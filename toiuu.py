import customtkinter as ctk
from tkinter import filedialog, messagebox
import sys
import os
import random
import time
import subprocess
import requests
import re
import concurrent.futures
import base64
import queue
import threading
from datetime import datetime, timedelta
import openpyxl

# Fix Unicode console output
sys.stdout.reconfigure(encoding='utf-8')

FIREFOX_API_URL = "http://www.firefox.fun/yhapi.ashx"

excel_lock = threading.Lock()

def log_to_excel(phone, otp_received, trigger_success, backup_success):
    excel_path = os.path.join(os.path.dirname(__file__), "campaign_stats.xlsx")
    with excel_lock:
        try:
            if not os.path.exists(excel_path):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Campaign Stats"
                ws.append(["Time", "Phone Number", "OTP Received", "Trigger Success", "Backup Success"])
            else:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
            
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws.append([now_str, str(phone), "Yes" if otp_received else "No", "Yes" if trigger_success else "No", "Yes" if backup_success else "No"])
            wb.save(excel_path)
        except Exception as e:
            print(f"Error saving to excel: {e}")

def translate_firefox_error(res, act=""):
    err_str = str(res).strip()
    if err_str.startswith("0|"):
        err_str = err_str.split("|", 1)[1]
    
    # Xử lý trường hợp trả về số giây chờ khi huỷ số (act=setRel)
    if err_str.isdigit() and act == "setRel":
        return f"Vui lòng đợi {err_str} giây nữa để huỷ số"
        
    mapping = {
        "login": {
            "-1": "Username cannot be empty",
            "-2": "Username length must be 3-30 characters",
            "-3": "Username cannot contain '|'",
            "-4": "Username cannot contain Chinese",
            "-5": "Password cannot be empty",
            "-6": "Password length must be 3-30 characters",
            "-7": "Request too fast, try again in 1 minute",
            "-8": "Account disabled",
            "-9": "Invalid username or password",
        },
        "myInfo": {
            "-1": "Token does not exist",
            "-2": "Invalid token, please re-login",
            "-3": "Please wait 60s before retrying",
        },
        "getPhone": {
            "-1": "No matching number found",
            "-2": "Token does not exist",
            "-3": "Service ID does not exist",
            "-4": "Invalid country",
            "-5": "Service not approved yet",
            "-6": "Service disabled",
            "-7": "Token disabled",
            "-8": "Insufficient balance, please top up",
            "-9": "Too many numbers held, please top up or release old ones",
            "-10": "Service does not allow specifying phone number",
        },
        "getPhoneCode": {
            "-1": "Token does not exist",
            "-2": "Invalid Pkey",
            "-3": "Waiting for verification code (no SMS yet)",
            "-4": "Phone number unavailable, please skip this one",
            "-5": "Phone number blacklisted, please skip",
        },
        "sendCode": {
            "-1": "Token does not exist",
            "-2": "Invalid Pkey",
            "-3": "Recipient number cannot be empty",
            "-4": "SMS content cannot be empty",
            "-5": "Unknown error",
            "-6": "Phone number unavailable, please skip",
            "-7": "Phone number unavailable, please skip",
            "-8": "Service does not allow sending SMS",
            "-9": "No keyword, contact support",
            "-10": "Invalid SMS content rule",
            "-11": "Do not send repeated requests",
        },
        "setRel": {
            "-1": "Token does not exist",
            "-2": "Invalid Pkey",
            "-3": "Phone number unavailable, please skip",
            "-4": "SMS received, cannot release number",
            "-5": "SMS is sending, cannot release",
            "-6": "Canceled too many times, account auto blacklisted",
        },
        "addBlack": {
            "-1": "Token does not exist",
            "-2": "Invalid Pkey",
            "-3": "Reason cannot be empty",
            "-4": "Phone number unavailable, please skip",
            "-5": "SMS not received, please call Release number API first",
            "-6": "No permission for Black number",
        }
    }
    
    # Generic fallback cho các trường hợp không khớp
    generic_mapping = {
        "-1": "Token does not exist",
        "-2": "Invalid Pkey/Token",
        "-3": "Parameter error / Waiting",
        "-4": "Phone number unavailable",
        "-5": "System error or blocked",
    }
    
    if act in mapping and err_str in mapping[act]:
        return f"{err_str} ({mapping[act][err_str]})"
        
    if err_str in generic_mapping:
        return f"{err_str} ({generic_mapping[err_str]})"
            
    if "ERROR|" in str(res):
        return str(res).replace("ERROR|", "NETWORK ERROR: ")
        
    return str(res)

def firefox_api(session, token, **params):
    params["token"] = token
    try:
        res = session.get(FIREFOX_API_URL, params=params, timeout=15)
        res.raise_for_status()
        res.encoding = 'utf-8' # Force utf-8 for Chinese characters
        return res.text.strip()
    except Exception as e:
        return f"ERROR|Network exception: {e}"

def firefox_get_phone(session, token, service_id, country="vn", mobile=None):
    params = {"act": "getPhone", "iid": service_id, "country": country}
    if mobile:
        params["mobile"] = mobile
    res = firefox_api(session, token, **params)
    parts = res.split("|")
    if parts[0] == "1":
        pkey = parts[1]
        phone = parts[7] if len(parts) > 7 and parts[7] else parts[4]
        return pkey, normalize_phone_number(phone)
    return None, translate_firefox_error(res, act="getPhone")

def firefox_set_rel(session, token, pkey):
    return firefox_api(session, token, act="setRel", pkey=pkey)

def firefox_add_black(session, token, pkey, reason="error"):
    return firefox_api(session, token, act="addBlack", pkey=pkey, reason=reason)

def firefox_api_return(session, token, pkey, remark="0"):
    return firefox_api(session, token, act="apiReturn", pkey=pkey, remark=remark)

def firefox_send_sms(session, token, pkey, receiver, content):
    return firefox_api(
        session,
        token,
        act="sendCode",
        pkey=pkey,
        receiver=receiver,
        smscontent=content
    )

def firefox_extract_otp(code="", sms_text=""):
    match = re.search(r"\b\d{4,8}\b", f"{code or ''} {sms_text or ''}")
    return match.group(0) if match else ""

def firefox_wait_sms_receipt(app, device_id, session, token, pkey, timeout=60):
    start = time.time()
    while time.time() - start < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return False
        res = firefox_api(session, token, act="getPhoneCode", pkey=pkey)
        if res.startswith("ERROR|"):
            time.sleep(3)
            continue
            
        parts = res.split("|")
        if parts[0] == "1" and len(parts) > 1:
            code = parts[1].strip()
            sms_text = "|".join(parts[2:]).strip()
            
            is_receipt = "发送成功" in code or "发送成功" in sms_text or "短信发送成功" in sms_text
            is_receipt_garbled = "å" in code or "é" in sms_text
            if is_receipt or is_receipt_garbled:
                app.log(f"[{device_id}] Đã nhận được biên lai gửi SMS thành công từ Firefox.")
                return True
        elif parts[0] == "0" and len(parts) > 1:
            err_code = parts[1]
            if err_code != "-3":
                app.log(f"[{device_id}] Lỗi khi chờ biên lai SMS Firefox: {translate_firefox_error(res, act='getPhoneCode')}")
                return False
        time.sleep(3)
    app.log(f"[{device_id}] Timeout chờ biên lai SMS Firefox.")
    return False

def firefox_wait_otp(app, device_id, session, token, pkey, timeout=120):
    start = time.time()
    while time.time() - start < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return None
        res = firefox_api(session, token, act="getPhoneCode", pkey=pkey)
        if res.startswith("ERROR|"):
            app.log(f"[{device_id}] Firefox API Error: {res}")
            time.sleep(3)
            continue
            
        parts = res.split("|")
        if parts[0] == "1" and len(parts) > 1:
            code = parts[1].strip()
            sms_text = "|".join(parts[2:]).strip()
            
            otp = firefox_extract_otp(code, sms_text)
            if otp:
                app.log(f"[{device_id}] Got Firefox OTP: {otp}")
                return otp
        elif parts[0] == "0" and len(parts) > 1:
            err_code = parts[1]
            if err_code == "-3":
                pass # Still waiting
            else:
                app.log(f"[{device_id}] Firefox OTP error: {translate_firefox_error(res, act='getPhoneCode')}")
                return None
        time.sleep(3)
        
    app.log(f"[{device_id}] Firefox OTP Timeout sau {timeout}s.")
    return None
    return None


def normalize_phone_number(phone):
    p = str(phone).strip().replace(" ", "").replace("\r", "").replace("\n", "")
    if p.startswith("+84"):
        p = "0" + p[3:]
    elif p.startswith("84") and len(p) >= 11:
        p = "0" + p[2:]
    elif len(p) == 9 and p[0] in "35789":
        p = "0" + p
    if not p.startswith("0") and len(p) > 0:
        p = "0" + p
    return p

def firefox_mobile_param(phone):
    p = str(phone).strip().replace(" ", "")
    if p.startswith("+84"):
        return p[3:]
    if p.startswith("84") and len(p) >= 11:
        return p[2:]
    if p.startswith("0") and len(p) == 10:
        return p[1:]
    return p

# ADB Configuration Constants
ADB_COMMAND_TIMEOUT = 10
UI_DUMP_TIMEOUT = 8
POLL_INTERVAL = 0.8
MAX_RETRIES = 2

# CustomTkinter setup
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# =====================================================================
# ===================== DESIGN SYSTEM CONSTANTS =======================
# =====================================================================

COLORS = {
    "bg_deep": "#080c14",
    "bg_panel": "#0f1520",
    "bg_card": "#151d2e",
    "bg_input": "#0a0f1a",
    "bg_header": "#0b1120",
    "border": "#1a2336",
    "border_light": "#243049",
    "border_accent": "#2d4a7a",
    "text_primary": "#f0f4ff",
    "text_secondary": "#94a3b8",
    "text_muted": "#475569",
    "accent_blue": "#3b82f6",
    "accent_cyan": "#22d3ee",
    "accent_green": "#10b981",
    "accent_red": "#ef4444",
    "accent_amber": "#f59e0b",
    "accent_purple": "#8b5cf6",
    "accent_pink": "#ec4899",
    "accent_emerald": "#34d399",
}

LOG_LEVELS = {
    "INFO":    {"color": "#38bdf8", "icon": "ℹ"},
    "SUCCESS": {"color": "#10b981", "icon": "✓"},
    "WARN":    {"color": "#f59e0b", "icon": "⚠"},
    "ERROR":   {"color": "#ef4444", "icon": "✗"},
    "DEBUG":   {"color": "#8b5cf6", "icon": "◆"},
    "STEP":    {"color": "#94a3b8", "icon": "▸"},
    "SYSTEM":  {"color": "#e0e7ff", "icon": "⚙"},
}


class ZaloAutoUIApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("⚡ ToolZL Auto UI — Master Controller v2.0")
        self.geometry("1100x820")
        self.minsize(1000, 720)
        self.configure(fg_color=COLORS["bg_deep"])

        # Application state
        self.is_running = False
        self.devices = []
        self.running_devices = []
        self.phone_numbers = []
        self.txt_path = ""
        self.device_ui_elements = {}  # device_id -> {frame, phone_lbl, status_lbl, ...}
        self.active_running_devices = set()  # Currently running device IDs
        self.active_workers_count = 0
        self.active_workers_lock = threading.Lock()
        self.total_phones = 0
        self.success_phones = 0
        self.failed_phones = 0
        self.otp_received_count = 0
        self.trigger_success_count = 0
        self.backup_success_count = 0
        self.completed_lock = threading.Lock()

        # Per-device locks for thread-safe ADB operations
        self.device_locks = {}

        # Retry tracking
        self.retry_counts = {}
        self.retry_lock = threading.Lock()

        # Enhanced UI state
        self.campaign_start_time = None
        self.log_entries = []       # Store all log entries for filtering
        self.current_log_filter = "ALL"
        self.auto_scroll_var = ctk.BooleanVar(value=True)

        # Grid layout — 5 rows
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=0)  # Header
        self.grid_rowconfigure(1, weight=0)  # Progress bar
        self.grid_rowconfigure(2, weight=1)  # Main panels
        self.grid_rowconfigure(3, weight=0)  # Stats (right) / left continues
        self.grid_rowconfigure(4, weight=3, minsize=250)  # Logs

        self.create_widgets()
        self.load_settings()
        self._start_live_clock()
        self.refresh_devices_list(on_startup=True)
        self.protocol("WM_DELETE_WINDOW", self.on_close_app)

    # =================================================================
    # ======================== CREATE WIDGETS =========================
    # =================================================================
    def create_widgets(self):

        # ==================== HEADER BAR ====================
        header = ctk.CTkFrame(self, height=60, fg_color=COLORS["bg_header"], corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        # Branding (left)
        brand = ctk.CTkFrame(header, fg_color="transparent")
        brand.grid(row=0, column=0, padx=15, pady=8, sticky="w")

        ctk.CTkLabel(brand, text="⚡", font=("Segoe UI", 26)).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(brand, text="ToolZL Master", font=("Segoe UI", 20, "bold"),
                     text_color=COLORS["text_primary"]).pack(side="left")

        version_lbl = ctk.CTkLabel(brand, text=" v2.0 ", font=("Segoe UI", 15, "bold"),
                                   fg_color=COLORS["accent_blue"], corner_radius=6,
                                   text_color="#ffffff", width=46, height=24)
        version_lbl.pack(side="left", padx=(8, 0))

        # Right side — uptime + clock
        clock_box = ctk.CTkFrame(header, fg_color="transparent")
        clock_box.grid(row=0, column=1, padx=15, pady=8, sticky="e")

        self.uptime_label = ctk.CTkLabel(clock_box, text="",
                                         font=("Consolas", 14), text_color=COLORS["accent_green"])
        self.uptime_label.pack(side="left", padx=(0, 15))

        self.clock_label = ctk.CTkLabel(clock_box, text="00:00:00",
                                        font=("Consolas", 17, "bold"), text_color=COLORS["text_muted"])
        self.clock_label.pack(side="left")

        # ==================== CAMPAIGN PROGRESS BAR ====================
        pbar_frame = ctk.CTkFrame(self, height=38, fg_color=COLORS["bg_header"], corner_radius=0)
        pbar_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        pbar_frame.grid_propagate(False)
        pbar_frame.grid_columnconfigure(1, weight=1)

        self.progress_label = ctk.CTkLabel(pbar_frame, text="Campaign Idle",
                                           font=("Segoe UI", 16), text_color=COLORS["text_muted"])
        self.progress_label.grid(row=0, column=0, padx=15, pady=4, sticky="w")

        self.campaign_progress_bar = ctk.CTkProgressBar(pbar_frame, height=10,
                                                        progress_color=COLORS["accent_green"],
                                                        fg_color=COLORS["border"], corner_radius=4)
        self.campaign_progress_bar.grid(row=0, column=1, padx=(0, 8), pady=4, sticky="ew")
        self.campaign_progress_bar.set(0)

        self.progress_pct_label = ctk.CTkLabel(pbar_frame, text="0%",
                                               font=("Consolas", 13, "bold"),
                                               text_color=COLORS["accent_cyan"])
        self.progress_pct_label.grid(row=0, column=2, padx=(0, 15), pady=4)

        # ==================== LEFT COLUMN — SETTINGS ====================
        left_panel = ctk.CTkScrollableFrame(self, corner_radius=12, fg_color=COLORS["bg_panel"],
                                            border_width=1, border_color=COLORS["border"])
        left_panel.grid(row=2, column=0, padx=(12, 6), pady=(6, 10), sticky="nsew")
        left_panel.grid_columnconfigure(0, weight=1)

        # Section: CONFIGURATION
        sec_hdr = ctk.CTkFrame(left_panel, fg_color=COLORS["bg_card"], corner_radius=8, height=42)
        sec_hdr.pack(fill="x", padx=10, pady=(10, 6))
        sec_hdr.pack_propagate(False)
        ctk.CTkLabel(sec_hdr, text="⚙️  CONFIGURATION", font=("Segoe UI", 15, "bold"),
                     text_color=COLORS["accent_cyan"]).pack(side="left", padx=12, pady=5)

        # ADB Path
        adb_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        adb_frame.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(adb_frame, text="ADB Executable", font=("Segoe UI", 13, "bold"),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 2))

        adb_row = ctk.CTkFrame(adb_frame, fg_color="transparent")
        adb_row.pack(fill="x")
        self.adb_path_entry = ctk.CTkEntry(
            adb_row, placeholder_text="Path to adb.exe", height=38,
            fg_color=COLORS["bg_input"], border_color=COLORS["border"], border_width=1,
            text_color=COLORS["text_primary"], placeholder_text_color=COLORS["text_muted"],
            font=("Consolas", 13))
        self.adb_path_entry.insert(0, r"D:\SIM\pmcong\platform-tools-latest-windows\platform-tools\adb.exe")
        self.adb_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))

        ctk.CTkButton(adb_row, text="📂", width=40, height=38,
                      fg_color=COLORS["bg_card"], border_width=1, border_color=COLORS["border_light"],
                      hover_color="#243049", text_color=COLORS["text_secondary"],
                      font=("Segoe UI", 20), command=self.browse_adb).pack(side="right")

        # Captcha Offset
        off_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        off_frame.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(off_frame, text="Captcha Offset (px)", font=("Segoe UI", 13, "bold"),
                     text_color=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 2))
        self.captcha_offset_entry = ctk.CTkEntry(
            off_frame, placeholder_text="0", height=38,
            fg_color=COLORS["bg_input"], border_color=COLORS["border"], border_width=1,
            text_color=COLORS["text_primary"], font=("Consolas", 13))
        self.captcha_offset_entry.insert(0, "0")
        self.captcha_offset_entry.pack(fill="x")

        # Firefox Config
        ff_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        ff_frame.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(ff_frame, text="Firefox Token | Service ID | Country", font=("Segoe UI", 13, "bold"), text_color=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 2))
        ff_input_row = ctk.CTkFrame(ff_frame, fg_color="transparent")
        ff_input_row.pack(fill="x")
        self.ff_token_entry = ctk.CTkEntry(ff_input_row, placeholder_text="Token", height=38, fg_color=COLORS["bg_input"], border_color=COLORS["border"], border_width=1, text_color=COLORS["text_primary"])
        self.ff_token_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.ff_token_entry.insert(0, "aba3dcd7dde85c51ec3454c56a85f77d_304014")
        self.ff_service_entry = ctk.CTkEntry(ff_input_row, placeholder_text="ID", width=60, height=38, fg_color=COLORS["bg_input"], border_color=COLORS["border"], border_width=1, text_color=COLORS["text_primary"])
        self.ff_service_entry.insert(0, "1049")
        self.ff_service_entry.pack(side="left", padx=(0, 4))
        self.ff_country_entry = ctk.CTkEntry(ff_input_row, placeholder_text="VN", width=50, height=38, fg_color=COLORS["bg_input"], border_color=COLORS["border"], border_width=1, text_color=COLORS["text_primary"])
        self.ff_country_entry.insert(0, "vnm")
        self.ff_country_entry.pack(side="left")

        # Balance Check
        self.ff_balance_frame = ctk.CTkFrame(ff_frame, fg_color="transparent")
        self.ff_balance_frame.pack(fill="x", pady=(8, 0))
        self.btn_check_balance = ctk.CTkButton(self.ff_balance_frame, text="🔍 Check Balance", 
                                               font=("Segoe UI", 12, "bold"), height=30, width=120,
                                               fg_color=COLORS["accent_blue"], hover_color="#1d4ed8",
                                               command=self.check_firefox_balance)
        self.btn_check_balance.pack(side="left")
        self.lbl_balance = ctk.CTkLabel(self.ff_balance_frame, text="Balance: --", font=("Segoe UI", 13, "bold"), text_color=COLORS["accent_emerald"])
        self.lbl_balance.pack(side="left", padx=(10, 0))

        # Separator
        ctk.CTkFrame(left_panel, height=1, fg_color=COLORS["border"]).pack(fill="x", padx=12, pady=10)

        # Section: PHONE DATA
        ph_hdr = ctk.CTkFrame(left_panel, fg_color=COLORS["bg_card"], corner_radius=8, height=42)
        ph_hdr.pack(fill="x", padx=10, pady=(0, 6))
        ph_hdr.pack_propagate(False)
        ctk.CTkLabel(ph_hdr, text="📞  PHONE DATA INPUT", font=("Segoe UI", 15, "bold"),
                     text_color=COLORS["accent_cyan"]).pack(side="left", padx=12, pady=5)

        # Firefox Qty
        self.firefox_qty_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        ctk.CTkLabel(self.firefox_qty_frame, text="Amount (Firefox Auto):", font=("Segoe UI", 13, "bold"), text_color=COLORS["text_secondary"]).pack(anchor="w", pady=(0, 2))

        qty_input_row = ctk.CTkFrame(self.firefox_qty_frame, fg_color="transparent")
        qty_input_row.pack(fill="x")

        self.firefox_qty_entry = ctk.CTkEntry(qty_input_row, placeholder_text="Example: 10", height=38, fg_color=COLORS["bg_input"], border_color=COLORS["border"], border_width=1, text_color=COLORS["text_primary"], font=("Consolas", 13))
        self.firefox_qty_entry.insert(0, "10")
        self.firefox_qty_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_update_qty = ctk.CTkButton(qty_input_row, text="Update", width=60, height=38, fg_color=COLORS["accent_blue"], command=self.update_qty_live)
        self.btn_update_qty.pack(side="right")
        
        self.firefox_qty_frame.pack(fill="x", padx=12, pady=10)

        # ==================== RIGHT COLUMN — DASHBOARD ====================
        right_panel = ctk.CTkFrame(self, corner_radius=12, fg_color=COLORS["bg_panel"],
                                   border_width=1, border_color=COLORS["border"])
        right_panel.grid(row=2, column=1, padx=(6, 12), pady=(6, 4), sticky="nsew")
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(1, weight=1)

        # Dashboard header
        dh = ctk.CTkFrame(right_panel, fg_color=COLORS["bg_card"], corner_radius=8, height=42)
        dh.grid(row=0, column=0, padx=10, pady=(10, 6), sticky="ew")
        dh.grid_propagate(False)
        dh.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(dh, text="🖥️  AUTOMATION DASHBOARD", font=("Segoe UI", 15, "bold"),
                     text_color=COLORS["accent_emerald"]).grid(row=0, column=0, padx=12, pady=5, sticky="w")

        # Scrollable device list
        self.devices_scroll_frame = ctk.CTkScrollableFrame(
            right_panel, fg_color=COLORS["bg_deep"], border_width=1,
            border_color=COLORS["border"], corner_radius=8,
            label_text="Connected Devices", label_text_color=COLORS["text_muted"],
            label_font=("Segoe UI", 16))
        self.devices_scroll_frame.grid(row=1, column=0, padx=10, pady=(0, 6), sticky="nsew")
        self.devices_scroll_frame.grid_columnconfigure(0, weight=1)

        # Control buttons
        ctrl = ctk.CTkFrame(right_panel, fg_color="transparent")
        ctrl.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")
        ctrl.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_refresh = ctk.CTkButton(ctrl, text="🔄 Refresh", font=("Segoe UI", 14, "bold"),
                                         fg_color=COLORS["bg_card"], border_width=1,
                                         border_color=COLORS["border_light"], hover_color="#243049",
                                         text_color=COLORS["text_secondary"], height=42,
                                         command=self.refresh_devices_list)
        self.btn_refresh.grid(row=0, column=0, padx=3, sticky="ew")

        self.btn_start = ctk.CTkButton(ctrl, text="🚀 START", font=("Segoe UI", 15, "bold"),
                                       fg_color=COLORS["accent_green"], hover_color="#059669",
                                       text_color="#ffffff", height=42, command=self.start_campaign)
        self.btn_start.grid(row=0, column=1, padx=3, sticky="ew")

        self.btn_stop = ctk.CTkButton(ctrl, text="🛑 STOP", font=("Segoe UI", 15, "bold"),
                                      fg_color=COLORS["accent_red"], hover_color="#dc2626",
                                      text_color="#ffffff", height=42, command=self.stop_campaign,
                                      state="disabled")
        self.btn_stop.grid(row=0, column=2, padx=3, sticky="ew")

        # ==================== STATS PANEL ====================
        self.stats_frame = ctk.CTkFrame(left_panel, fg_color=COLORS["bg_panel"], corner_radius=10,
                                        border_width=1, border_color=COLORS["border"])
        self.stats_frame.pack(fill="x", padx=12, pady=10, side="bottom")
        self.stats_frame.grid_columnconfigure((0, 1), weight=1)

        self.lbl_total_phones = ctk.CTkLabel(self.stats_frame, text="📊 Total: 0",
                                             font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_blue"])
        self.lbl_total_phones.grid(row=0, column=0, pady=(12, 4), padx=15, sticky="w")

        self.lbl_remaining_phones = ctk.CTkLabel(self.stats_frame, text="⏳ Remaining: 0",
                                                 font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_amber"])
        self.lbl_remaining_phones.grid(row=0, column=1, pady=(12, 4), padx=15, sticky="w")

        self.lbl_success_phones = ctk.CTkLabel(self.stats_frame, text="✅ Success: 0",
                                               font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_green"])
        self.lbl_success_phones.grid(row=1, column=0, pady=4, padx=15, sticky="w")

        self.lbl_otp_received = ctk.CTkLabel(self.stats_frame, text="✉️ OTP: 0",
                                             font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_cyan"])
        self.lbl_otp_received.grid(row=1, column=1, pady=4, padx=15, sticky="w")

        self.lbl_failed_phones = ctk.CTkLabel(self.stats_frame, text="❌ Failed: 0",
                                              font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_red"])
        self.lbl_failed_phones.grid(row=2, column=0, pady=4, padx=15, sticky="w")

        self.lbl_active_threads = ctk.CTkLabel(self.stats_frame, text="⚙️ Threads: 0",
                                               font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_blue"])
        self.lbl_active_threads.grid(row=2, column=1, pady=4, padx=15, sticky="w")

        self.lbl_trigger_success = ctk.CTkLabel(self.stats_frame, text="🎯 Trigger: 0",
                                                font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_amber"])
        self.lbl_trigger_success.grid(row=3, column=0, pady=4, padx=15, sticky="w")

        self.lbl_backup_success = ctk.CTkLabel(self.stats_frame, text="💾 Backup: 0",
                                               font=("Segoe UI", 14, "bold"), text_color=COLORS["accent_green"])
        self.lbl_backup_success.grid(row=3, column=1, pady=4, padx=15, sticky="w")

        self.lbl_elapsed = ctk.CTkLabel(self.stats_frame, text="⏱ 00:00",
                                        font=("Consolas", 14, "bold"), text_color=COLORS["text_muted"])
        self.lbl_elapsed.grid(row=4, column=0, pady=(4, 12), padx=15, sticky="w")



        # ==================== LOG PANEL ====================
        log_panel = ctk.CTkFrame(self, corner_radius=12, fg_color=COLORS["bg_panel"],
                                 border_width=1, border_color=COLORS["border"])
        log_panel.grid(row=4, column=0, columnspan=2, padx=12, pady=(4, 10), sticky="nsew")
        log_panel.grid_columnconfigure(0, weight=1)
        log_panel.grid_rowconfigure(1, weight=1)

        # Log toolbar
        log_bar = ctk.CTkFrame(log_panel, fg_color="transparent", height=40)
        log_bar.grid(row=0, column=0, padx=10, pady=(8, 3), sticky="ew")
        log_bar.grid_columnconfigure(6, weight=1)

        ctk.CTkLabel(log_bar, text="📋 CONSOLE", font=("Segoe UI", 14, "bold"),
                     text_color=COLORS["text_secondary"]).grid(row=0, column=0, padx=(4, 10), sticky="w")

        # Filter buttons
        filter_defs = [
            ("All",     "ALL",     COLORS["text_secondary"]),
            ("Errors",  "ERROR",   COLORS["accent_red"]),
            ("Success", "SUCCESS", COLORS["accent_green"]),
            ("Warn",    "WARN",    COLORS["accent_amber"]),
        ]
        self.filter_buttons = {}
        for idx, (label, ftype, color) in enumerate(filter_defs):
            btn = ctk.CTkButton(
                log_bar, text=label, width=70, height=28,
                fg_color=COLORS["border_light"] if ftype == "ALL" else COLORS["bg_card"],
                hover_color=COLORS["border_light"], text_color=color,
                font=("Segoe UI", 12, "bold"), corner_radius=6,
                command=lambda ft=ftype: self.filter_logs(ft))
            btn.grid(row=0, column=idx + 1, padx=2, sticky="w")
            self.filter_buttons[ftype] = btn

        # Right toolbar items
        rtb = ctk.CTkFrame(log_bar, fg_color="transparent")
        rtb.grid(row=0, column=6, sticky="e")

        self.auto_scroll_cb = ctk.CTkCheckBox(
            rtb, text="Auto-scroll", variable=self.auto_scroll_var,
            fg_color=COLORS["accent_blue"], hover_color="#1d4ed8",
            text_color=COLORS["text_muted"], font=("Segoe UI", 15),
            width=18, height=18, checkbox_width=15, checkbox_height=15)
        self.auto_scroll_cb.pack(side="left", padx=(0, 8))

        ctk.CTkButton(rtb, text="📥 Export", width=62, height=22,
                      fg_color=COLORS["bg_card"], hover_color=COLORS["border_light"],
                      text_color=COLORS["text_muted"], font=("Segoe UI", 15, "bold"),
                      corner_radius=5, command=self.export_logs).pack(side="left", padx=(0, 4))

        ctk.CTkButton(rtb, text="🗑 Clear", width=58, height=22,
                      fg_color=COLORS["bg_card"], hover_color=COLORS["border_light"],
                      text_color=COLORS["text_muted"], font=("Segoe UI", 15, "bold"),
                      corner_radius=5, command=self.clear_logs).pack(side="left")

        # Log text container
        self.log_container = ctk.CTkFrame(log_panel, fg_color="transparent")
        self.log_container.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.log_container.grid_rowconfigure(0, weight=1)

        self.device_log_terminals = {}
        self.system_log_terminal = None

        self._rebuild_log_terminals([])

        # Initial log message
        self.log("System console initialized. Ready to operate.", level="SYSTEM")

        # Show correct phone input mode
        

    # =================================================================
    # ==================== LOGGING SYSTEM ============================
    # =================================================================

    def _rebuild_log_terminals(self, devices):
        if not hasattr(self, 'log_container'): return
        
        for child in self.log_container.winfo_children():
            child.destroy()
            
        self.device_log_terminals.clear()
        
        cols = len(devices) + 1
        for i in range(cols):
            self.log_container.grid_columnconfigure(i, weight=1)
            
        # Create SYSTEM log
        sys_frame = ctk.CTkFrame(self.log_container, fg_color="transparent")
        sys_frame.grid(row=0, column=0, padx=2, pady=2, sticky="nsew")
        sys_frame.grid_rowconfigure(1, weight=1)
        sys_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(sys_frame, text="⚙️ SYSTEM", font=("Segoe UI", 12, "bold"), text_color=COLORS["accent_blue"]).grid(row=0, column=0, sticky="w")
        self.system_log_terminal = ctk.CTkTextbox(sys_frame, fg_color=COLORS["bg_deep"],
                                           text_color=COLORS["accent_blue"], border_width=1,
                                           border_color=COLORS["border"], font=("Consolas", 12),
                                           corner_radius=8)
        self.system_log_terminal.grid(row=1, column=0, sticky="nsew")
        self.system_log_terminal.configure(state="disabled")
        self._setup_log_tags(self.system_log_terminal)
        
        # Create device logs
        for idx, dev_id in enumerate(devices):
            dev_frame = ctk.CTkFrame(self.log_container, fg_color="transparent")
            dev_frame.grid(row=0, column=idx+1, padx=2, pady=2, sticky="nsew")
            dev_frame.grid_rowconfigure(1, weight=1)
            dev_frame.grid_columnconfigure(0, weight=1)
            
            ctk.CTkLabel(dev_frame, text=f"📱 {dev_id}", font=("Segoe UI", 12, "bold"), text_color=COLORS["text_secondary"]).grid(row=0, column=0, sticky="w")
            dev_term = ctk.CTkTextbox(dev_frame, fg_color=COLORS["bg_deep"],
                                               text_color=COLORS["accent_blue"], border_width=1,
                                               border_color=COLORS["border"], font=("Consolas", 12),
                                               corner_radius=8)
            dev_term.grid(row=1, column=0, sticky="nsew")
            dev_term.configure(state="disabled")
            self._setup_log_tags(dev_term)
            self.device_log_terminals[dev_id] = dev_term

        self.filter_logs(getattr(self, 'current_log_filter', 'ALL'))

    def _setup_log_tags(self, terminal=None):
        """Configure colored text tags on the underlying tkinter Text widget."""
        try:
            if not terminal: return
            tw = terminal._textbox
            tw.tag_configure("ts", foreground="#475569", font=("Consolas", 12))
            tw.tag_configure("device_id", foreground="#60a5fa", font=("Consolas", 12, "bold"))
            for lname, lcfg in LOG_LEVELS.items():
                tw.tag_configure(lname, foreground=lcfg["color"])
                tw.tag_configure(f"{lname}_b", foreground=lcfg["color"], font=("Consolas", 12, "bold"))
        except Exception:
            pass

    def log(self, text, level=None):
        """Thread-safe, color-coded, filterable log method."""
        if level is None:
            tl = text.lower()
            if any(k in text for k in ["❌", "CRASH"]) or any(k in tl for k in ["error", "failed", "crash", "timeout", "not found"]):
                level = "ERROR"
            elif any(k in text for k in ["✅", "🟢", "🎉"]) or any(k in tl for k in ["success", "completed"]):
                level = "SUCCESS"
            elif any(k in text for k in ["⚠️", "⚠"]) or "warning" in tl:
                level = "WARN"
            elif any(k in text for k in ["🔄"]) or any(k in tl for k in ["starting", "setting up", "launching", "waiting"]):
                level = "STEP"
            else:
                level = "INFO"

        ts = datetime.now().strftime('%H:%M:%S')

        if level == "ERROR":
            m = re.match(r'\[([^\]]+)\]\s*(.*)', text, re.DOTALL)
            if m:
                if not hasattr(self, 'device_last_error'):
                    self.device_last_error = {}
                self.device_last_error[m.group(1)] = m.group(2)
        self.log_entries.append({"time": ts, "level": level, "text": text})

        # Skip display if current filter doesn't match
        if self.current_log_filter != "ALL" and level != self.current_log_filter:
            return

        def _append():
            target_terminals = []
            dev_match = re.match(r'\[([^\]]+)\]\s*(.*)', text, re.DOTALL)
            
            if dev_match:
                dev_id = dev_match.group(1)
                msg_body = dev_match.group(2)
                if hasattr(self, 'device_log_terminals') and dev_id in self.device_log_terminals:
                    target_terminals.append((self.device_log_terminals[dev_id], msg_body))
                else:
                    target_terminals.append((getattr(self, 'system_log_terminal', None), text))
            else:
                target_terminals.append((getattr(self, 'system_log_terminal', None), text))

            for term, msg in target_terminals:
                if not term: continue
                try:
                    term.configure(state="normal")
                    tw = term._textbox
                    icon = LOG_LEVELS.get(level, {}).get("icon", "")
                    tw.insert("end", f"[{ts}]", ("ts",))
                    tw.insert("end", f" {icon} ", (f"{level}_b",))
                    tw.insert("end", f"{msg}\n", (level,))
                    
                    if self.auto_scroll_var.get():
                        term.see("end")
                    term.configure(state="disabled")
                except Exception:
                    pass
        self.after(0, _append)

    def filter_logs(self, filter_type):
        """Re-render log terminal showing only entries matching the filter."""
        self.current_log_filter = filter_type
        if hasattr(self, 'filter_buttons'):
            for ft, btn in self.filter_buttons.items():
                btn.configure(fg_color=COLORS["border_light"] if ft == filter_type else COLORS["bg_card"])

        def _rerender():
            terminals = [getattr(self, 'system_log_terminal', None)] + list(getattr(self, 'device_log_terminals', {}).values())
            for t in terminals:
                if t:
                    t.configure(state="normal")
                    try:
                        t._textbox.delete("1.0", "end")
                    except:
                        pass

            for e in self.log_entries:
                if filter_type != "ALL" and e["level"] != filter_type:
                    continue
                lv = e["level"]
                icon = LOG_LEVELS.get(lv, {}).get("icon", "")
                
                dev_m = re.match(r'\[([^\]]+)\]\s*(.*)', e["text"], re.DOTALL)
                target_term = getattr(self, 'system_log_terminal', None)
                msg = e["text"]
                if dev_m:
                    dev_id = dev_m.group(1)
                    msg = dev_m.group(2)
                    if hasattr(self, 'device_log_terminals') and dev_id in self.device_log_terminals:
                        target_term = self.device_log_terminals[dev_id]

                if not target_term: continue
                
                try:
                    target_term.insert("end", f"[{e['time']}]", ("ts",))
                    target_term.insert("end", f" {icon} ", (f"{lv}_b",))
                    target_term.insert("end", f"{msg}\n", (lv,))
                except:
                    pass
                
            for t in terminals:
                if t:
                    try:
                        t.see("end")
                        t.configure(state="disabled")
                    except:
                        pass
        self.after(0, _rerender)

    def export_logs(self):
        """Export all log entries to a text file."""
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"campaign_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    for e in self.log_entries:
                        f.write(f"[{e['time']}] [{e['level']}] {e['text']}\n")
                self.log(f"Log exported → {path}", level="SUCCESS")
            except Exception as ex:
                self.log(f"Export failed: {ex}", level="ERROR")

    def clear_logs(self):
        """Clear all log entries and terminal content."""
        self.log_entries.clear()
        def _clr():
            terminals = [getattr(self, 'system_log_terminal', None)] + list(getattr(self, 'device_log_terminals', {}).values())
            for t in terminals:
                if t:
                    t.configure(state="normal")
                    try:
                        t._textbox.delete("1.0", "end")
                    except:
                        pass
                    t.configure(state="disabled")
        self.after(0, _clr)
        self.log("Console cleared.", level="SYSTEM")

    # =================================================================
    # ==================== LIVE CLOCK & TIMERS =======================
    # =================================================================

    def _start_live_clock(self):
        self._tick_clock()

    def _tick_clock(self):
        now = datetime.now()
        self.clock_label.configure(text=now.strftime("%H:%M:%S"))

        # Campaign uptime
        if self.is_running and self.campaign_start_time:
            elapsed = now - self.campaign_start_time
            secs = int(elapsed.total_seconds())
            h, rem = divmod(secs, 3600)
            m, s = divmod(rem, 60)
            ts_str = f"{h:02d}:{m:02d}:{s:02d}"
            self.uptime_label.configure(text=f"▶ {ts_str}")
            self.lbl_elapsed.configure(text=f"⏱ {ts_str}")

            # Speed update removed
        else:
            self.uptime_label.configure(text="")

        # Per-device elapsed times
        if self.is_running:
            ct = time.time()
            for dev_id, el in self.device_ui_elements.items():
                if 'task_start_time' in el and dev_id in self.active_running_devices and 'elapsed_lbl' in el:
                    d = int(ct - el['task_start_time'])
                    dm, ds = divmod(d, 60)
                    el['elapsed_lbl'].configure(text=f"⏱ {dm:02d}:{ds:02d}")

        self.after(1000, self._tick_clock)

    # =================================================================
    # ==================== STATS / PROGRESS ==========================
    # =================================================================

    def update_stats_ui(self):
        def _update():
            remaining = max(0, self.total_phones - self.success_phones - self.failed_phones)
            processed = self.success_phones + self.failed_phones
            active_threads = len(self.active_running_devices)

            self.lbl_total_phones.configure(text=f"📊 Total: {self.total_phones}")
            self.lbl_success_phones.configure(text=f"✅ Success: {self.success_phones}")
            self.lbl_failed_phones.configure(text=f"❌ Failed: {self.failed_phones}")
            self.lbl_remaining_phones.configure(text=f"⏳ Remaining: {remaining}")
            self.lbl_otp_received.configure(text=f"✉️ OTP: {self.otp_received_count}")
            self.lbl_active_threads.configure(text=f"⚙️ Threads: {active_threads}")
            self.lbl_trigger_success.configure(text=f"🎯 Trigger: {self.trigger_success_count}")
            self.lbl_backup_success.configure(text=f"💾 Backup: {self.backup_success_count}")

            if self.total_phones > 0:
                pct = processed / self.total_phones
                self.campaign_progress_bar.set(pct)
                self.progress_pct_label.configure(text=f"{int(pct * 100)}%")
                self.progress_label.configure(
                    text=f"Campaign — {processed}/{self.total_phones} processed",
                    text_color=COLORS["accent_green"] if pct >= 1.0 else COLORS["text_secondary"])
                # Dynamic color
                if pct >= 1.0:
                    self.campaign_progress_bar.configure(progress_color=COLORS["accent_green"])
                elif pct >= 0.5:
                    self.campaign_progress_bar.configure(progress_color=COLORS["accent_cyan"])
                else:
                    self.campaign_progress_bar.configure(progress_color=COLORS["accent_blue"])
        self.after(0, _update)

    # =================================================================
    # ==================== DEVICE UI UPDATES =========================
    # =================================================================

    def _infer_pipeline_progress(self, status_text):
        """Map a device status string to a pipeline progress 0.0–1.0."""
        t = status_text.lower()
        mapping = [
            (["idle", "ready"], 0.0),
            (["reverse proxy", "setting up"], 0.05),
            (["xtoolz"], 0.08),
            (["rebooting", "device id"], 0.12),
            (["booting", "wait adb"], 0.20),
            (["4g", "wifi", "verifying"], 0.30),
            (["launching zalo"], 0.40),
            (["entering phone"], 0.48),
            (["ticking"], 0.52),
            (["resolving"], 0.55),
            (["captcha"], 0.60),
            (["sms", "otp"], 0.68),
            (["submitting otp"], 0.72),
            (["profile"], 0.78),
            (["birthday", "sex"], 0.82),
            (["7up", "mini app", "opening"], 0.88),
            (["token", "saving"], 0.93),
            (["backup"], 0.96),
            (["done", "completed", "✅"], 1.0),
        ]
        for kws, pv in mapping:
            if any(k in t for k in kws):
                return pv
        if any(k in t for k in ["failed", "crash", "terminated", "stopped"]):
            return 0.0
        return 0.0

    def update_device_ui(self, device_id, phone=None, status_text=None, text_color="#e2e8f0"):
        def _update():
            if device_id in self.device_ui_elements:
                el = self.device_ui_elements[device_id]
                if phone is not None:
                    el['phone_lbl'].configure(text=f"📞 {phone}")
                    el['task_start_time'] = time.time()
                if status_text is not None:
                    el['status_lbl'].configure(text=status_text, text_color=text_color)
                    pv = self._infer_pipeline_progress(status_text)
                    if 'mini_progress' in el:
                        el['mini_progress'].set(pv)
                        if pv >= 0.9:
                            el['mini_progress'].configure(progress_color=COLORS["accent_green"])
                        elif pv >= 0.5:
                            el['mini_progress'].configure(progress_color=COLORS["accent_cyan"])
                        else:
                            el['mini_progress'].configure(progress_color=COLORS["accent_blue"])
                    if 'step_lbl' in el:
                        clean = status_text
                        for prefix in ["🔄 ", "✅ ", "❌ ", "🧩 ", "✉️ ", "🎮 ", "💤 ", "🔴 ", "🛑 ", "● "]:
                            clean = clean.replace(prefix, "")
                        el['step_lbl'].configure(text=f"{int(pv*100)}% — {clean}")
        self.after(0, _update)

    # =================================================================
    # ==================== DEVICE LIST ===============================
    # =================================================================

    def refresh_devices_list(self, on_startup=False):
        adb_exec = self.adb_path_entry.get().strip()
        if not os.path.exists(adb_exec):
            for child in self.devices_scroll_frame.winfo_children():
                child.destroy()
            self.device_ui_elements.clear()
            ctk.CTkLabel(self.devices_scroll_frame,
                         text="⚠️ ADB not found. Set the correct path and click Refresh.",
                         text_color=COLORS["accent_red"], font=("Segoe UI", 14, "italic")).pack(pady=20)
            if not on_startup:
                self.log(f"ADB Not Found at: '{adb_exec}'", level="ERROR")
            return

        new_scan_devices = []
        try:
            result = subprocess.run(f'"{adb_exec}" devices', shell=True, capture_output=True, text=True, timeout=5)
            lines = result.stdout.strip().split('\n')[1:]
            for line in lines:
                parts = line.strip().split()
                if len(parts) >= 2 and parts[1] == "device":
                    new_scan_devices.append(parts[0])
                    if parts[0] not in self.device_locks:
                        self.device_locks[parts[0]] = threading.Lock()
        except subprocess.TimeoutExpired:
            self.log("ADB devices command timed out", level="ERROR")
            return
        except Exception as e:
            self.log(f"Error listing ADB devices: {e}", level="ERROR")
            return

        # Clear stale placeholder labels
        for child in list(self.devices_scroll_frame.winfo_children()):
            if isinstance(child, ctk.CTkLabel) and ("No ADB Devices" in child.cget("text") or "ADB not found" in child.cget("text")):
                child.destroy()

        # Mark disconnected
        for dev_id in list(self.device_ui_elements.keys()):
            if dev_id not in new_scan_devices:
                self.update_device_ui(dev_id, status_text="⚠️ Disconnected", text_color=COLORS["accent_red"])

        # Create cards for new devices
        new_connections = [d for d in new_scan_devices if d not in self.device_ui_elements]
        if new_connections:
            self.log(f"New device(s): {', '.join(new_connections)}", level="SUCCESS")
            for dev_id in new_connections:
                self._create_device_card(dev_id)

        self.devices = list(new_scan_devices)
        self._rebuild_log_terminals(self.devices)

        if not self.devices:
            ctk.CTkLabel(self.devices_scroll_frame,
                         text="⚠️ No ADB Devices found! Connect devices & enable USB Debugging.",
                         text_color=COLORS["accent_red"], font=("Segoe UI", 14, "italic")).pack(pady=20)
            if not on_startup:
                self.log("No devices detected.", level="WARN")

    def _create_device_card(self, dev_id):
        """Build a premium device status card with mini progress bar."""
        card = ctk.CTkFrame(self.devices_scroll_frame, fg_color=COLORS["bg_card"],
                            border_width=1, border_color=COLORS["border"], corner_radius=10, height=110)
        card.pack(fill="x", pady=4, padx=4)
        card.pack_propagate(False)

        # Icon
        ctk.CTkLabel(card, text="📱", font=("Segoe UI", 26)).pack(side="left", padx=(12, 8))

        # Info block
        info = ctk.CTkFrame(card, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True, pady=8)

        dev_lbl = ctk.CTkLabel(info, text=dev_id, font=("Consolas", 17, "bold"),
                               text_color=COLORS["accent_blue"], anchor="w")
        dev_lbl.pack(fill="x")

        detail_row = ctk.CTkFrame(info, fg_color="transparent")
        detail_row.pack(fill="x")

        phone_lbl = ctk.CTkLabel(detail_row, text="📞 Idle", font=("Segoe UI", 16),
                                 text_color=COLORS["text_muted"], anchor="w")
        phone_lbl.pack(side="left")

        elapsed_lbl = ctk.CTkLabel(detail_row, text="⏱ 00:00", font=("Consolas", 12),
                                   text_color=COLORS["text_muted"], anchor="e")
        elapsed_lbl.pack(side="right", padx=(8, 0))

        # Progress row
        prog_row = ctk.CTkFrame(info, fg_color="transparent")
        prog_row.pack(fill="x", pady=(3, 0))

        mini_progress = ctk.CTkProgressBar(prog_row, height=5, width=100,
                                           progress_color=COLORS["accent_blue"],
                                           fg_color=COLORS["border"], corner_radius=3)
        mini_progress.pack(side="left", fill="x", expand=True, padx=(0, 6))
        mini_progress.set(0)

        step_lbl = ctk.CTkLabel(prog_row, text="0% — Ready", font=("Segoe UI", 14),
                                text_color=COLORS["text_muted"], anchor="e", width=110)
        step_lbl.pack(side="right")

        # Buttons
        btn_box = ctk.CTkFrame(card, fg_color="transparent")
        btn_box.pack(side="right", padx=8)

        run_btn = ctk.CTkButton(btn_box, text="▶", width=38, height=32,
                                fg_color=COLORS["accent_green"], hover_color="#059669",
                                text_color="#fff", font=("Segoe UI", 15, "bold"), corner_radius=6,
                                command=lambda d=dev_id: self.start_single_device(d))
        run_btn.pack(side="left", padx=2)

        stop_btn = ctk.CTkButton(btn_box, text="⏹", width=38, height=32,
                                 fg_color=COLORS["accent_red"], hover_color="#dc2626",
                                 text_color="#fff", font=("Segoe UI", 15, "bold"), corner_radius=6,
                                 state="disabled",
                                 command=lambda d=dev_id: self.stop_single_device(d))
        stop_btn.pack(side="left", padx=2)

        status_lbl = ctk.CTkLabel(card, text="● Ready", font=("Segoe UI", 20, "bold"),
                                  text_color=COLORS["accent_emerald"], anchor="e", width=85)
        status_lbl.pack(side="right", padx=6)

        self.device_ui_elements[dev_id] = {
            'frame': card,
            'phone_lbl': phone_lbl,
            'status_lbl': status_lbl,
            'run_btn': run_btn,
            'stop_btn': stop_btn,
            'mini_progress': mini_progress,
            'step_lbl': step_lbl,
            'elapsed_lbl': elapsed_lbl,
            'task_start_time': time.time(),
        }

    # =================================================================
    # ==================== FILE / BROWSE HELPERS =====================
    # =================================================================

    def browse_adb(self):
        file_path = filedialog.askopenfilename(filetypes=[("ADB Executable", "adb.exe"), ("All Files", "*.*")])
        if file_path:
            old_state = self.adb_path_entry.cget("state")
            self.adb_path_entry.configure(state="normal")
            self.adb_path_entry.delete(0, "end")
            self.adb_path_entry.insert(0, file_path)
            self.adb_path_entry.configure(state=old_state)
            self.log(f"Selected ADB path: {file_path}")
            self.save_settings()

            self.save_settings()

    def load_settings(self):
        import json
        settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "toiuu_settings.json")
        
        defaults = {
            "adb_path": r"D:\SIM\pmcong\platform-tools-latest-windows\platform-tools\adb.exe",
            "captcha_offset": "0",
            "phone_file_path": "",
            "ff_token": "aba3dcd7dde85c51ec3454c56a85f77d_304014",
            "ff_service": "1049",
            "ff_country": "vnm"
        }
        
        if os.path.exists(settings_path):
            try:
                with open(settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    defaults.update(data)
            except Exception as e:
                self.log(f"⚠️ Error loading settings JSON: {e}", level="WARN")
                
        # Populate UI entries safely
        self.adb_path_entry.delete(0, "end")
        self.adb_path_entry.insert(0, defaults["adb_path"])
        
        self.captcha_offset_entry.delete(0, "end")
        self.captcha_offset_entry.insert(0, defaults["captcha_offset"])
        
        self.ff_token_entry.delete(0, "end")
        self.ff_token_entry.insert(0, defaults["ff_token"])
        
        self.ff_service_entry.delete(0, "end")
        self.ff_service_entry.insert(0, defaults["ff_service"])
        
        self.ff_country_entry.delete(0, "end")
        self.ff_country_entry.insert(0, defaults["ff_country"])
        
        if defaults.get("phone_file_path") and os.path.exists(defaults["phone_file_path"]):
            self.txt_path = defaults["phone_file_path"]

    def save_settings(self):
        import json
        settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "toiuu_settings.json")
        data = {
            "adb_path": self.adb_path_entry.get().strip(),
            "captcha_offset": self.captcha_offset_entry.get().strip(),
            "phone_file_path": getattr(self, "txt_path", ""),
            "ff_token": self.ff_token_entry.get().strip(),
            "ff_service": self.ff_service_entry.get().strip(),
            "ff_country": self.ff_country_entry.get().strip()
        }
        try:
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self.log(f"⚠️ Error saving settings: {e}", level="WARN")

    def on_close_app(self):
        self.save_settings()
        self.destroy()

    def check_firefox_balance(self):
        token = self.ff_token_entry.get().strip()
        if not token:
            messagebox.showwarning("Warning", "Please enter Firefox Token first.")
            return
            
        self.lbl_balance.configure(text="Checking...", text_color=COLORS["accent_amber"])
        self.update_idletasks()
        
        def fetch():
            try:
                import requests
                res = firefox_api(requests.Session(), token, act="myInfo")
                if res.startswith("1|"):
                    parts = res.split("|")
                    if len(parts) >= 2:
                        balance = parts[1]
                        self.lbl_balance.configure(text=f"💰 {balance} đ", text_color=COLORS["accent_green"])
                    else:
                        self.lbl_balance.configure(text=f"Lỗi: {translate_firefox_error(res, act='myInfo')}", text_color=COLORS["accent_red"])
                else:
                    self.lbl_balance.configure(text=f"Lỗi Token: {translate_firefox_error(res, act='myInfo')}", text_color=COLORS["accent_red"])
            except Exception as e:
                self.lbl_balance.configure(text="Network error", text_color=COLORS["accent_red"])
                
        threading.Thread(target=fetch, daemon=True).start()

    # =================================================================
    # ==================== PHONE LOADING =============================
    # =================================================================

    def update_qty_live(self):
        if not getattr(self, 'is_running', False):
            messagebox.showinfo("Info", "This feature is for modifying amount while running. Tool is stopped, just edit amount and press Start.")
            return
            
        try:
            new_qty = int(self.firefox_qty_entry.get().strip())
            if new_qty <= 0:
                messagebox.showerror("Error", "Amount must be > 0")
                return
                
            if new_qty > self.total_phones:
                diff = new_qty - self.total_phones
                for _ in range(diff):
                    self.phone_queue.put("FIREFOX_AUTO")
                self.total_phones = new_qty
                self.update_stats_ui()
                self.log(f"Đã thêm {diff} luồng chờ. Tổng số lượng mới: {self.total_phones}", level="SUCCESS")
                
            elif new_qty < self.total_phones:
                diff = self.total_phones - new_qty
                removed = 0
                for _ in range(diff):
                    try:
                        self.phone_queue.get_nowait()
                        removed += 1
                    except queue.Empty:
                        break
                
                self.total_phones -= removed
                self.update_stats_ui()
                if removed < diff:
                    self.log(f"Chỉ giảm được {removed} vì các máy đã bắt đầu chạy. Tổng số mới: {self.total_phones}", level="WARN")
                else:
                    self.log(f"Đã giảm {removed} luồng chờ. Tổng số lượng mới: {self.total_phones}", level="SUCCESS")
            else:
                self.log("Amount unchanged.")
                
            self.save_settings()
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid integer.")

    def load_phone_numbers_from_ui(self):
        adb_exec = self.adb_path_entry.get().strip()
        if not os.path.exists(adb_exec):
            messagebox.showerror("Error", f"ADB executable not found at specified path:\\n{adb_exec}")
            return False

        try:
            self.captcha_offset = int(self.captcha_offset_entry.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Captcha Offset must be a numeric integer value (e.g. -30, 0, 50).")
            return False

        try:
            qty = int(self.firefox_qty_entry.get().strip())
            if qty <= 0:
                messagebox.showerror("Error", "Amount must be greater than 0.")
                return False
            self.phone_numbers = ["FIREFOX_AUTO"] * qty
            self.log(f"Set to run {qty} accounts via Firefox API.", level="INFO")
        except ValueError:
            messagebox.showerror("Error", "Amount must be a valid integer.")
            return False

        self.save_settings()
        return True

    # =================================================================
    # ==================== CAMPAIGN CONTROL ==========================
    # =================================================================

    def start_campaign(self):
        if not self.load_phone_numbers_from_ui():
            return

        if not self.devices:
            messagebox.showerror("Error", "No active connected ADB devices to delegate tasks to.")
            return

        self.is_running = True
        self.campaign_start_time = datetime.now()
        self.btn_start.configure(state="disabled")
        self.btn_refresh.configure(state="normal")
        self.btn_stop.configure(state="normal")
        self.adb_path_entry.configure(state="disabled")
        self.captcha_offset_entry.configure(state="disabled")
        self.ff_token_entry.configure(state="disabled")
        self.ff_service_entry.configure(state="disabled")
        self.ff_country_entry.configure(state="disabled")

        self.campaign_progress_bar.set(0)
        self.progress_label.configure(text="Campaign Starting…", text_color=COLORS["accent_green"])

        self.phone_queue = queue.Queue()
        for p in self.phone_numbers:
            self.phone_queue.put((p, 0))

        self.total_phones = len(self.phone_numbers)
        self.success_phones = 0
        self.failed_phones = 0
        self.trigger_success_count = 0
        self.backup_success_count = 0
        self.update_stats_ui()
        
        with self.active_workers_lock:
            self.active_workers_count = 0
        self.running_devices = list(self.devices)
        self.active_running_devices.clear()
        
        adb_exec = self.adb_path_entry.get().strip()
        for dev_id in self.running_devices:
            self.active_running_devices.add(dev_id)
            if dev_id in self.device_ui_elements:
                self.device_ui_elements[dev_id]['run_btn'].configure(state="disabled")
                self.device_ui_elements[dev_id]['stop_btn'].configure(state="normal")
            self.start_worker_thread(dev_id, adb_exec)

    def start_single_device(self, device_id):
        if not hasattr(self, 'phone_queue') or self.phone_queue.empty():
            if not self.load_phone_numbers_from_ui():
                return
            
            self.is_running = True
            self.campaign_start_time = datetime.now()
            self.btn_start.configure(state="disabled")
            self.btn_refresh.configure(state="normal")
            self.btn_stop.configure(state="normal")
            self.adb_path_entry.configure(state="disabled")
            self.captcha_offset_entry.configure(state="disabled")
            self.ff_token_entry.configure(state="disabled")
            self.ff_service_entry.configure(state="disabled")
            self.ff_country_entry.configure(state="disabled")
            
            self.phone_queue = queue.Queue()
            for p in self.phone_numbers:
                self.phone_queue.put(p)
                
            self.total_phones = len(self.phone_numbers)
            self.success_phones = 0
            self.failed_phones = 0
            self.trigger_success_count = 0
            self.backup_success_count = 0
            self.update_stats_ui()
            
            with self.active_workers_lock:
                self.active_workers_count = 0
            self.running_devices = []

        adb_exec = self.adb_path_entry.get().strip()
        
        with self.active_workers_lock:
            if device_id in self.running_devices:
                return
            self.running_devices.append(device_id)
            self.active_running_devices.add(device_id)

        if device_id in self.device_ui_elements:
            self.device_ui_elements[device_id]['run_btn'].configure(state="disabled")
            self.device_ui_elements[device_id]['stop_btn'].configure(state="normal")
            self.update_device_ui(device_id, status_text="● Ready", text_color=COLORS["accent_emerald"])

        self.start_worker_thread(device_id, adb_exec)

    def stop_single_device(self, device_id):
        with self.active_workers_lock:
            if device_id in self.active_running_devices:
                self.active_running_devices.remove(device_id)
        self.log(f"[{device_id}] Stopping device worker… waiting to finish current task.", level="WARN")
        if device_id in self.device_ui_elements:
            self.device_ui_elements[device_id]['stop_btn'].configure(state="disabled")
            self.update_device_ui(device_id, status_text="🛑 Stopping…", text_color=COLORS["accent_red"])

    def start_worker_thread(self, dev_id, adb_path):
        with self.active_workers_lock:
            self.active_workers_count += 1
        self.log(f"[{dev_id}] Worker thread spawned.", level="STEP")
        threading.Thread(target=self.device_worker, args=(dev_id, adb_path), daemon=True).start()

    def device_worker(self, device_id, adb_path):
        # Create per-device session for operations
        session = requests.Session()
        
        while self.is_running and device_id in self.active_running_devices and not self.phone_queue.empty():
            try:
                phone_data = self.phone_queue.get_nowait()
                if isinstance(phone_data, tuple):
                    phone, retry_count = phone_data
                else:
                    phone = phone_data
                    retry_count = 0
            except queue.Empty:
                break

            firefox_pkey = None
            self.update_device_ui(device_id, phone=phone, status_text="🔄 Starting task…", text_color=COLORS["accent_amber"])
            self.log(f"[{device_id}] Starting campaign task for phone {phone} (Retry {retry_count}/{MAX_RETRIES})")
            
            status_str = "FAILED"
            try:
                success = process_device(self, device_id, phone, adb_path, self.captcha_offset, session, firefox_pkey=firefox_pkey)
                if not self.is_running or device_id not in self.active_running_devices:
                    self.log(f"[{device_id}] Thread stopped by user command.", level="WARN")
                    self.update_device_ui(device_id, status_text="🔴 Terminated", text_color=COLORS["accent_red"])
                    self.phone_queue.put((phone, retry_count)) # Put back numbers that weren't fully processed
                    status_str = "TERMINATED"
                    break
                
                if success is True:
                    status_str = "SUCCESS"
                    self.log(f"[{device_id}] ✅ Task completed successfully for {phone}!", level="SUCCESS")
                    self.update_device_ui(device_id, status_text="✅ Done", text_color=COLORS["accent_green"])
                elif success in ["TIMEOUT", "UI_UNKNOWN"]:
                    if retry_count < MAX_RETRIES:
                        status_str = "RETRY"
                        self.log(f"[{device_id}] ⚠️ Task issue ({success}), returning {phone} to queue for retry.", level="WARN")
                        self.update_device_ui(device_id, status_text="⚠️ Retry", text_color=COLORS["accent_amber"])
                        self.phone_queue.put((phone, retry_count + 1))
                    else:
                        status_str = "FAILED"
                        self.log(f"[{device_id}] ❌ Max retries reached for {phone}. Marking as FAILED.", level="ERROR")
                        self.update_device_ui(device_id, status_text="❌ Failed", text_color=COLORS["accent_red"])
                elif success == "TERMINATED":
                    status_str = "TERMINATED"
                    self.log(f"[{device_id}] Thread stopped/disconnected.", level="WARN")
                    self.update_device_ui(device_id, status_text="🔴 Terminated", text_color=COLORS["accent_red"])
                else:
                    last_err = getattr(self, 'device_last_error', {}).get(device_id, "Unknown error")
                    self.log(f"❌ [{device_id}] Task failed for {phone}. Lý do: {last_err}", level="ERROR")
                    self.update_device_ui(device_id, status_text="❌ Failed", text_color=COLORS["accent_red"])
            except Exception as e:
                status_str = "CRASH"
                self.log(f"[{device_id}] ❌ System Error: {e}", level="ERROR")
                self.update_device_ui(device_id, status_text="❌ Crash", text_color=COLORS["accent_red"])
            
            # Retrieve the actual phone number if it was dynamically allocated
            actual_phone = getattr(self, 'actual_phones', {}).get(device_id, phone)
            
            # Increment tracking counters, write file & update UI (thread-safe)
            with self.completed_lock:
                if status_str == "SUCCESS":
                    self.success_phones += 1
                    try:
                        log_file_path = os.path.join(os.path.dirname(__file__), "completed_phones.txt")
                        with open(log_file_path, "a", encoding="utf-8") as lf:
                            lf.write(f"{actual_phone}\n")
                    except Exception as e:
                        self.log(f"[{device_id}] ⚠️ Error writing to completed_phones.txt: {e}", level="WARN")
                elif status_str in ["FAILED", "CRASH"]:
                    self.failed_phones += 1
            self.update_stats_ui()
            
            self.phone_queue.task_done()
            app_sleep(self, 0.5, device_id)

        self.update_device_ui(device_id, status_text="💤 Idle", text_color=COLORS["text_muted"])
        
        # Reset buttons back to idle state
        if device_id in self.device_ui_elements:
            self.device_ui_elements[device_id]['run_btn'].configure(state="normal")
            self.device_ui_elements[device_id]['stop_btn'].configure(state="disabled")

        # Clean from running devices list
        with self.active_workers_lock:
            if device_id in self.running_devices:
                self.running_devices.remove(device_id)
            if device_id in self.active_running_devices:
                self.active_running_devices.remove(device_id)
                
            self.active_workers_count -= 1
            if self.active_workers_count == 0:
                self.after(0, self.campaign_finished)

    def campaign_finished(self):
        self.is_running = False
        elapsed_str = ""
        if self.campaign_start_time:
            el = datetime.now() - self.campaign_start_time
            h, rem = divmod(int(el.total_seconds()), 3600)
            m, s = divmod(rem, 60)
            elapsed_str = f" in {h:02d}:{m:02d}:{s:02d}"
        self.campaign_start_time = None
        self.reset_ui_controls()

        total = self.success_phones + self.failed_phones
        rate = (self.success_phones / total * 100) if total > 0 else 0
        self.log(f"🎉 CAMPAIGN COMPLETE — {self.success_phones}/{total} success ({rate:.0f}%){elapsed_str}", level="SYSTEM")

        self.progress_label.configure(text="Campaign Complete ✅", text_color=COLORS["accent_green"])
        self.uptime_label.configure(text="")

    def stop_campaign(self):
        if self.is_running:
            self.is_running = False
            with self.active_workers_lock:
                self.active_running_devices.clear()
            self.log("🛑 Force stop requested. Waiting for threads to finish…", level="WARN")
            self.btn_stop.configure(state="disabled")
            self.progress_label.configure(text="Stopping…", text_color=COLORS["accent_red"])

    def reset_ui_controls(self):
        self.btn_start.configure(state="normal")
        self.btn_refresh.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.adb_path_entry.configure(state="normal")
        self.captcha_offset_entry.configure(state="normal")
        # Reset all device buttons back to idle state
        for dev_id, elements in self.device_ui_elements.items():
            if 'run_btn' in elements:
                elements['run_btn'].configure(state="normal")
            if 'stop_btn' in elements:
                elements['stop_btn'].configure(state="disabled")

# =====================================================================
# ================= AUTOMATION CORE INTEGRATED FUNCTIONS =============
# =====================================================================

def run_adb(command, timeout=ADB_COMMAND_TIMEOUT, device_id=None, app=None):
    """Execute ADB command with timeout protection, bypassing shell where possible."""
    use_shell = False
    if isinstance(command, str):
        if any(op in command for op in [">", "<", "|", "&&", "||"]):
            use_shell = True
        else:
            import shlex
            try:
                # Use posix=False to preserve Windows backslashes and strip surrounding double quotes
                parts = shlex.split(command, posix=False)
                command = [p.strip('"') for p in parts]
                use_shell = False
            except Exception:
                use_shell = True
                
    try:
        result = subprocess.run(
            command,
            shell=use_shell,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=timeout
        )
        return result
    except subprocess.TimeoutExpired:
        if app and device_id:
            app.log(f"[{device_id}] ⏱️ ADB command timeout (>{timeout}s)")
        return None
    except Exception as e:
        if app and device_id:
            app.log(f"[{device_id}] ⚠️ ADB error: {e}")
        return None

def detect_page(xml_data):
    """Detect current page state from UI XML using text and content-desc attribute values to avoid false positives."""
    if not xml_data:
        return "UNKNOWN"
    
    # Extract text="..." and content-desc="..." values from XML to avoid matching resource-name, class, etc.
    text_values = ' '.join(re.findall(r'(?:text|content-desc)="([^"]*)"', xml_data)).lower()
    
    if any(t in text_values for t in ["banned", "locked", "blocked", "invalid", "suspended", "locked", "temporarily locked", "invalid", "disabled", "banned"]):
        return "BANNED"
    if any(t in text_values for t in [
        "verification code",
        "enter your verification",
        "enter the 6-digit code",
        "send sms",
        "text zalo to",
        "8500",
        "7539",
        "didn't receive code",
        "resend",
        "verification code",
        "enter code",
        "send message",
        "compose",
        "verify",
    ]):
        return "OTP"
    if any(t in text_values for t in ["in progress", "processing", "processing"]):
        return "PROCESSING"
    if any(t in text_values for t in ["captcha", "puzzle", "puzzle", "slide"]):
        return "CAPTCHA"
    if any(t in text_values for t in ["profile name", "your name", "your name", "enter name", "to provide"]):
        return "PROFILE_NAME"
    if any(t in text_values for t in ["birthday", "birthday", "gender", "gender"]):
        return "BIRTHDAY_GENDER"
    if any(t in text_values for t in ["khui 7up", "7up", "mini app"]):
        return "MINI_APP"
    if any(t in text_values for t in ["follow", "interested", "follow oa"]):
        return "FOLLOW"
    if any(t in text_values for t in ["joined", "success", "success", "completed"]):
        return "SUCCESS_PAGE"
    
    return "UNKNOWN"

def is_7up_campaign_screen(xml_data):
    if not xml_data:
        return False

    text_values = ' '.join(
        re.findall(r'(?:text|content-desc)="([^"]*)"', xml_data)
    ).lower()

    keywords = [
        "khui 7up",
        "hunt gifts now",
        "the le tham du",
        "rules of participation",
        "toi da doc",
        "i have read",
        "chinh sach quyen rieng tu",
        "privacy policy",
        "home",
        "scan qr",
        "profile",
    ]

    return any(k.lower() in text_values for k in keywords)

def quick_scan_7up_screen(app, device_id, adb_path, timeout=6):
    start_time = time.time()

    while time.time() - start_time < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return False

        xml_data = get_ui_xml(device_id, adb_path, app)
        if is_7up_campaign_screen(xml_data):
            app.log(f"[{device_id}] Detected 7UP campaign screen by quick scan.")
            return True

        app_sleep(app, 0.1, device_id)

    return False

def get_screen_texts(xml_data, limit=20):
    if not xml_data:
        return []

    texts = re.findall(r'(?:text|content-desc)="([^"]*)"', xml_data)
    texts = [t.strip() for t in texts if t and t.strip()]
    return texts[:limit]


def scan_page_with_log(app, device_id, adb_path, note=""):
    xml_data = get_ui_xml(device_id, adb_path, app)
    page = detect_page(xml_data)
    texts = get_screen_texts(xml_data)

    app.log(f"[{device_id}] Screen scan {note}: page={page}, texts={texts}")
    return page, xml_data, texts

def wait_page_after_captcha(app, device_id, adb_path, timeout=30):
    start = time.time()
    captcha_seen_since = None

    while time.time() - start < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return "STOPPED"

        xml_data = get_ui_xml(device_id, adb_path, app)
        page = detect_page(xml_data)

        if page in ["OTP", "BANNED", "PROFILE_NAME"]:
            return page

        if page == "PROCESSING":
            captcha_seen_since = None
            texts = get_screen_texts(xml_data)
            app.log(f"[{device_id}] Zalo đang xử lý sau captcha, tiếp tục chờ... texts={texts}")
        elif page == "CAPTCHA":
            if captcha_seen_since is None:
                captcha_seen_since = time.time()

            # chỉ coi là vẫn captcha nếu giữ nguyên >= 10 giây
            if time.time() - captcha_seen_since >= 10:
                return "CAPTCHA"
        else:
            captcha_seen_since = None
            texts = get_screen_texts(xml_data)
            app.log(f"[{device_id}] Chưa rõ màn hình sau captcha, page={page}, texts={texts}")

        app_sleep(app, 1, device_id)

    return "TIMEOUT"

def app_sleep(app, seconds, device_id=None):
    start = time.time()
    while time.time() - start < seconds:
        if not app.is_running:
            break
        if device_id and device_id not in app.active_running_devices:
            break
        time.sleep(0.1)

def adb_click(app, device_id, x_y_string, adb_path):
    coords = x_y_string.split()
    with app.device_locks.setdefault(device_id, threading.Lock()):
        if len(coords) == 2:
            cmd = [adb_path, "-s", device_id, "shell", "input", "tap", coords[0], coords[1]]
            run_adb(cmd, timeout=5, device_id=device_id, app=app)
        else:
            run_adb(f'"{adb_path}" -s {device_id} shell input tap {x_y_string}', timeout=5, device_id=device_id, app=app)
    app_sleep(app, 0.4, device_id)

def adb_type(app, device_id, text, adb_path, slow=False):
    text_safe = text.replace(" ", "%s")
    with app.device_locks.setdefault(device_id, threading.Lock()):
        if slow:
            for char in text_safe:
                cmd = [adb_path, "-s", device_id, "shell", "input", "text", char]
                run_adb(cmd, timeout=3, device_id=device_id, app=app)
                time.sleep(0.15)
        else:
            cmd = [adb_path, "-s", device_id, "shell", "input", "text", text_safe]
            run_adb(cmd, timeout=5, device_id=device_id, app=app)
    app_sleep(app, 0.4, device_id)

def adb_type_digits(app, device_id, digits, adb_path):
    """Sends native hardware-like digit keyevents (7-16) to bypass Vietnamese IME/Telex auto-correct issues."""
    with app.device_locks.setdefault(device_id, threading.Lock()):
        for char in digits:
            if char.isdigit():
                keycode = str(7 + int(char))
                run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", keycode], timeout=3, device_id=device_id, app=app)
                time.sleep(0.05)
    app_sleep(app, 0.4, device_id)

def get_screen_size(device_id, adb_path, app=None):
    try:
        result = run_adb(f'"{adb_path}" -s {device_id} shell wm size', timeout=8, device_id=device_id, app=app)
        if result:
            match = re.search(r'(\d+)x(\d+)', result.stdout)
            if match:
                return int(match.group(1)), int(match.group(2))
    except:
        pass
    return 1080, 2220

def get_ui_xml(device_id, adb_path, app=None):
    import unicodedata
    for attempt in range(2):
        with app.device_locks.setdefault(device_id, threading.Lock()) if app else threading.Lock():
            result = run_adb(
                f'"{adb_path}" -s {device_id} shell "uiautomator dump /data/local/tmp/window_dump.xml >/dev/null && cat /data/local/tmp/window_dump.xml"',
                timeout=UI_DUMP_TIMEOUT,
                device_id=device_id,
                app=app
            )
        if result:
            xml = result.stdout
            if xml and "<node" in xml:
                if app and hasattr(app, 'device_errors'):
                    app.device_errors[device_id] = 0
                return unicodedata.normalize('NFC', xml)
        time.sleep(0.3)
        
    if app:
        if not hasattr(app, 'device_errors'):
            app.device_errors = {}
        app.device_errors[device_id] = app.device_errors.get(device_id, 0) + 1
        if app.device_errors[device_id] >= 4:
            app.log(f"[{device_id}] ⚠️ PHÁT HIỆN THIẾT BỊ TREO/LAG (UI Dump thất bại 4 lần liên tiếp). Ép dừng tiến trình để tránh kẹt số!", level="ERROR")
            app.active_running_devices.discard(device_id)
            
    return ""

def adb_focus_input(app, device_id, adb_path):
    xml_data = get_ui_xml(device_id, adb_path, app)
    match = re.search(r'node.*?class="[^"]*EditText".*?bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', xml_data)
    if match:
        x1, y1, x2, y2 = match.groups()
        click_x = (int(x1) + int(x2)) // 2
        click_y = (int(y1) + int(y2)) // 2
        app.log(f"[{device_id}] Focus input EditText...")
        adb_click(app, device_id, f"{click_x} {click_y}", adb_path)
        return True
    return False

def adb_click_text(app, device_id, target_text, adb_path, click_all=False, align="center", exact_match=False, override_x=None, match_index=0, silent=False):
    import unicodedata
    if not isinstance(target_text, list):
        target_text = [target_text]
        
    xml_data = get_ui_xml(device_id, adb_path, app)
    if not xml_data:
        return False
    # Parse text, content-desc, and bounds in correct order
    matches = re.findall(r'node[^>]*text="([^"]*)"[^>]*content-desc="([^"]*)"[^>]*bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', xml_data)
    
    clicked = False
    current_index = 0
    for text, desc, x1, y1, x2, y2 in matches:
        match_found = False
        clean_text = text.strip()
        clean_desc = desc.strip()
        
        # Normalize targets to NFC
        normalized_targets = [unicodedata.normalize('NFC', t).lower() for t in target_text]
        
        if exact_match:
            match_found = any(nt == clean_text.lower() or nt == clean_desc.lower() for nt in normalized_targets)
        else:
            match_found = any(nt in clean_text.lower() or nt in clean_desc.lower() for nt in normalized_targets)
            
        if match_found:
            if current_index < match_index:
                current_index += 1
                continue
                
            if align == "left":
                click_x = int(x1) + 20
            elif align == "outside_left":
                click_x = int(x1) - 20
            else:
                click_x = (int(x1) + int(x2)) // 2
                
            click_y = (int(y1) + int(y2)) // 2
            
            if override_x is not None:
                click_x = override_x
            
            if not silent:
                app.log(f"[{device_id}] Click '{text or desc}' at ({click_x}, {click_y})")
            adb_click(app, device_id, f"{click_x} {click_y}", adb_path)
            clicked = True
            if not click_all:
                return True
            
    return clicked

def check_text_exists(device_id, target_text, adb_path, app=None):
    import unicodedata
    if not isinstance(target_text, list):
        target_text = [target_text]
    xml_data = get_ui_xml(device_id, adb_path, app)
    if not xml_data:
        return False
    visible_texts = ' '.join(re.findall(r'(?:text|content-desc)="([^"]*)"', xml_data)).lower()
    return any(unicodedata.normalize('NFC', t).lower() in visible_texts for t in target_text)

def wait_for_text(app, device_id, target_text, adb_path, timeout=30):
    if not isinstance(target_text, list):
        target_text = [target_text]
    
    app.log(f"[{device_id}] ⏳ Waiting for text {target_text} (Max: {timeout}s)", level="STEP")
    start_time = time.time()
    while time.time() - start_time < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return False
        if check_text_exists(device_id, target_text, adb_path, app):
            app_sleep(app, 0.2, device_id)
            return True
        app_sleep(app, POLL_INTERVAL, device_id)
    app.log(f"[{device_id}] ❌ Timeout waiting for text {target_text}", level="ERROR")
    return False

def solve_zalo_captcha(app, device_id, screen_w, screen_h, adb_path, offset_captcha, session):
    app.log(f"[{device_id}] Solving Zalo puzzle slider captcha...")
    safe_id = device_id.replace(':', '_').replace('.', '_')
    local_img_path = os.path.abspath(os.path.join(os.path.dirname(__file__), f"captcha_{safe_id}.png"))
    
    # Take screenshot binary-safely with timeout and compatibility fallback
    screencap_success = False
    try:
        result = subprocess.run(
            [adb_path, "-s", device_id, "exec-out", "screencap", "-p"],
            capture_output=True,
            timeout=15,
            shell=False
        )
        if result.returncode == 0 and result.stdout and result.stdout.startswith(b"\x89PNG"):
            with open(local_img_path, "wb") as f:
                f.write(result.stdout)
            screencap_success = True
    except Exception as e:
        app.log(f"[{device_id}] exec-out screencap failed: {e}. Trying compatibility fallback...")
        
    if not screencap_success:
        try:
            device_temp_path = "/data/local/tmp/temp_cap.png"
            run_adb([adb_path, "-s", device_id, "shell", "screencap", "-p", device_temp_path], timeout=12, device_id=device_id, app=app)
            pull_res = subprocess.run([adb_path, "-s", device_id, "pull", device_temp_path, local_img_path], capture_output=True, timeout=12)
            run_adb([adb_path, "-s", device_id, "shell", "rm", device_temp_path], timeout=5, device_id=device_id, app=app)
            if pull_res.returncode == 0 and os.path.exists(local_img_path) and os.path.getsize(local_img_path) > 0:
                screencap_success = True
        except Exception as e:
            app.log(f"[{device_id}] Fallback screencap exception: {e}")

    if not screencap_success or not os.path.exists(local_img_path) or os.path.getsize(local_img_path) == 0:
        app.log(f"[{device_id}] Screencap file missing or empty.")
        return False
        
    try:
        try:
            from PIL import Image
            import io
            with Image.open(local_img_path) as img:
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                # Nén ảnh thành JPEG quality 65 giúp giảm 80-90% dung lượng file (từ ~1MB xuống ~80KB)
                # Kích thước ảnh siêu nhẹ giúp gửi lên server OmoCaptcha mất chưa tới 0.1 giây (thay vì 1-2s)
                buffer = io.BytesIO()
                img.save(buffer, format="JPEG", quality=65)
                img_b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
                app.log(f"[{device_id}] Đã nén ảnh Captcha siêu nhẹ để tăng tốc độ upload.")
        except ImportError:
            # Fallback phòng hờ máy chưa cài thư viện Pillow
            with open(local_img_path, "rb") as img_file:
                img_b64 = base64.b64encode(img_file.read()).decode("utf-8")
            
        client_key = "PKG_NBFIXS0EN93XZQNZJ2DCC6ELE4VXIDL9ILDEJ7GP6KV3EGY5NFYS7UBGTHXDIQ1781082254"
        payload = {
            "clientKey": client_key,
            "task": {
                "type": "ZaloSliderPhoneTask",
                "imageBase64": img_b64
            }
        }
        
        app.log(f"[{device_id}] Submitting task to OmoCaptcha...")
        task_id = None
        max_retries = 3
        for attempt in range(max_retries):
            if not app.is_running or device_id not in app.active_running_devices: return False
            try:
                res = session.post("https://api.omocaptcha.com/v2/createTask", json=payload, timeout=8)
                data = res.json()
                if data.get("errorId") != 0:
                    err_code = data.get("errorCode")
                    app.log(f"[{device_id}] OmoCaptcha Error: {err_code}")
                    if err_code in ["ERROR_SERVICE_UNAVAILABLE", "ERROR_REDIS_UNAVAILABLE", "ERROR_TASK_IS_MAINTENANCE"]:
                        time.sleep(2 ** attempt)
                        continue
                    return False
                else:
                    task_id = data.get("taskId")
                    break
            except Exception as e:
                time.sleep(2 ** attempt)
                
        if not task_id:
            return False
            
        app.log(f"[{device_id}] Task created: {task_id}. Awaiting solution...")
        swipe_distance = 0
        api_start_x = None
        api_start_y = None
        api_end_x = None
        
        for _ in range(50):
            if not app.is_running or device_id not in app.active_running_devices: return False
            time.sleep(0.7)
            try:
                poll_res = session.post("https://api.omocaptcha.com/v2/getTaskResult", json={"clientKey": client_key, "taskId": task_id}, timeout=5)
                poll_data = poll_res.json()
                if poll_data.get("errorId") != 0:
                    continue
                status = poll_data.get("status")
                if status == "ready":
                    solution = poll_data.get("solution", {})
                    app.log(f"[{device_id}] Captcha solution fetched!")
                    if "start" in solution and "end" in solution:
                        api_start_x = int(float(solution["start"].get("x", 0)))
                        api_start_y = int(float(solution["start"].get("y", 0)))
                        api_end_x = int(float(solution["end"].get("x", 0)))
                        swipe_distance = api_end_x - api_start_x
                    elif "distance" in solution:
                        swipe_distance = int(float(solution["distance"]))
                    elif "x" in solution:
                        swipe_distance = int(float(solution["x"]))
                    break
                elif status == "fail":
                    return False
            except Exception as e:
                pass
                
        if swipe_distance <= 0:
            return False
            
        if api_start_x and api_start_y and api_end_x:
            btn_start_x = api_start_x
            btn_start_y = api_start_y
            end_x = api_end_x + offset_captcha
        else:
            btn_start_x = int(screen_w * (299 / 1080))
            btn_start_y = int(screen_h * (1250 / 2220))
            end_x = btn_start_x + swipe_distance + offset_captcha
        
        duration = random.randint(250, 400)
        cmd = [
            adb_path, "-s", device_id, "shell", "input", "swipe",
            str(btn_start_x), str(btn_start_y), str(end_x),
            str(btn_start_y + random.randint(-1, 1)), str(duration)
        ]
        run_adb(cmd, timeout=8, device_id=device_id, app=app)
        return True
    except Exception as e:
        app.log(f"[{device_id}] Captcha Exception: {e}")
        return False
    finally:
        if os.path.exists(local_img_path):
            try: os.remove(local_img_path)
            except: pass

# Main Single Device execution stream
def process_device(app, device_id, phone, adb_path, offset_captcha, session, firefox_pkey=None):
    if hasattr(app, 'device_last_error'):
        app.device_last_error.pop(device_id, None)
    ff_ctx = {"pkey": firefox_pkey, "otp_received": False, "trigger_success": False, "backup_success": False}
    status_str = "FAILED"
    try:
        result = _process_device_internal(app, device_id, phone, adb_path, offset_captcha, session, ff_ctx)
        
        if not app.is_running or device_id not in app.active_running_devices:
            result = "TERMINATED"

        if result is True:
            status_str = "SUCCESS"
        elif result in ["TIMEOUT", "UI_UNKNOWN"]:
            status_str = "TIMEOUT"
        elif result == "TERMINATED":
            status_str = "TERMINATED"
        return result
    finally:
        # EXPORT STATS TO EXCEL
        actual_phone = getattr(app, 'actual_phones', {}).get(device_id, phone)
        if actual_phone and actual_phone != "FIREFOX_AUTO":
            log_to_excel(actual_phone, ff_ctx["otp_received"], ff_ctx["trigger_success"], ff_ctx["backup_success"])

        ff_token = app.ff_token_entry.get().strip()
        if ff_ctx.get("pkey"):
            if status_str == "SUCCESS":
                firefox_api_return(session, ff_token, ff_ctx["pkey"], "0")
                app.log(f"[{device_id}] Đã feedback SUCCESS lên Firefox cho số {phone}")
            elif status_str in ["FAILED", "TIMEOUT", "UI_UNKNOWN"]:
                firefox_add_black(session, ff_token, ff_ctx["pkey"])
                last_err = getattr(app, 'device_last_error', {}).get(device_id, "Unknown error")
                app.log(f"❌ [{device_id}] Đã Blacklist số {phone} trên hệ thống Firefox. Lý do: {last_err}", level="ERROR")
            elif status_str in ["TERMINATED", "RETRY"]:
                firefox_set_rel(session, ff_token, ff_ctx["pkey"])
                app.log(f"[{device_id}] Đã release/setRel số {phone} trên Firefox (trạng thái: {status_str}).")

def _process_device_internal(app, device_id, phone, adb_path, offset_captcha, session, ff_ctx):
    try:
        if not app.is_running or device_id not in app.active_running_devices: return "TERMINATED"
        
        screen_w, screen_h = get_screen_size(device_id, adb_path, app)
        app.log(f"[{device_id}] Resolution: {screen_w}x{screen_h}")
        
        # Setup ADB reverse port forwarding for local Server webhook
        app.update_device_ui(device_id, status_text="🔄 Setting up reverse proxy...", text_color="#3b82f6")
        run_adb(f'"{adb_path}" -s {device_id} reverse tcp:5000 tcp:5000', timeout=8, device_id=device_id, app=app)
        app.log(f"[{device_id}] Configured ADB reverse tcp:5000 tcp:5000")
        
        COORD_O_NHAP_SDT = f"{int(screen_w * 0.5)} {int(screen_h * 0.31)}"
        COORD_BTN_TIEP_TUC = f"{int(screen_w * 0.89)} {int(screen_h * 0.94)}"
        COORD_BTN_XAC_NHAN = f"{int(screen_w * 0.5)} {int(screen_h * 0.52)}"
        SWIPE_SCROLL_UP = f"{int(screen_w * 0.5)} {int(screen_h * 0.73)} {int(screen_w * 0.5)} {int(screen_h * 0.31)} 300"
        
        app.update_device_ui(device_id, status_text="🔄 Starting Xtoolz...", text_color="#38bdf8")
        
        # Reset phone ID
        app.update_device_ui(device_id, status_text="🔄 Rebooting device ID...", text_color="#e0f2fe")
        
        max_reset_retry = 5
        reset_accepted = False
        offline_timeout_loops = int(35 / 0.5) # 35 seconds
        for attempt in range(1, max_reset_retry + 1):
            if not app.is_running or device_id not in app.active_running_devices: return False
            app.log(f"[{device_id}] Reset 4G attempt {attempt}/{max_reset_retry}")

            XTOOLZ_PACKAGE = "app.pixel.xtoolz"
            launch_res = run_adb(
                [adb_path, "-s", device_id, "shell", "monkey", "-p", XTOOLZ_PACKAGE, "-c", "android.intent.category.LAUNCHER", "1"],
                timeout=8,
                device_id=device_id,
                app=app
            )

            if not launch_res or launch_res.returncode != 0:
                stdout = launch_res.stdout.strip() if launch_res else ""
                stderr = launch_res.stderr.strip() if launch_res else ""
                err_text = f"{stdout} {stderr}".lower()

                if any(x in err_text for x in ["device not found", "offline", "no devices", "not found"]):
                    app.log(f"[{device_id}] ADB đã ngắt trong lúc mở Xtoolz, coi như reset đã ăn. Chuyển sang chờ reconnect...")
                    reset_accepted = True
                    break

                app.log(
                    f"[{device_id}] Không mở được Xtoolz package={XTOOLZ_PACKAGE}. stdout={stdout} stderr={stderr}",
                    level="ERROR"
                )
                app_sleep(app, 5, device_id)
                continue
            
            app.log(f"[{device_id}] Launched Xtoolz app. Waiting for UI (10s)...")
            app_sleep(app, 4.0, device_id)

            # Đợi có mạng (4G) trước khi click Reset
            app.update_device_ui(device_id, status_text="🔄 Waiting for 4G...", text_color="#fbbf24")
            app.log(f"[{device_id}] Đang chờ kết nối 4G ổn định...")
            while True:
                if not app.is_running or device_id not in app.active_running_devices: return False
                res = run_adb(f'"{adb_path}" -s {device_id} shell "ping -c 1 -w 2 8.8.8.8"', timeout=5, device_id=device_id, app=app)
                if res and res.returncode == 0:
                    app.log(f"[{device_id}] Đã có kết nối mạng (4G)!")
                    break
                app_sleep(app, 0.5, device_id)


            # Try dynamic click first
            if not adb_click_text(app, device_id, ["Reset", "New ID", "Reboot", "Change", "Random"], adb_path, silent=False):
                app.log(f"[{device_id}] ⚠️ Dynamic reset button not found, clicking fallback coordinates...", level="WARN")
                COORD_BTN_RESET_XTOOLZ = f"{int(screen_w * 0.86)} {int(screen_h * 0.09)}"
                adb_click(app, device_id, COORD_BTN_RESET_XTOOLZ, adb_path)
            
            # Monitor adb disconnection
            app.log(f"[{device_id}] Waiting for device offline...")
            offline = False
            for _ in range(offline_timeout_loops):
                if not app.is_running or device_id not in app.active_running_devices: return False
                res = run_adb(f'"{adb_path}" -s {device_id} get-state', timeout=5, device_id=device_id, app=app)
                if res and res.stdout.strip() != "device":
                    offline = True
                    break
                app_sleep(app, 0.5, device_id)
                
            if offline:
                app.log(f"[{device_id}] Device went offline, reset accepted.")
                reset_accepted = True
                break
                
            app.log(f"[{device_id}] ADB chưa ngắt sau reset, quay lại Xtoolz bấm lại...", level="WARN")

        if not reset_accepted:
            app.log(f"[{device_id}] Reset 4G failed: ADB never went offline after retries.", level="ERROR")
            return False
        
        # Monitor adb connection
        app.update_device_ui(device_id, status_text="🔄 Rebooting (Wait ADB)...", text_color="#fbbf24")
        app.log(f"[{device_id}] Waiting for device to reconnect...")
        reconnect_start = time.time()
        reconnect_success = False
        while app.is_running and device_id in app.active_running_devices:
            if time.time() - reconnect_start > 120: # 2 minutes timeout
                break
            res = run_adb(f'"{adb_path}" -s {device_id} get-state', timeout=5, device_id=device_id, app=app)
            if res and res.stdout.strip() == "device":
                reconnect_success = True
                break
            app_sleep(app, 1.0, device_id)
            
        if not reconnect_success or not app.is_running or device_id not in app.active_running_devices:
            app.log(f"[{device_id}] ❌ Reconnect timeout or stopped by user.")
            return False
        
        # Buffer wait after reconnection to allow system to stabilize
        app.log(f"[{device_id}] Device reconnected. Waiting for system stabilization...")
        app_sleep(app, 3.0, device_id)
        
        # Wait boot completed
        app.update_device_ui(device_id, status_text="🔄 Booting OS...", text_color="#fbbf24")
        boot_start = time.time()
        boot_success = False
        while app.is_running and device_id in app.active_running_devices:
            if time.time() - boot_start > 120: # 2 minutes timeout
                break
            result = run_adb(f'"{adb_path}" -s {device_id} shell getprop sys.boot_completed', timeout=8, device_id=device_id, app=app)
            if result and "1" in result.stdout:
                boot_success = True
                break
            app_sleep(app, 1.5, device_id)
            
        if not boot_success or not app.is_running or device_id not in app.active_running_devices:
            app.log(f"[{device_id}] ❌ Boot completion timeout or stopped by user.")
            return False
        
        # Critical buffer - system services still initializing even after boot_completed=1
        # Device needs at least 6-8 seconds after boot_completed before any operations
        app.log(f"[{device_id}] ✅ Boot completed detected. Waiting for full system initialization (6s)...")
        app_sleep(app, 20, device_id)
        
        # Verify network connectivity - MUST have 4G/WiFi before launching Zalo
        app.update_device_ui(device_id, status_text="🔄 Verifying 4G/WiFi...", text_color="#fbbf24")
        app.log(f"[{device_id}] Checking 4G/WiFi connectivity...")
        net_start = time.time()
        net_success = False
        
        while app.is_running and device_id in app.active_running_devices:
            if time.time() - net_start > 1000000: # 4 minutes max wait for network
                break
            
            ping_res = run_adb(f'"{adb_path}" -s {device_id} shell ping -c 1 -W 3 8.8.8.8', timeout=8, device_id=device_id, app=app)
            if ping_res and ("1 received" in ping_res.stdout or "1 packets received" in ping_res.stdout or "0% packet loss" in ping_res.stdout):
                net_success = True
            else:
                # Fallback check using dumpsys connectivity
                conn_res = run_adb(f'"{adb_path}" -s {device_id} shell dumpsys connectivity', timeout=5, device_id=device_id, app=app)
                if conn_res and any(x in conn_res.stdout.lower() for x in ["state: connected", "status: connected", "connected/connected"]):
                    net_success = True
                    
            if net_success:
                app.log(f"[{device_id}] ✅ Network online - 4G/WiFi confirmed. Ready to launch Zalo")
                # Sync Android system time with network NTP to prevent Zalo rejecting OTP due to clock skew
                run_adb([adb_path, "-s", device_id, "shell", "settings", "put", "global", "auto_time", "1"], timeout=5, device_id=device_id, app=app)
                run_adb([adb_path, "-s", device_id, "shell", "settings", "put", "global", "auto_time_zone", "1"], timeout=5, device_id=device_id, app=app)
                app.log(f"[{device_id}] Synchronized system time via network NTP.")
                break
            
            app_sleep(app, 2.0, device_id)
            
        # Network REQUIRED - cannot proceed without 4G/WiFi
        if not net_success or not app.is_running or device_id not in app.active_running_devices:
            app.log(f"[{device_id}] ❌ Network unavailable after {int(time.time() - net_start)}s - aborting Zalo launch")
            return False
        
        # Network confirmed - safe to launch Zalo
        app_sleep(app, 1.0, device_id)
        
       
        
        # Launch Zalo
        app.update_device_ui(device_id, status_text="🔄 Launching Zalo...", text_color="#a855f7")
        
        zalo_launched = False
        for launch_attempt in range(2):
            if not app.is_running or device_id not in app.active_running_devices: return False
            
            # Wake up and dismiss keyguard to make sure Zalo opens on top and screen is interactable
            run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "224"], timeout=5, device_id=device_id, app=app)
            app_sleep(app, 0.3, device_id)
            run_adb([adb_path, "-s", device_id, "shell", "wm", "dismiss-keyguard"], timeout=5, device_id=device_id, app=app)
            app_sleep(app, 0.3, device_id)
            
            run_adb(f'"{adb_path}" -s {device_id} shell monkey -p com.zing.zalo -c android.intent.category.LAUNCHER 1', timeout=8, device_id=device_id, app=app)
            app.log(f"[{device_id}] Launched Zalo app (Attempt {launch_attempt+1}/2). Waiting for UI...")
            
            # Dismiss Google Smart Lock / "Sign in with ease" popups that overlay Zalo
            for _ in range(2):
                if check_text_exists(device_id, ["Sign in with ease", "None of the above", "Skip", "Smart Lock"], adb_path):
                    app.log(f"[{device_id}] Detected Google popup overlay. Dismissing...")
                    adb_click_text(app, device_id, ["None of the above", "Skip", "Cancel", "No thanks"], adb_path)
                    app_sleep(app, 1.0, device_id)
            
            if wait_for_text(app, device_id, ["Create new account", "Sign up", "Sign up"], adb_path, timeout=30):
                zalo_launched = True
                break
                
            app.log(f"[{device_id}] ⚠️ Zalo sign-up screen not detected on attempt {launch_attempt+1}, retrying launch...")
            app_sleep(app, 2.0, device_id)
            
        if not zalo_launched:
            app.log(f"[{device_id}] Zalo did not show sign-up screen after retries")
            return False
        
        # Click Sign up
        for i in range(3):
            if adb_click_text(app, device_id, ["Create new account", "Sign up", "Sign up"], adb_path, silent=(i > 0)):
                app_sleep(app, 0.1, device_id)
            else:
                break
                
        phone_screen_ok = wait_for_text(
            app, device_id,
            ["Phone number", "Phone number", "Enter phone number"],
            adb_path,
            timeout=40
        )
        
        if not phone_screen_ok:
            app.log(f"[{device_id}] Phone input text not detected, trying input focus fallback...", level="WARN")
        
            # Tap expected phone input area once
            input_x = int(screen_w * 0.5)
            input_y = int(screen_h * 0.38)
            run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(input_x), str(input_y)], timeout=5, device_id=device_id, app=app)
            app_sleep(app, 0.5, device_id)
        
            # Recheck XML after fallback
            phone_screen_ok = wait_for_text(
                app, device_id,
                ["Phone number", "Phone number", "Enter phone number"],
                adb_path,
                timeout=15
            )
        
        if not phone_screen_ok:
            xml = get_ui_xml(device_id, adb_path, app)
            texts = re.findall(r'(?:text|content-desc)="([^"]*)"', xml)
            app.log(f"[{device_id}] Phone screen still unknown. Visible texts: {texts[:20]}", level="WARN")
            return "UI_UNKNOWN"
        
        # Enter Phone
        if phone == "FIREFOX_AUTO":
            app.update_device_ui(device_id, status_text="🔄 Getting Phone...", text_color=COLORS["accent_amber"])
            ff_token = app.ff_token_entry.get().strip()
            ff_service = app.ff_service_entry.get().strip()
            ff_country = app.ff_country_entry.get().strip()
            
            for retry in range(3):
                pkey, ff_phone_or_err = firefox_get_phone(session, ff_token, ff_service, ff_country)
                if pkey:
                    break
                app.log(f"[{device_id}] Lỗi thuê số (lần {retry+1}/3): {ff_phone_or_err}", level="WARN")
                if isinstance(ff_phone_or_err, str):
                    if any(fatal in ff_phone_or_err for fatal in ["No number", "Invalid token", "Insufficient balance", "Too many numbers", "disabled"]):
                        break # Stop retrying on fatal errors
                app_sleep(app, 5.0, device_id)
                
            if not pkey:
                app.log(f"[{device_id}] Bỏ cuộc, không thuê được số sau 3 lần thử.", level="ERROR")
                app.update_device_ui(device_id, status_text="❌ Rent error", text_color=COLORS["accent_red"])
                return False
                
            ff_ctx["pkey"] = pkey
            phone = ff_phone_or_err
            if not hasattr(app, 'actual_phones'):
                app.actual_phones = {}
            app.actual_phones[device_id] = phone
            app.update_device_ui(device_id, phone=phone)
            app.log(f"[{device_id}] Đã lấy số thành công: {phone}")

        app.update_device_ui(device_id, status_text="🔄 Entering Phone...", text_color="#22c55e")
        adb_click(app, device_id, COORD_O_NHAP_SDT, adb_path)
        adb_type_digits(app, device_id, phone, adb_path)
        
        # Tick terms checkboxes
        app.update_device_ui(device_id, status_text="🔄 Ticking terms...", text_color="#38bdf8")
        
        # Try dynamic lookup from XML layout dump first
        xml_data = get_ui_xml(device_id, adb_path, app)
        checkbox_matches = re.findall(r'node[^>]*class="[^"]*CheckBox"[^>]*checked="([^"]*)"[^>]*bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', xml_data)
        if not checkbox_matches:
            checkbox_matches = re.findall(r'node[^>]*checkable="true"[^>]*checked="([^"]*)"[^>]*bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', xml_data)
        
        clicked_dynamic = False
        if checkbox_matches:
            app.log(f"[{device_id}] Found {len(checkbox_matches)} checkboxes dynamically.")
            clicked_dynamic = True
            for idx, cb in enumerate(checkbox_matches):
                checked_state, x1, y1, x2, y2 = cb
                cx = (int(x1) + int(x2)) // 2
                cy = (int(y1) + int(y2)) // 2
                if checked_state.lower() == "false":
                    app.log(f"[{device_id}] Checkbox {idx+1} is unchecked. Clicking at ({cx}, {cy})")
                    adb_click(app, device_id, f"{cx} {cy}", adb_path)
                else:
                    app.log(f"[{device_id}] Checkbox {idx+1} is already checked. Skipping click.")
        
        if not clicked_dynamic:
            # Fallback to static coordinates ratio if dynamic detection fails
            app.log(f"[{device_id}] Checkbox XML parsing failed, using scaled coordinate fallback...")
            cb1_x = int(screen_w * (100 / 1080))
            cb1_y = int(screen_h * (615 / 2220))
            cb2_x = int(screen_w * (100 / 1080))
            cb2_y = int(screen_h * (730 / 2220))
            
            adb_click(app, device_id, f"{cb1_x} {cb1_y}", adb_path)
            adb_click(app, device_id, f"{cb2_x} {cb2_y}", adb_path)
        
        if not adb_click_text(app, device_id, ["Next", "Continue"], adb_path):
            adb_click(app, device_id, COORD_BTN_TIEP_TUC, adb_path)
            
        app.log(f"[{device_id}] Zalo entered, waiting for confirmation dialog...")
        confirm_ok = wait_for_text(app, device_id, ["Next", "Confirm", "Agree", "Confirm", "Continue", "Continue"], adb_path, timeout=10)
        if confirm_ok:
            app.log(f"[{device_id}] Clicking Confirm on phone number dialog...")
            adb_click_text(app, device_id, ["Next", "Confirm", "Agree", "Confirm", "Continue", "Continue"], adb_path)
            app_sleep(app, 2.0, device_id)
        else:
            app.log(f"[{device_id}] ⚠️ Confirmation dialog not found, returning UI_UNKNOWN", level="WARN")
            return "UI_UNKNOWN"

        if not app.is_running or device_id not in app.active_running_devices: return "TERMINATED"
            
        if not adb_click_text(app, device_id, ["Next", "Confirm", "Continue", "Confirm"], adb_path):
            adb_click(app, device_id, COORD_BTN_XAC_NHAN, adb_path)
            
        # Fast scanning for next page type
        app.update_device_ui(device_id, status_text="🔄 Resolving screen...", text_color="#6366f1")
        start_time = time.time()
        page_detected = None
        
        while time.time() - start_time < 45:
            if not app.is_running or device_id not in app.active_running_devices: return "TERMINATED"
            xml_data = get_ui_xml(device_id, adb_path, app).lower()
            if any(t in xml_data for t in ["banned", "locked", "blocked", "invalid", "suspended", "locked", "temporarily locked", "invalid", "disabled", "banned"]):
                page_detected = "banned"
                break
            if any(t in xml_data for t in ["captcha", "puzzle", "puzzle", "slide"]):
                page_detected = "captcha"
                break
            if any(t in xml_data for t in ["verification", "enter code", "send message", "please enter", "enter your verification", "verification code", "enter code", "send message", "compose"]):
                page_detected = "otp"
                break
            app_sleep(app, 0.5, device_id)
            
        if not page_detected:
            xml_data = get_ui_xml(device_id, adb_path, app)
            texts = re.findall(r'(?:text|content-desc)="([^"]*)"', xml_data)
            app.log(f"[{device_id}] Unknown screen texts: {texts[:15]}", level="WARN")
            app.log(f"[{device_id}] Screen state timeout, retrying Next/check instead of blacklisting...", level="WARN")
            return "TIMEOUT"
            
        if page_detected == "banned":
            app.log(f"[{device_id}] ⚠️ Number {phone} is Banned or Invalid Zalo Account.")
            return False
            
        if page_detected == "captcha":
            app.update_device_ui(device_id, status_text="🧩 Solving Captcha...", text_color="#e11d48")
            attempt = 0
            max_captcha_attempts = 3
            captcha_solved = False
            
            while app.is_running and device_id in app.active_running_devices and attempt < max_captcha_attempts:
                # Kiểm tra trạng thái trang trước khi giải
                xml_data = get_ui_xml(device_id, adb_path, app)
                current_page = detect_page(xml_data)
                
                if current_page == "OTP":
                    captcha_solved = True
                    break
                elif current_page == "PROCESSING":
                    page_after = wait_page_after_captcha(app, device_id, adb_path, timeout=30)
                    if page_after == "OTP":
                        captcha_solved = True
                        break
                    continue
                elif current_page != "CAPTCHA":
                    app.log(f"[{device_id}] Không ở captcha, quét lại màn hình trước khi fail. page={current_page}")
                    page_after = wait_page_after_captcha(app, device_id, adb_path, timeout=10)
                    if page_after == "OTP":
                        captcha_solved = True
                        break
                    return "TIMEOUT"
                    
                attempt += 1
                t0 = time.time()
                app.log(f"[{device_id}] Solver trial {attempt}...")
                if solve_zalo_captcha(app, device_id, screen_w, screen_h, adb_path, offset_captcha, session):
                    app.log(f"[{device_id}] Captcha solve took {time.time() - t0:.1f}s")
                    
                    page_after = wait_page_after_captcha(app, device_id, adb_path, timeout=60)

                    if page_after == "OTP":
                        app.update_device_ui(device_id, status_text="✉️ Waiting SMS OTP...", text_color="#10b981")
                        captcha_solved = True
                        break

                    if page_after == "PROCESSING":
                        continue

                    if page_after == "CAPTCHA":
                        app.log(f"[{device_id}] Vẫn ở captcha, thử lại.")
                        continue

                    if page_after == "BANNED":
                        return False

                    if page_after in ["TIMEOUT", "STOPPED", "UNKNOWN"]:
                        app.log(f"[{device_id}] Timeout/Lỗi màn sau captcha: {page_after}")
                        return "UI_UNKNOWN"
                else:
                    app.log(f"[{device_id}] solve_zalo_captcha returned False. Took {time.time() - t0:.1f}s")
                    if not app.is_running or device_id not in app.active_running_devices:
                        break
                    
                    page_now, xml_now, texts_now = scan_page_with_log(
                        app, device_id, adb_path, note="during captcha"
                    )

                    if page_now == "OTP":
                        app.log(f"[{device_id}] Đã thấy màn OTP trong lúc trạng thái còn captcha, nhảy sang bước OTP.")
                        app.update_device_ui(device_id, status_text="✉️ Waiting SMS OTP...", text_color="#10b981")
                        captcha_solved = True
                        break

                    if page_now == "PROCESSING":
                        app.log(f"[{device_id}] Zalo đang xử lý sau captcha, tiếp tục chờ...")
                        
                        continue

                    if page_now == "CAPTCHA":
                        app.log(f"[{device_id}] Vẫn ở captcha, chờ thêm trước khi retry.")
                        
                        adb_click(app, device_id, f"{int(screen_w*0.82)} {int(screen_h*0.35)}", adb_path)
                        app_sleep(app, 1.5, device_id)
                        continue

                    if page_now == "BANNED":
                        return False
                    
                    # If page is something else, try to refresh captcha just in case
                    adb_click(app, device_id, f"{int(screen_w*0.82)} {int(screen_h*0.35)}", adb_path)
                    app_sleep(app, 1.5, device_id)

            if not captcha_solved:
                app.log(f"[{device_id}] ❌ Thất bại Captcha sau {max_captcha_attempts} lần.")
                return "TIMEOUT"

        # Verify OTP screen explicitly
        page_now, xml_data, texts_now = scan_page_with_log(
            app, device_id, adb_path, note="before OTP verify"
        )
        if page_now != "OTP":
            if page_now == "BANNED":
                app.log(f"[{device_id}] Bị khóa sau khi giải Captcha. Page={page_now}")
                return False
            else:
                app.log(f"[{device_id}] Không phải màn OTP, chưa gửi SMS. Page={page_now}, Texts={texts_now}", level="WARN")
                return "UI_UNKNOWN"

        # OTP Validation screen
        app.update_device_ui(device_id, status_text="✉️ Waiting SMS OTP...", text_color="#10b981")
        xml_data = get_ui_xml(device_id, adb_path, app)
        target_recipient = None
        texts_on_screen = re.findall(r'text="([^"]*)"', xml_data)
        for text_node in texts_on_screen:
            # Match 4-5 digit shortcodes (6xxx, 7xxx, 8xxx, 9xxx, etc.)
            match = re.search(r'\b(\d{4,5})\b', text_node)
            if match:
                code = match.group(1)
                # Filter out unlikely shortcodes (year-like numbers, phone fragments)
                if len(code) <= 5 and not code.startswith('0') and int(code) >= 1000:
                    target_recipient = code
                    break
                
        # Send SMS Trigger
        if target_recipient:
            sms_content = "ZALO"
            for text_node in texts_on_screen:
                match_content = re.search(r'(?:soạn|send)\s+(.*?)\s+(?:gửi|to)\s+' + target_recipient, text_node, re.IGNORECASE)
                if not match_content:
                    match_content = re.search(r'(?:soạn|send)\s+(.*?)\s+(?:gửi|to)', text_node, re.IGNORECASE)
                if match_content:
                    parsed_content = match_content.group(1).strip()
                    parsed_content = re.sub(r'["\']', '', parsed_content).strip()
                    if parsed_content:
                        sms_content = parsed_content
                        app.log(f"[{device_id}] Dynamically parsed SMS content: '{sms_content}'")
                        break
            
            app.log(f"[{device_id}] Zalo requires MO SMS to shortcode {target_recipient} with content '{sms_content}'")
            ff_token = app.ff_token_entry.get().strip()
            ff_service = app.ff_service_entry.get().strip()
            ff_country = app.ff_country_entry.get().strip()

            app.log(f"[{device_id}] Gửi SMS Firefox tới {target_recipient}: {sms_content}")
            res = firefox_send_sms(session, ff_token, ff_ctx["pkey"], target_recipient, sms_content)

            if res and res.startswith("1|"):
                app.log(f"[{device_id}] Chờ biên lai gửi SMS...")
                receipt_status = firefox_wait_sms_receipt(app, device_id, session, ff_token, ff_ctx["pkey"], timeout=60)
                if receipt_status:
                    app.log(f"[{device_id}] Đang xin cấp lại pkey mới cho số {phone}...")
                    new_pkey, reuse_res = firefox_get_phone(session, ff_token, ff_service, ff_country, mobile=firefox_mobile_param(phone))
                    if new_pkey:
                        ff_ctx["pkey"] = new_pkey
                        app.log(f"[{device_id}] Đã lấy pkey mới thành công. Bắt đầu chờ mã OTP Zalo...")
                        otp = firefox_wait_otp(app, device_id, session, ff_token, ff_ctx["pkey"], timeout=120)
                    else:
                        app.log(f"[{device_id}] Không thể lấy lại pkey mới cho số {phone}. Response: {reuse_res}", level="ERROR")
                        return False
                else:
                    app.log(f"[{device_id}] Không nhận được biên lai gửi SMS.", level="ERROR")
                    return False
            else:
                app.log(f"[{device_id}] Gửi SMS Firefox lỗi: {translate_firefox_error(res, act='sendCode')}", level="ERROR")
                return False
        else:
            app.log(f"[{device_id}] Zalo auto-sent OTP SMS verification code.")
            ff_token = app.ff_token_entry.get().strip()
            otp = firefox_wait_otp(app, device_id, session, ff_token, ff_ctx["pkey"], timeout=120)
            
        # Xử lý các loại lỗi khác nhau
        if otp is None:
            app.log(f"[{device_id}] ❌ Thất bại nhận OTP từ Firefox - Bỏ qua.")
            return False
            
        with app.completed_lock:
            app.otp_received_count += 1
        ff_ctx["otp_received"] = True
        app.update_stats_ui()
            
        # Type OTP
        if otp != "AUTO_VERIFIED":
            app.update_device_ui(device_id, status_text="🔄 Submitting OTP...", text_color="#10b981")
            
            # Format and normalize OTP (must be exactly 6 digits)
            otp = str(otp).strip()
            otp = ''.join(filter(str.isdigit, otp))
            if len(otp) < 6:
                otp = otp.zfill(6)
            elif len(otp) > 6:
                otp = otp[:6]
                
            app.log(f"[{device_id}] Typing normalized OTP: {otp}")
            
            if not adb_focus_input(app, device_id, adb_path):
                adb_click(app, device_id, f"{int(screen_w/2)} {int(screen_h*0.35)}", adb_path)
                
            # Clear any existing text first by sending keyevent 67 (delete) 8 times
            for _ in range(8):
                run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "67"], timeout=3, device_id=device_id, app=app)
            
            # Type OTP digit-by-digit using native keyevents to bypass Telex/IME issues
            adb_type_digits(app, device_id, otp, adb_path)
            
            # Wait for screen to respond to OTP input before clicking Next
            app.log(f"[{device_id}] Waiting for system to process OTP input...")
            
            # Verify OTP screen state BEFORE proceeding
            if not check_text_exists(device_id, ["To provide", "Profile name", "Your name", "Name", "To provide", "Your name", "Enter name"], adb_path):
                app.log(f"[{device_id}] Next screen not detected, clicking Next button...")
                adb_click_text(app, device_id, ["Next", "Continue"], adb_path)
                
            # MUST wait for next screen to appear
            start_wait = time.time()
            success_advanced = False
            target_texts = ["To provide", "Profile name", "Your name", "Name", "To provide", "Your name", "Enter name"]
            fail_texts = ["account already exists", "account locked", "was linked to this account", "account exists"]
            import unicodedata
            app.log(f"[{device_id}] Waiting for next screen after OTP (Timeout: 30s)")
            while time.time() - start_wait < 30:
                if not app.is_running or device_id not in app.active_running_devices: return False
                xml_data = get_ui_xml(device_id, adb_path, app)
                if xml_data:
                    visible_texts = ' '.join(re.findall(r'(?:text|content-desc)="([^"]*)"', xml_data)).lower()
                    visible_texts = unicodedata.normalize('NFC', visible_texts)
                    
                    if any(t in visible_texts for t in fail_texts):
                        app.log(f"[{device_id}] ❌ Phát hiện lỗi: Tài khoản đã tồn tại hoặc Bị khóa. Ghi số {phone} vào file.")
                        try:
                            with open("sdt_da_ton_tai.txt", "a", encoding="utf-8") as f:
                                f.write(f"{phone}\n")
                        except Exception as e:
                            app.log(f"[{device_id}] Lỗi ghi file sdt: {e}", level="ERROR")
                        return False
                        
                    if any(t in visible_texts for t in [unicodedata.normalize('NFC', t).lower() for t in target_texts]):
                        app_sleep(app, 0.2, device_id)
                        success_advanced = True
                        break
                app_sleep(app, POLL_INTERVAL, device_id)
                
            if not success_advanced:
                app.log(f"[{device_id}] ❌ Failed to advance past OTP screen")
                return False
        
        
        # Policy accept UI
        if check_text_exists(device_id, ["To provide", "To provide"], adb_path):
            if not adb_click_text(app, device_id, "To provide", adb_path, align="outside_left"):
                adb_click_text(app, device_id, "To provide", adb_path, align="outside_left")
            run_adb([adb_path, "-s", device_id, "shell", "input", "swipe"] + SWIPE_SCROLL_UP.split(), timeout=8, device_id=device_id, app=app)
            
            if not adb_click_text(app, device_id, ["Agree", "Next", "Agree", "Continue"], adb_path):
                btn_dongy_x = int(screen_w * 0.5)
                btn_dongy_y = int(screen_h * 0.92)
                run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_dongy_x), str(btn_dongy_y)], timeout=5, device_id=device_id, app=app)
            
            if not wait_for_text(app, device_id, ["Profile name", "Your name", "Name", "Your name", "Enter name"], adb_path, timeout=15):
                app.log(f"[{device_id}] Profile name screen not found")
                return "UI_UNKNOWN"
 
        # Enter Zalo Profile Name
        app.update_device_ui(device_id, status_text="🔄 Setting Profile Name...", text_color="#10b981")
        adb_focus_input(app, device_id, adb_path)
        name_list = [
            "nguyen van minh", "tran thi lan", "le hoang nam", "pham thu huong",
            "vu duc hai", "dang ngoc anh", "bui xuan truong", "hoang thi mai",
            "nguyen hai dang", "tran quoc bao", "le thi hoa", "pham minh tuan",
            "nguyen van a", "nguyen quoc b", "nguyen van c", "nguyen van d",
            "nguyen thi huong", "tran van hieu", "le minh chau", "pham quang huy",
            "vu thi ngoc", "dang thanh tung", "bui thi kim", "hoang gia bao",
            "nguyen dinh tuan", "tran minh triet", "le thi mai anh", "pham hoang yen",
            "vu hoang long", "dang dinh khoi", "bui thanh hang", "hoang quoc anh",
            "nguyen thi thuy", "tran duy khanh", "le hoang bach", "pham ngoc mai",
            "vu thi thu trang", "dang quang huy", "bui minh hieu", "hoang thi tuyet",
            "nguyen xuan phuc", "tran anh dung", "le thi thanh tam", "pham duc anh",
            "nguyen phuong thao", "tran thanh hai", "le minh tu", "pham quoc khanh",
            "vu hoang nam", "dang thuy chi", "bui duc anh", "hoang ngoc lan",
            "nguyen van quyet", "tran thi thu", "le quang vinh", "pham thi diep",
            "vu ngoc diep", "dang minh quan", "bui thu thao", "hoang van dong",
            "nguyen thanh cong", "tran ngoc trinh", "le khac huy", "pham thi thanh"
        ]
        random_name = random.choice(name_list)
        adb_type(app, device_id, random_name, adb_path)
        
        # Verify name was entered before proceeding
        app.log(f"[{device_id}] Profile name entered: {random_name}. Proceeding...")
        
        clicked_next = adb_click_text(app, device_id, ["Continue", "Next"], adb_path)
        
        if not clicked_next:
            app.log(f"[{device_id}] ⚠️ Next button after name not found, trying keyboard Enter/fallback tap...", level="WARN")
            run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "66"], timeout=5, device_id=device_id, app=app)
            app_sleep(app, 1, device_id)
            
            next_x = int(screen_w * 0.5)
            next_y = int(screen_h * 0.88)
            adb_click(app, device_id, f"{next_x} {next_y}", adb_path)
        
        # MUST verify birthday/gender screen appears
        if not wait_for_text(app, device_id, ["Birthday", "Birthday", "Gender", "Gender"], adb_path, timeout=20):
            app.log(f"[{device_id}] ❌ Birthday screen did not appear after name input")
            return "UI_UNKNOWN"
        
        # Scroll values for age fields
        app.update_device_ui(device_id, status_text="🔄 Setting Birthday/Gender...", text_color="#22d3ee")
        if not app.is_running or device_id not in app.active_running_devices: return "TERMINATED"
        
        # Gender selection logic
        gender = random.choice(["Male", "Female"])
        app.log(f"[{device_id}] Selecting gender: {gender}")
        
        if not wait_for_text(app, device_id, ["Male", "Female", "Nam", "Female"], adb_path, timeout=UI_DUMP_TIMEOUT):
            app.log(f"[{device_id}] Không tìm thấy màn hình chọn Giới tính", level="WARN")
            return "UI_UNKNOWN"
        
        if adb_click_text(app, device_id, ["Birthday", "Birthday", "199"], adb_path): 
            duration_year = random.randint(8000, 12000)
            cx1 = int(screen_w * (728 / 1080))
            cy1 = int(screen_h * (1285 / 2220))
            run_adb([adb_path, "-s", device_id, "shell", "input", "swipe", str(cx1), str(cy1), str(cx1), str(cy1), str(duration_year)], timeout=25, device_id=device_id, app=app)
            
            duration_month = random.randint(1000, 5000)
            cx2 = int(screen_w * (550 / 1080))
            cy2 = int(screen_h * (1285 / 2220))
            run_adb([adb_path, "-s", device_id, "shell", "input", "swipe", str(cx2), str(cy2), str(cx2), str(cy2), str(duration_month)], timeout=10, device_id=device_id, app=app)
            
            duration_day = random.randint(1000, 5000)
            cx3 = int(screen_w * (362 / 1080))
            cy3 = int(screen_h * (1285 / 2220))
            run_adb([adb_path, "-s", device_id, "shell", "input", "swipe", str(cx3), str(cy3), str(cx3), str(cy3), str(duration_day)], timeout=10, device_id=device_id, app=app)
            
            adb_click_text(app, device_id, ["Select", "Select"], adb_path, exact_match=True)
 
        if adb_click_text(app, device_id, ["Gender", "Gender"], adb_path):
            gender_select = random.choice([["Male", "Nam"], ["Female", "Female"]])
            adb_click_text(app, device_id, gender_select, adb_path)
            
        adb_click_text(app, device_id, ["Next", "Continue"], adb_path)
        
        app.log(f"[{device_id}] Waiting for Avatar/Cover screen or failure toast...")
        start_wait = time.time()
        screen_success = False
        while time.time() - start_wait < 30:
            if not app.is_running or device_id not in app.active_running_devices: return False
            
            if check_text_exists(device_id, ["Account creation failed", "failed. (2008)", "Account creation error", "unsuccessful"], adb_path, app):
                app.log(f"[{device_id}] ❌ Lỗi tạo tài khoản (Account creation failed).", level="ERROR")
                return False
                
            if check_text_exists(device_id, ["Next", "Skip", "Later", "Continue", "Skip", "Later"], adb_path, app):
                screen_success = True
                break
                
            app_sleep(app, 0.5, device_id)
            
        if not screen_success:
            app.log(f"[{device_id}] ❌ Skip screen timeout")
            return "UI_UNKNOWN"
        
        # Skip sub-flows
        adb_click_text(app, device_id, ["Next", "Continue"], adb_path)
        app_sleep(app, 0.2, device_id)
        
        if adb_click_text(app, device_id, ["Skip", "Skip"], adb_path):
            adb_click_text(app, device_id, ["Skip", "Skip"], adb_path, match_index=1)
            
        # Handle Update avatar screen
        if wait_for_text(app, device_id, ["Update your avatar", "Update", "Skip", "Update avatar"], adb_path, timeout=5):
            app.log(f"[{device_id}] Update avatar screen detected, clicking Skip...")
            if not adb_click_text(app, device_id, ["Skip", "Skip"], adb_path):
                skip_x = int(screen_w * 0.5)
                skip_y = int(screen_h * (2020 / 2220))
                run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(skip_x), str(skip_y)], timeout=5, device_id=device_id, app=app)
            
        if not wait_for_text(app, device_id, ["Later", "Stay connected", "Search", "Later", "Search", "Find friends"], adb_path, timeout=10):
            app.log(f"[{device_id}] Later/Search screen timeout")
        adb_click_text(app, device_id, ["Later", "Later"], adb_path)
        app_sleep(app, 1, device_id)

        run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "4"], timeout=5, device_id=device_id, app=app)
        
        # ------------------ MINI APP CAMPAIGN ------------------
        app.update_device_ui(device_id, status_text="🎮 Opening 7up App...", text_color="#f59e0b")
        mini_app_success = False

        already_on_7up = quick_scan_7up_screen(app, device_id, adb_path, timeout=5)

        if already_on_7up:
            app.log(f"[{device_id}] 7UP page already loaded, skipping search flow.")
        else:
            if wait_for_text(app, device_id, ["Search", "Find friends", "Search", "Find friends", "Find friends"], adb_path, timeout=30):
                if adb_click_text(app, device_id, ["Search", "Find friends", "Search", "Find friends", "Find friends"], adb_path):
                    app_sleep(app, 1, device_id)
                    adb_type(app, device_id, "7up", adb_path)
                
                app.log(f"[{device_id}] Checking View results before searching 7up app...")
                if check_text_exists(device_id, ["View results", "View results"], adb_path, app):
                    if not adb_click_text(app, device_id, ["View results", "View results"], adb_path):
                        view_x = int(screen_w * 0.5)
                        view_y = int(screen_h * (720 / 2220))
                        app.log(f"[{device_id}] View results text found but click failed, tapping fallback ({view_x},{view_y})")
                        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(view_x), str(view_y)], timeout=5, device_id=device_id, app=app)
                
                # Scroll through search results to find 7up app
                search_attempts = 0
                while search_attempts < 5:
                    search_attempts += 1
                    app.log(f"[{device_id}] Looking for 7up app (attempt {search_attempts}/5)...")
                    
                    if wait_for_text(app, device_id, ["Khui 7up", "Open 7up", "Hunt Gifts Now"], adb_path, timeout=8):
                        app.log(f"[{device_id}] Found 7up app on attempt {search_attempts}")
                        break
                    
                    if check_text_exists(device_id, ["View results", "View results"], adb_path, app):
                        app.log(f"[{device_id}] Still on View results screen, clicking it again...")
                        adb_click_text(app, device_id, ["View results", "View results"], adb_path)
                        continue
                    
                    # Scroll down to find more results
                    scroll_x = int(screen_w * 0.5)
                    scroll_top = int(screen_h * 0.3)
                    scroll_bot = int(screen_h * 0.8)
                    run_adb([adb_path, "-s", device_id, "shell", "input", "swipe", str(scroll_x), str(scroll_top), str(scroll_x), str(scroll_bot), "400"], timeout=5, device_id=device_id, app=app)
                    app_sleep(app, 1, device_id)
                
                # Click 7up app link
                for click_attempt in range(3):
                    if not app.is_running or device_id not in app.active_running_devices:
                        return False
                    if adb_click_text(app, device_id, ["Khui 7up", "Open 7up", "Hunt Gifts Now"], adb_path, silent=(click_attempt > 0)):
                        app_sleep(app, 1.5, device_id)
                        break
                    app_sleep(app, 0.5, device_id)
                
        # Proceed with checkbox clicks regardless of how we got to 7UP
        for welcome_retry in range(3):
            if not app.is_running or device_id not in app.active_running_devices: return False
            cb_welcome_x = int(screen_w * 0.12)
            cb_welcome_y = int(screen_h * 0.68)
            run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(cb_welcome_x), str(cb_welcome_y)], timeout=5, device_id=device_id, app=app)
            app_sleep(app, 1, device_id)
 
            btn_cont_x = int(screen_w * (553 / 1080))
            btn_cont_y = int(screen_h * (1724 / 2220))
            run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_cont_x), str(btn_cont_y)], timeout=5, device_id=device_id, app=app)
            app_sleep(app, 1, device_id)

            # Dynamic wait for welcome screen to close (up to 5s)
            start_welcome = time.time()
            while time.time() - start_welcome < 7:
                if not app.is_running or device_id not in app.active_running_devices: return False
                if not check_text_exists(device_id, ["Welcome to Zalo", "To get started", "Terms of Use", "Welcome", "To get started", "Terms"], adb_path):
                    break
            
            if not check_text_exists(device_id, ["Welcome to Zalo", "To get started", "Terms of Use", "Welcome", "To get started", "Terms"], adb_path):
                break
        
            
        # Click Tab Ca Nhan - VERIFY screen is ready BEFORE clicking
        tab_x = int(screen_w * 0.875)
        tab_y = int(screen_h * 0.95)
        app_sleep(app, 1, device_id)
 
        # Scan screen before clicking tab
        xml_before = get_ui_xml(device_id, adb_path, app)
        if not xml_before or "<node" not in xml_before:
            app.log(f"[{device_id}] Screen not fully loaded, waiting...")
            app_sleep(app, 1, device_id)
        
        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(tab_x), str(tab_y)], timeout=5, device_id=device_id, app=app)
        
        # VERIFY tab loaded by checking XML changes
        start_tab = time.time()
        tab_loaded = False
        while time.time() - start_tab < 4:
            if not app.is_running or device_id not in app.active_running_devices: 
                return False
            xml = get_ui_xml(device_id, adb_path, app)
            if xml and "<node" in xml:
                tab_loaded = True
                break
            app_sleep(app, 1, device_id)
        
        if not tab_loaded:
            app.log(f"[{device_id}] ⚠️ Tab content did not load, retrying...")
            app_sleep(app, 2.0, device_id)
        
        cb_x = int(screen_w * (125 / 1080))
        cb1_y = int(screen_h * (1576 / 2220))
        cb2_y = int(screen_h * (1815 / 2220))
        cb3_y = int(screen_h * (1960 / 2220))
        
        # Wait until mini app terms/checkbox area is really loaded
        app.log(f"[{device_id}] Waiting for 7up checkbox area to load...")
        
        checkbox_ready = False
        for i in range(3):  
            if not app.is_running or device_id not in app.active_running_devices:
                return False
        
            xml = get_ui_xml(device_id, adb_path, app)
            texts = xml.lower()
        
            if (
                "i have read" in texts
                or "i agree" in texts
                or "terms" in texts
                or "follow oa" in texts
                or "followed successfully" in texts
            ):
                checkbox_ready = True
                break
        
            app_sleep(app, 1, device_id)
        
        if not checkbox_ready:
            app.log(f"[{device_id}] 7up checkbox area not ready, waiting extra 3s...", level="WARN")
            
        app_sleep(app, 2.5, device_id) # Đợi toast 'Followed successfully' biến mất để khỏi che nút tick
        
        # Tap checkbox 1 only after visible area loaded
        app.log(f"[{device_id}] Tapping first terms checkbox...")
        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(cb_x), str(cb1_y)], timeout=5, device_id=device_id, app=app)
        app_sleep(app, 1, device_id)
        
        # Bỏ qua check_text vì popup Zalo Mini App dạng Webview không lấy được text qua uiautomator
        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(cb_x), str(cb2_y)], timeout=5, device_id=device_id, app=app)
        app.log(f"[{device_id}] Tapped checkbox 2, waiting for phone permission popup...")
        app_sleep(app, 1, device_id) # Chờ 3 giây cho popup hiển thị hoàn toàn
        
        btn_phone_x = 952
        btn_phone_y = 1833
        app.log(f"[{device_id}] BEFORE TAP phone ({btn_phone_x},{btn_phone_y})")
        res = run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_phone_x), str(btn_phone_y)], timeout=5, device_id=device_id, app=app)
        app.log(f"[{device_id}] AFTER TAP phone rc={getattr(res, 'returncode', None)} err={getattr(res, 'stderr', '')}")
        app_sleep(app, 2, device_id)
        
        btn_allow_x = 780
        btn_allow_y = 2053
        app.log(f"[{device_id}] BEFORE TAP allow ({btn_allow_x},{btn_allow_y})")
        res = run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_allow_x), str(btn_allow_y)], timeout=5, device_id=device_id, app=app)
        app.log(f"[{device_id}] AFTER TAP allow rc={getattr(res, 'returncode', None)} err={getattr(res, 'stderr', '')}")
        app_sleep(app, 2, device_id)
        
        # VERIFY screen before clicking checkbox 3
        app.log(f"[{device_id}] Verified. Clicking final checkbox...")
        app_sleep(app, 1, device_id)
        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(cb_x), str(cb3_y)], timeout=5, device_id=device_id, app=app)
        app_sleep(app, 6, device_id)
 
        # Clicking Follow button
        app.log(f"[{device_id}] Clicking Follow...")
        btn_follow_x = int(screen_w * (785 / 1080))
        btn_follow_y = int(screen_h * (2075 / 2220))
        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_follow_x), str(btn_follow_y)], timeout=5, device_id=device_id, app=app)
        app_sleep(app, 2, device_id)
        
        # VERIFY screen before scrolling and clicking participate
        app.log(f"[{device_id}] Scrolling to participate button...")
        app_sleep(app, 2, device_id)
        swipe_x = int(screen_w * (500 / 1080))
        swipe_y1 = int(screen_h * (1800 / 2220))
        swipe_y2 = int(screen_h * (300 / 2220))
        run_adb([adb_path, "-s", device_id, "shell", "input", "swipe", str(swipe_x), str(swipe_y1), str(swipe_x), str(swipe_y2), "300"], timeout=5, device_id=device_id, app=app)
        app_sleep(app, 5, device_id)
        
        # VERIFY participate button is visible before clicking
        btn_thamgia_x = int(screen_w * (588 / 1080))
        btn_thamgia_y = int(screen_h * (1755 / 2220))
        
        app.log(f"[{device_id}] Scanning for participate button state...")
        run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_thamgia_x), str(btn_thamgia_y)], timeout=5, device_id=device_id, app=app)
        app_sleep(app, 5, device_id)
 
        # Dynamic wait for participation response (up to 7s)
        mini_app_success = False
        start_thamgia = time.time()
        while time.time() - start_thamgia < 25:
            if not app.is_running or device_id not in app.active_running_devices: return "TERMINATED"
            if check_text_exists(device_id, ["tham gia", "success", "success", "congratulations", "gift", "confirm", "close"], adb_path):
                app.log(f"[{device_id}] 7up mini app joined successfully!", level="SUCCESS")
                mini_app_success = True
                break
            app_sleep(app, 5, device_id)
        
        # Minimize Zalo to background (press HOME key instead of force-stop)
        run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "3"], timeout=5, device_id=device_id, app=app)
        
        if not mini_app_success:
            app.log(f"[{device_id}] ⚠️ Mini App 7up failed — aborting")
            return False
        
        # Open Token Extractor App
        app.update_device_ui(device_id, status_text="🔄 Saving Zalo Token...", text_color="#10b981")

        # Setup adb reverse to allow direct webhook pushing bypassing PC firewall
        run_adb(f'"{adb_path}" -s {device_id} reverse tcp:5000 tcp:5000', timeout=5, device_id=device_id, app=app)
        
        # Force stop Zalo first via ADB to guarantee Zalo starts fresh even if Autopee app lacks root
        run_adb(f'"{adb_path}" -s {device_id} shell am force-stop com.zing.zalo', timeout=5, device_id=device_id, app=app)
       
         # Bật Tailscale
        
        
        # Open com.example.hack
        token_pkg = "com.example.hack"
        run_adb(f'"{adb_path}" -s {device_id} shell am start -n com.example.hack/.MainActivity', timeout=8, device_id=device_id, app=app)
        
        # VERIFY Autopee app loaded before clicking trigger
        if not wait_for_text(app, device_id, ["TRIGGER", "adb shell am broadcast", "com.autopee"], adb_path, timeout=5):
            app.log(f"[{device_id}] ⚠️ Autopee app did not load in time")
            return False
        
        app.log(f"[{device_id}] Autopee app loaded. Clicking trigger...")
        
        # Click TRIGGER dynamically by text first, fallback to coordinates if not found
        if not adb_click_text(app, device_id, ["TRIGGER"], adb_path):
            app.log(f"[{device_id}] ⚠️ Click trigger via text failed, using coordinate fallback...")
            btn_trigger_x = int(screen_w * (573 / 1080))
            btn_trigger_y = int(screen_h * (1026 / 2220))
            run_adb([adb_path, "-s", device_id, "shell", "input", "tap", str(btn_trigger_x), str(btn_trigger_y)], timeout=5, device_id=device_id, app=app)
        
        app.trigger_success_count += 1
        ff_ctx["trigger_success"] = True
        app.update_stats_ui()
        app_sleep(app, 1, device_id)

        start_token = time.time()
        token_pulled = False
        while time.time() - start_token < 24:
            if not app.is_running or device_id not in app.active_running_devices: 
                return False
            
            # Check primary: /sdcard/Download/
            token_file_device = None
            check_file = run_adb([adb_path, "-s", device_id, "shell", "ls", "/sdcard/Download/autopee_token.json"], timeout=5, device_id=device_id, app=app)
            if check_file and "autopee_token.json" in check_file.stdout:
                token_file_device = "/sdcard/Download/autopee_token.json"
            else:
                # Fallback: /data/local/tmp/ (ZaloHooker ghi vao day neu /sdcard that bai)
                check_tmp = run_adb([adb_path, "-s", device_id, "shell", "ls", "/data/local/tmp/autopee_token.json"], timeout=5, device_id=device_id, app=app)
                if check_tmp and "autopee_token.json" in check_tmp.stdout:
                    token_file_device = "/data/local/tmp/autopee_token.json"

            if token_file_device:
                app.log(f"[{device_id}] Token file found at {token_file_device}! Pulling to server...")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:19]
                local_filename = f"token_pulled_{device_id}_{timestamp}.json"
                local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), local_filename)

                run_adb([adb_path, "-s", device_id, "pull", token_file_device, local_path], timeout=15, device_id=device_id, app=app)
                run_adb([adb_path, "-s", device_id, "shell", "rm", token_file_device], timeout=8, device_id=device_id, app=app)

                app.log(f"[{device_id}] Token successfully pulled: {local_filename}")

                try:
                    with open(local_path, "r", encoding="utf-8") as f:
                        import json
                        token_data = json.load(f)
                    import requests
                    requests.post("http://127.0.0.1:5000/token", json=token_data, timeout=3)
                    app.log(f"[{device_id}] Token forwarded to local server.py successfully.")
                except Exception as e:
                    app.log(f"[{device_id}] Failed to forward token to server.py: {e}")

                token_pulled = True
                break

            if check_text_exists(device_id, ["CAPTURED", "captured", "Saved successfully"], adb_path):
                app.log(f"[{device_id}] Token capture detected on Autopee screen, waiting for file...")
            app_sleep(app, 0.8, device_id)
        
        # Back to Xtoolz backup
        app.update_device_ui(device_id, status_text="🔄 Backing up Profile...", text_color="#38bdf8")
        run_adb(f'"{adb_path}" -s {device_id} shell am force-stop {token_pkg}', timeout=5, device_id=device_id, app=app)
        # Locate package dynamically
        XTOOLZ_PACKAGE = "app.pixel.xtoolz"
        launch_res = run_adb(
            [adb_path, "-s", device_id, "shell", "monkey", "-p", XTOOLZ_PACKAGE, "-c", "android.intent.category.LAUNCHER", "1"],
            timeout=8,
            device_id=device_id,
            app=app
        )

        if not launch_res or launch_res.returncode != 0:
            app.log(f"[{device_id}] Không mở được Xtoolz package={XTOOLZ_PACKAGE}", level="ERROR")
            return False
        
        # VERIFY Xtoolz loaded before performing backup
        app.log(f"[{device_id}] Waiting for Xtoolz app to load...")
        app_sleep(app, 0.5, device_id)
        
        xml_xtoolz = get_ui_xml(device_id, adb_path, app)
        if not xml_xtoolz or "<node" not in xml_xtoolz:
            app.log(f"[{device_id}] ⚠️ Xtoolz screen not ready, waiting more...")
            app_sleep(app, 0.5, device_id)
        
        # Perform XToolz backup flows
        app.log(f"[{device_id}] Initiating Xtoolz backup process...")
        cx1 = int(screen_w * (683 / 1080))
        cy1 = int(screen_h * (2000 / 2220))
        adb_click(app, device_id, f"{cx1} {cy1}", adb_path)
        
        cx2 = int(screen_w * (776 / 1080))
        cy2 = int(screen_h * (164 / 2220))
        adb_click(app, device_id, f"{cx2} {cy2}", adb_path)
        
        # VERIFY backup item appears before proceeding
        start_backup = time.time()
        backup_found = False
        while time.time() - start_backup < 6:
            if not app.is_running or device_id not in app.active_running_devices: 
                return False
            if check_text_exists(device_id, ["zalo", "com.zing.zalo"], adb_path):
                backup_found = True
                app.log(f"[{device_id}] Zalo backup item detected")
                break
        
        if not backup_found:
            app.log(f"[{device_id}] ⚠️ Backup item not found, proceeding anyway...")
        
        cx3 = int(screen_w * (583 / 1080))
        cy3 = int(screen_h * (338 / 2220))
        app.log(f"[{device_id}] Long-pressing backup item to rename...")
        run_adb([adb_path, "-s", device_id, "shell", "input", "swipe", str(cx3), str(cy3), str(cx3), str(cy3), "3000"], timeout=8, device_id=device_id, app=app)
        
        # VERIFY rename menu appeared before clicking
        app.log(f"[{device_id}] Scanning for rename menu...")
        if not adb_click_text(app, device_id, ["Rename", "Change name"], adb_path):
            app.log(f"[{device_id}] ⚠️ Rename option not found, trying alternative...")
            app_sleep(app, 0.1, device_id)
            adb_click_text(app, device_id, ["Rename", "Change name"], adb_path)
        
        
        # Clear existing text
        for _ in range(15):
            run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "67"], timeout=3, device_id=device_id, app=app)
        
        # VERIFY input field is cleared before typing new name
        app.log(f"[{device_id}] Text field cleared. Entering phone number as backup name...")
        adb_type(app, device_id, f"{phone} tool", adb_path)
        
        # VERIFY new name entered before confirming
        if not adb_click_text(app, device_id, ["OK", "Agree", "Confirm", "Xong"], adb_path):
            app.log(f"[{device_id}] ⚠️ OK button not found, retrying...")
            adb_click_text(app, device_id, ["OK", "Agree", "Confirm", "Xong"], adb_path)
        
        
        run_adb([adb_path, "-s", device_id, "shell", "input", "keyevent", "4"], timeout=5, device_id=device_id, app=app)
        
        app.log(f"✅ [{device_id}] COMPLETED profile setup & backup successfully!")
        app.backup_success_count += 1
        ff_ctx["backup_success"] = True
        app.update_stats_ui()
        return True
            
    except Exception as e:
        app.log(f"[{device_id}] ❌ Process Error: {e}", level="ERROR")
        return False


if __name__ == "__main__":
    app = ZaloAutoUIApp()
    app.log(f"Running script file: {__file__}")
    app.mainloop()
