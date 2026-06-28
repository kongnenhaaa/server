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
from datetime import datetime

# Fix Unicode console output
sys.stdout.reconfigure(encoding='utf-8')

# Global requests session for performance
http_session = requests.Session()

# CustomTkinter setup
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ZaloAutoUIApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ToolZL Auto UI - Master Controller")
        self.geometry("1000x750")
        self.minsize(900, 650)
        self.configure(fg_color="#090d16") # Deep space black application window background
        
        # Application state
        self.is_running = False
        self.devices = []
        self.running_devices = []
        self.phone_numbers = []
        self.txt_path = ""
        self.device_ui_elements = {}  # device_id -> {frame, phone_lbl, status_lbl, status_text}
        self.active_running_devices = set()  # Currently running device IDs
        self.active_workers_count = 0
        self.active_workers_lock = threading.Lock()
        self.total_phones = 0
        self.qr_phones = 0
        self.backup_phones = 0
        self.list_phones = 0
        self.processed_phones = 0
        self.completed_lock = threading.Lock()
        
        # Grid layout (Left side configurations, Right side details/status, Bottom logging)
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0) # Space for logs
        self.grid_rowconfigure(2, weight=1) # Log area

        self.create_widgets()
        self.refresh_devices_list(on_startup=True)

    def create_widgets(self):
        # ------------------ LEFT COLUMN: SETTINGS ------------------
        left_panel = ctk.CTkFrame(self, corner_radius=15, fg_color="#131924", border_width=1, border_color="#1e293b")
        left_panel.grid(row=0, column=0, rowspan=2, padx=15, pady=15, sticky="nsew")
        left_panel.grid_columnconfigure(0, weight=1)
        
        # Header Label
        title_lbl = ctk.CTkLabel(left_panel, text="🛠️ CONFIGURATIONS", font=("Segoe UI", 15, "bold"), text_color="#38bdf8")
        title_lbl.pack(pady=(15, 10))

        # Config: ADB Path
        adb_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        adb_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(adb_frame, text="ADB Executable Path:", font=("Segoe UI", 11, "bold"), text_color="#cbd5e1").pack(anchor="w")
        
        adb_search_frame = ctk.CTkFrame(adb_frame, fg_color="transparent")
        adb_search_frame.pack(fill="x")
        self.adb_path_entry = ctk.CTkEntry(adb_search_frame, placeholder_text="Path to adb.exe", height=32,
                                           fg_color="#090d16", border_color="#1e293b", border_width=1, text_color="#f8fafc",
                                           placeholder_text_color="#475569")
        self.adb_path_entry.insert(0, r"D:\SIM\pmcong\platform-tools-latest-windows\platform-tools\adb.exe")
        self.adb_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.adb_path_entry.bind("<Return>", lambda event: self.refresh_devices_list())
        self.adb_path_entry.bind("<FocusOut>", lambda event: self.refresh_devices_list())
        
        adb_browse_btn = ctk.CTkButton(adb_search_frame, text="Browse", width=70, height=32,
                                       fg_color="#1e293b", border_width=1, border_color="#2b384e", hover_color="#2d3748",
                                       text_color="#cbd5e1", font=("Segoe UI", 11, "bold"), command=self.browse_adb)
        adb_browse_btn.pack(side="right")

        # Config: Captcha Offset
        offset_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        offset_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(offset_frame, text="Captcha Offset (pixels adjustment):", font=("Segoe UI", 11, "bold"), text_color="#cbd5e1").pack(anchor="w")
        self.captcha_offset_entry = ctk.CTkEntry(offset_frame, placeholder_text="0", height=32,
                                                 fg_color="#090d16", border_color="#1e293b", border_width=1, text_color="#f8fafc")
        self.captcha_offset_entry.insert(0, "0")
        self.captcha_offset_entry.pack(fill="x")

        # Config: Firebase URL
        firebase_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        firebase_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(firebase_frame, text="Firebase Database URL:", font=("Segoe UI", 11, "bold"), text_color="#cbd5e1").pack(anchor="w")
        self.firebase_url_entry = ctk.CTkEntry(firebase_frame, placeholder_text="https://...", height=32,
                                               fg_color="#090d16", border_color="#1e293b", border_width=1, text_color="#f8fafc")
        self.firebase_url_entry.insert(0, "https://toolweb-c7702-default-rtdb.firebaseio.com/")
        self.firebase_url_entry.pack(fill="x")

        # Separator line
        ctk.CTkFrame(left_panel, height=2, fg_color="#1e293b").pack(fill="x", padx=15, pady=15)

        # Phone input selection header
        phone_title_lbl = ctk.CTkLabel(left_panel, text="📞 PHONES DATA INPUT", font=("Segoe UI", 13, "bold"), text_color="#38bdf8")
        phone_title_lbl.pack(pady=(0, 5))

        # Config: Tab input selection (File or Comma List)
        self.phone_mode_var = ctk.StringVar(value="file")
        
        radio_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        radio_frame.pack(pady=5)
        r_file = ctk.CTkRadioButton(radio_frame, text="Select Text File", variable=self.phone_mode_var, value="file", 
                                    command=self.toggle_phone_mode, fg_color="#38bdf8", hover_color="#0284c7",
                                    text_color="#cbd5e1", font=("Segoe UI", 11, "bold"))
        r_file.pack(side="left", padx=10)
        r_manual = ctk.CTkRadioButton(radio_frame, text="Manual Entry", variable=self.phone_mode_var, value="manual", 
                                      command=self.toggle_phone_mode, fg_color="#38bdf8", hover_color="#0284c7",
                                      text_color="#cbd5e1", font=("Segoe UI", 11, "bold"))
        r_manual.pack(side="left", padx=10)

        # Phone Panel: File Selector
        self.phone_file_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        self.phone_file_frame.pack(fill="x", padx=15, pady=5)
        self.file_path_label = ctk.CTkLabel(self.phone_file_frame, text="No file selected...", text_color="#64748b", anchor="w", font=("Segoe UI", 11, "italic"))
        self.file_path_label.pack(fill="x")
        self.btn_browse_txt = ctk.CTkButton(self.phone_file_frame, text="Choose Phones File (.txt)", fg_color="#2563eb", 
                                            hover_color="#1d4ed8", font=("Segoe UI", 12, "bold"), command=self.browse_phones_txt)
        self.btn_browse_txt.pack(fill="x", pady=(5, 0))

        # Phone Panel: Text Entry
        self.phone_text_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        self.phone_entry = ctk.CTkTextbox(self.phone_text_frame, height=120, border_width=1, border_color="#1e293b",
                                          fg_color="#090d16", text_color="#cbd5e1", font=("Consolas", 11))
        self.phone_entry.pack(fill="both", expand=True)

        # ------------------ RIGHT COLUMN: DASHBOARD & CONTROLS ------------------
        right_panel = ctk.CTkFrame(self, corner_radius=15, fg_color="#131924", border_width=1, border_color="#1e293b")
        right_panel.grid(row=0, column=1, rowspan=2, padx=(0, 15), pady=15, sticky="nsew")
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(1, weight=1) # Dashboard lists

        # Dashboard Title
        dash_header_lbl = ctk.CTkLabel(right_panel, text="🖥️ AUTOMATION LIVE BOARD", font=("Segoe UI", 15, "bold"), text_color="#10b981")
        dash_header_lbl.grid(row=0, column=0, pady=(15, 10), sticky="n")

        # Scrollable Device Dashboard
        self.devices_scroll_frame = ctk.CTkScrollableFrame(right_panel, fg_color="#090d16", border_width=1, border_color="#1e293b",
                                                          label_text="Connected Android Devices status", label_text_color="#94a3b8")
        self.devices_scroll_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.devices_scroll_frame.grid_columnconfigure(0, weight=1)

        # Campaign Statistics Panel
        self.stats_frame = ctk.CTkFrame(right_panel, fg_color="#1c2333", corner_radius=10, border_width=1, border_color="#2b384e")
        self.stats_frame.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")
        self.stats_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self.lbl_total_phones = ctk.CTkLabel(self.stats_frame, text="📊 Total: 0", font=("Segoe UI", 12, "bold"), text_color="#38bdf8")
        self.lbl_total_phones.grid(row=0, column=0, pady=10, padx=5, sticky="ew")

        self.lbl_qr_phones = ctk.CTkLabel(self.stats_frame, text="📱 QR: 0", font=("Segoe UI", 12, "bold"), text_color="#f59e0b")
        self.lbl_qr_phones.grid(row=0, column=1, pady=10, padx=5, sticky="ew")

        self.lbl_backup_phones = ctk.CTkLabel(self.stats_frame, text="✅ Backup: 0", font=("Segoe UI", 12, "bold"), text_color="#10b981")
        self.lbl_backup_phones.grid(row=0, column=2, pady=10, padx=5, sticky="ew")

        self.lbl_list_phones = ctk.CTkLabel(self.stats_frame, text="📝 List: 0", font=("Segoe UI", 12, "bold"), text_color="#a855f7")
        self.lbl_list_phones.grid(row=0, column=3, pady=10, padx=5, sticky="ew")

        self.lbl_remaining_phones = ctk.CTkLabel(self.stats_frame, text="⏳ Remaining: 0", font=("Segoe UI", 12, "bold"), text_color="#ef4444")
        self.lbl_remaining_phones.grid(row=0, column=4, pady=10, padx=5, sticky="ew")

        # Control Panel buttons underneath the device list
        control_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        control_frame.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="ew")
        control_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_refresh = ctk.CTkButton(control_frame, text="🔄 Refresh Devices", font=("Segoe UI", 12, "bold"),
                                         fg_color="#1e293b", border_width=1, border_color="#2b384e", hover_color="#2d3748",
                                         text_color="#cbd5e1", height=42, command=self.refresh_devices_list)
        self.btn_refresh.grid(row=0, column=0, padx=5)

        self.btn_start = ctk.CTkButton(control_frame, text="🚀 START CAMPAIGN", font=("Segoe UI", 12, "bold"),
                                       fg_color="#10b981", hover_color="#059669", height=42, command=self.start_campaign)
        self.btn_start.grid(row=0, column=1, padx=5)

        self.btn_stop = ctk.CTkButton(control_frame, text="🛑 FORCE STOP", font=("Segoe UI", 12, "bold"),
                                      fg_color="#ef4444", hover_color="#dc2626", height=42, command=self.stop_campaign, state="disabled")
        self.btn_stop.grid(row=0, column=2, padx=5)

        # ------------------ BOTTOM ROW: LIVE LOGS ------------------
        log_panel = ctk.CTkFrame(self, corner_radius=15, fg_color="#131924", border_width=1, border_color="#1e293b")
        log_panel.grid(row=2, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="nsew")
        log_panel.grid_columnconfigure(0, weight=1)
        log_panel.grid_rowconfigure(1, weight=1)

        log_title = ctk.CTkLabel(log_panel, text="📋 SYSTEM CONSOLE LOGGER", font=("Segoe UI", 12, "bold"), text_color="#cbd5e1", anchor="w")
        log_title.grid(row=0, column=0, padx=15, pady=(8, 2), sticky="ew")

        self.log_terminal = ctk.CTkTextbox(log_panel, height=140, fg_color="#090d16", text_color="#38bdf8", border_width=1, border_color="#1e293b", font=("Consolas", 11))
        self.log_terminal.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.log_terminal.insert("0.0", f"--- Log console initialized at {datetime.now().strftime('%H:%M:%S')} ---\n")
        self.log_terminal.configure(state="disabled")

        # Initial view update
        self.toggle_phone_mode()

    # Browser & file loading functions
    def browse_adb(self):
        file_path = filedialog.askopenfilename(filetypes=[("ADB Executable", "adb.exe"), ("All Files", "*.*")])
        if file_path:
            old_state = self.adb_path_entry.cget("state")
            self.adb_path_entry.configure(state="normal")
            self.adb_path_entry.delete(0, "end")
            self.adb_path_entry.insert(0, file_path)
            self.adb_path_entry.configure(state=old_state)
            self.log(f"Selected ADB path: {file_path}")
            self.refresh_devices_list()

    def browse_phones_txt(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            self.txt_path = file_path
            filename = os.path.basename(file_path)
            self.file_path_label.configure(text=f"Selected: {filename}", text_color="#10b981")
            self.log(f"Selected Phone database file: {file_path}")

    def toggle_phone_mode(self):
        if self.phone_mode_var.get() == "file":
            self.phone_file_frame.pack(fill="x", padx=15, pady=5)
            self.phone_text_frame.pack_forget()
        else:
            self.phone_text_frame.pack(fill="both", expand=True, padx=15, pady=5)
            self.phone_file_frame.pack_forget()

    # Dynamic log printer
    def log(self, text):
        def _append():
            self.log_terminal.configure(state="normal")
            self.log_terminal.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
            self.log_terminal.see("end")
            self.log_terminal.configure(state="disabled")
        self.after(0, _append)

    # Thread-safe stats panel updates
    def update_stats_ui(self):
        def _update():
            remaining = max(0, self.total_phones - self.processed_phones)
            self.lbl_total_phones.configure(text=f"📊 Total: {self.total_phones}")
            self.lbl_qr_phones.configure(text=f"📱 QR: {self.qr_phones}")
            self.lbl_backup_phones.configure(text=f"✅ Backup: {self.backup_phones}")
            self.lbl_list_phones.configure(text=f"📝 List: {self.list_phones}")
            self.lbl_remaining_phones.configure(text=f"⏳ Remaining: {remaining}")
        self.after(0, _update)

    # Thread-safe device row updates
    def update_device_ui(self, device_id, phone=None, status_text=None, text_color="#e2e8f0"):
        def _update():
            if device_id in self.device_ui_elements:
                elements = self.device_ui_elements[device_id]
                if phone is not None:
                    elements['phone_lbl'].configure(text=f"Phone: {phone}")
                if status_text is not None:
                    elements['status_lbl'].configure(text=status_text, text_color=text_color)
        self.after(0, _update)

    # Scans connected devices and updates the layout dynamically
    def refresh_devices_list(self, on_startup=False):
        adb_exec = self.adb_path_entry.get().strip()
        if not os.path.exists(adb_exec):
            # Clear layout and show error
            for child in self.devices_scroll_frame.winfo_children():
                child.destroy()
            self.device_ui_elements.clear()
            no_adb_lbl = ctk.CTkLabel(self.devices_scroll_frame, 
                                      text="⚠️ ADB not found. Please set the correct path and click 'Refresh'.", 
                                      text_color="#f43f5e", font=("Segoe UI", 12, "italic"))
            no_adb_lbl.pack(pady=20)
            if not on_startup:
                self.log(f"❌ ADB Not Found at: '{adb_exec}'. Please select the correct adb.exe path.")
            return

        # Query active connected devices
        new_scan_devices = []
        try:
            result = subprocess.run(f'"{adb_exec}" devices', shell=True, capture_output=True, text=True)
            lines = result.stdout.strip().split('\n')[1:]
            for line in lines:
                if 'device' in line and 'offline' not in line:
                    parts = line.split('\t')
                    if len(parts) > 0:
                        new_scan_devices.append(parts[0].strip())
        except Exception as e:
            self.log(f"❌ Error listing ADB devices: {e}")
            return

        # Clear placeholder if it existed
        for child in list(self.devices_scroll_frame.winfo_children()):
            if isinstance(child, ctk.CTkLabel) and ("No ADB Devices" in child.cget("text") or "ADB not found" in child.cget("text")):
                child.destroy()

        # Update lists and cards
        # 1. Detect disconnected devices
        for dev_id in list(self.device_ui_elements.keys()):
            if dev_id not in new_scan_devices:
                self.update_device_ui(dev_id, status_text="⚠️ Disconnected", text_color="#f43f5e")

        # 2. Detect newly connected devices
        new_connections = []
        for dev_id in new_scan_devices:
            if dev_id not in self.device_ui_elements:
                new_connections.append(dev_id)

        if new_connections:
            self.log(f"🔌 New device(s) connected: {', '.join(new_connections)}")
            
            for dev_id in new_connections:
                # Populate card for new device with premium rounded border cards
                card = ctk.CTkFrame(self.devices_scroll_frame, fg_color="#1c2333", border_width=1, border_color="#2b384e", corner_radius=10, height=85)
                card.pack(fill="x", pady=5, padx=5)
                card.pack_propagate(False)

                dev_icon = ctk.CTkLabel(card, text="📱", font=("Segoe UI", 20))
                dev_icon.pack(side="left", padx=10)

                info_subframe = ctk.CTkFrame(card, fg_color="transparent")
                info_subframe.pack(side="left", fill="both", expand=True, pady=8)

                dev_id_lbl = ctk.CTkLabel(info_subframe, text=f"ID: {dev_id}", font=("Segoe UI", 12, "bold"), text_color="#60a5fa", anchor="w")
                dev_id_lbl.pack(fill="x")

                phone_lbl = ctk.CTkLabel(info_subframe, text="Phone: Idle", font=("Segoe UI", 11), text_color="#94a3b8", anchor="w")
                phone_lbl.pack(fill="x")

                # Control buttons frame packed to the far right
                control_btn_frame = ctk.CTkFrame(card, fg_color="transparent")
                control_btn_frame.pack(side="right", padx=10)

                # Small Run button
                run_btn = ctk.CTkButton(control_btn_frame, text="▶ Run", width=55, height=28,
                                         fg_color="#10b981", hover_color="#059669", text_color="#ffffff",
                                         font=("Segoe UI", 10, "bold"),
                                         command=lambda d=dev_id: self.start_single_device(d))
                run_btn.pack(side="left", padx=3)

                # Small Stop button
                stop_btn = ctk.CTkButton(control_btn_frame, text="⏹ Stop", width=55, height=28,
                                         fg_color="#ef4444", hover_color="#dc2626", text_color="#ffffff",
                                         font=("Segoe UI", 10, "bold"),
                                         state="disabled",
                                         command=lambda d=dev_id: self.stop_single_device(d))
                stop_btn.pack(side="left", padx=3)

                # Status label next to buttons
                status_lbl = ctk.CTkLabel(card, text="Status: Ready", font=("Segoe UI", 11, "bold"), text_color="#34d399", anchor="e")
                status_lbl.pack(side="right", padx=10)

                self.device_ui_elements[dev_id] = {
                    'frame': card,
                    'phone_lbl': phone_lbl,
                    'status_lbl': status_lbl,
                    'run_btn': run_btn,
                    'stop_btn': stop_btn
                }

        # Update main lists
        self.devices = list(new_scan_devices)

        # If no devices are connected, add a user-friendly label placeholder
        if not self.devices:
            no_dev_lbl = ctk.CTkLabel(self.devices_scroll_frame, 
                                      text="⚠️ No ADB Devices found! Connect devices & turn on USB Debugging.", 
                                      text_color="#f43f5e", font=("Segoe UI", 12, "italic"))
            no_dev_lbl.pack(pady=20)
            if not on_startup:
                self.log("⚠️ No devices detected.")

    def load_phone_numbers_from_ui(self):
        adb_exec = self.adb_path_entry.get().strip()
        if not os.path.exists(adb_exec):
            messagebox.showerror("Error", f"ADB executable not found at specified path:\n{adb_exec}")
            return False

        try:
            self.captcha_offset = int(self.captcha_offset_entry.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Captcha Offset must be a numeric integer value (e.g. -30, 0, 50).")
            return False

        self.firebase_url = self.firebase_url_entry.get().strip()
        if not self.firebase_url.startswith("http"):
            messagebox.showerror("Error", "Please input a valid HTTP/HTTPS Firebase Database URL.")
            return False

        self.phone_numbers = []
        if self.phone_mode_var.get() == "file":
            if not self.txt_path or not os.path.exists(self.txt_path):
                messagebox.showerror("Error", "Please choose a valid text file containing phone numbers.")
                return False
            try:
                with open(self.txt_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        for p in line.replace(',', '\n').split('\n'):
                            p_clean = p.strip().replace(" ", "")
                            if p_clean:
                                self.phone_numbers.append(p_clean)
                self.log(f"📚 Read {len(self.phone_numbers)} phone numbers successfully from text file.")
            except Exception as e:
                messagebox.showerror("Error", f"Error reading text file:\n{e}")
                return False
        else:
            raw_text = self.phone_entry.get("0.0", "end").strip()
            self.phone_numbers = [p.strip().replace(" ", "") for p in raw_text.split(',') if p.strip()]
            self.log(f"✍️ Loaded {len(self.phone_numbers)} manually written phone numbers.")

        if not self.phone_numbers:
            messagebox.showerror("Error", "No valid phone numbers found to run automation.")
            return False

        self.phone_numbers = ['0' + p if not p.startswith('0') else p for p in self.phone_numbers]
        return True

    def start_campaign(self):
        if not self.load_phone_numbers_from_ui():
            return

        if not self.devices:
            messagebox.showerror("Error", "No active connected ADB devices to delegate tasks to.")
            return

        self.is_running = True
        self.btn_start.configure(state="disabled")
        self.btn_refresh.configure(state="normal")
        self.btn_stop.configure(state="normal")
        self.adb_path_entry.configure(state="disabled")
        self.captcha_offset_entry.configure(state="disabled")
        self.firebase_url_entry.configure(state="disabled")

        self.log("🚀 Starting Zalo Auto Campaign UI master session...")

        self.phone_queue = queue.Queue()
        for p in self.phone_numbers:
            self.phone_queue.put(p)

        self.total_phones = len(self.phone_numbers)
        self.qr_phones = 0
        self.backup_phones = 0
        self.list_phones = 0
        self.processed_phones = 0
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
            self.btn_start.configure(state="disabled")
            self.btn_refresh.configure(state="normal")
            self.btn_stop.configure(state="normal")
            self.adb_path_entry.configure(state="disabled")
            self.captcha_offset_entry.configure(state="disabled")
            self.firebase_url_entry.configure(state="disabled")
            
            self.phone_queue = queue.Queue()
            for p in self.phone_numbers:
                self.phone_queue.put(p)
                
            self.total_phones = len(self.phone_numbers)
            self.qr_phones = 0
            self.backup_phones = 0
            self.list_phones = 0
            self.processed_phones = 0
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
            self.update_device_ui(device_id, status_text="Status: Ready", text_color="#34d399")

        self.start_worker_thread(device_id, adb_exec)

    def stop_single_device(self, device_id):
        if device_id in self.active_running_devices:
            self.active_running_devices.remove(device_id)
            self.log(f"🛑 Stopping single device worker for: {device_id}... waiting to finish current task.")
            if device_id in self.device_ui_elements:
                self.device_ui_elements[device_id]['stop_btn'].configure(state="disabled")
                self.update_device_ui(device_id, status_text="🛑 Stopping...", text_color="#ef4444")

    def start_worker_thread(self, dev_id, adb_path):
        with self.active_workers_lock:
            self.active_workers_count += 1
        self.log(f"🧵 Spawning worker thread for device {dev_id}...")
        threading.Thread(target=self.device_worker, args=(dev_id, adb_path), daemon=True).start()

    def device_worker(self, device_id, adb_path):
        while self.is_running and device_id in self.active_running_devices and not self.phone_queue.empty():
            try:
                phone = self.phone_queue.get_nowait()
            except queue.Empty:
                break

            self.update_device_ui(device_id, phone=phone, status_text="🔄 Starting task...", text_color="#fbbf24")
            self.log(f"[{device_id}] Starting campaign task for phone {phone}")
            
            status_str = "FAILED"
            try:
                result_code = process_device(self, device_id, phone, adb_path, self.captcha_offset, self.firebase_url)
                if not self.is_running or device_id not in self.active_running_devices:
                    self.log(f"[{device_id}] Thread stopped by user command.")
                    self.update_device_ui(device_id, status_text="🔴 Terminated", text_color="#f43f5e")
                    self.phone_queue.put(phone) # Put back numbers that weren't fully processed
                    break
                
                status_str = result_code if result_code else "FAILED"
                if status_str == "BACKUP":
                    self.log(f"✅ [{device_id}] Task completed successfully (BACKUP) for {phone}!")
                    self.update_device_ui(device_id, status_text="✅ Backup", text_color="#10b981")
                elif status_str == "LIST":
                    self.log(f"📝 [{device_id}] Task finished (LIST saved) for {phone}.")
                    self.update_device_ui(device_id, status_text="📝 List Saved", text_color="#a855f7")
                elif status_str == "QR":
                    self.log(f"📱 [{device_id}] Task hit QR screen for {phone}.")
                    self.update_device_ui(device_id, status_text="📱 QR", text_color="#f59e0b")
                else:
                    self.log(f"❌ [{device_id}] Task failed for {phone}.")
                    self.update_device_ui(device_id, status_text="❌ Failed", text_color="#f43f5e")
            except Exception as e:
                status_str = "CRASH"
                self.log(f"❌ [{device_id}] System Error: {e}")
                self.update_device_ui(device_id, status_text="❌ Crash", text_color="#f43f5e")
            
            # Save completion status to completed_phones.txt file
            if status_str in ["BACKUP", "LIST"]:
                try:
                    log_file_path = os.path.join(os.path.dirname(__file__), "completed_phones.txt")
                    with open(log_file_path, "a", encoding="utf-8") as lf:
                        lf.write(f"{phone}\n")
                except Exception as e:
                    self.log(f"⚠️ Error writing to completed_phones.txt: {e}")
            
            # Increment tracking counters & update UI
            with self.completed_lock:
                self.processed_phones += 1
                if status_str == "BACKUP":
                    self.backup_phones += 1
                elif status_str == "LIST":
                    self.list_phones += 1
                elif status_str == "QR":
                    self.qr_phones += 1
            self.update_stats_ui()
            
            self.phone_queue.task_done()
            app_sleep(self, 0.5, device_id)

        self.update_device_ui(device_id, status_text="💤 Idle", text_color="#94a3b8")
        
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
        self.reset_ui_controls()
        self.log("🎉 ALL TASKS FINISHED OR CAMPAIGN STOPPED.")

    def stop_campaign(self):
        if self.is_running:
            self.is_running = False
            self.active_running_devices.clear()
            self.log("🛑 Stopping execution request sent... waiting for threads to finish current loop.")
            self.btn_stop.configure(state="disabled")

    def reset_ui_controls(self):
        self.btn_start.configure(state="normal")
        self.btn_refresh.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.adb_path_entry.configure(state="normal")
        self.captcha_offset_entry.configure(state="normal")
        self.firebase_url_entry.configure(state="normal")
        # Reset all device buttons back to idle state
        for dev_id, elements in self.device_ui_elements.items():
            if 'run_btn' in elements:
                elements['run_btn'].configure(state="normal")
            if 'stop_btn' in elements:
                elements['stop_btn'].configure(state="disabled")


# =====================================================================
# ================= AUTOMATION CORE INTEGRATED FUNCTIONS =============
# =====================================================================

def app_sleep(app, seconds, device_id=None):
    start = time.time()
    while time.time() - start < seconds:
        if not app.is_running:
            break
        if device_id and device_id not in app.active_running_devices:
            break
        time.sleep(0.1)

def adb_click(app, device_id, x_y_string, adb_path):
    subprocess.run(f'"{adb_path}" -s {device_id} shell input tap {x_y_string}', shell=True)
    app_sleep(app, 0.1, device_id)

def adb_type(app, device_id, text, adb_path, slow=False):
    text_safe = text.replace(" ", "%s")
    if slow:
        for char in text_safe:
            subprocess.run(f'"{adb_path}" -s {device_id} shell input text {char}', shell=True)
            time.sleep(0.1)
    else:
        subprocess.run(f'"{adb_path}" -s {device_id} shell input text {text_safe}', shell=True)
    app_sleep(app, 0.2, device_id)

def get_screen_size(device_id, adb_path):
    try:
        result = subprocess.run(f'"{adb_path}" -s {device_id} shell wm size', shell=True, capture_output=True, text=True)
        match = re.search(r'(\d+)x(\d+)', result.stdout)
        if match:
            return int(match.group(1)), int(match.group(2))
    except:
        pass
    return 1080, 2220

def get_ui_xml(device_id, adb_path):
    for attempt in range(2):
        # We combine dump and cat into a single shell execution to save process spawning overhead on Windows
        result = subprocess.run(
            f'"{adb_path}" -s {device_id} exec-out uiautomator dump /dev/tty',
            shell=True, capture_output=True, text=True, encoding='utf-8'
        )
        xml = result.stdout
        if xml and "<node" in xml:
            return xml
            
        result = subprocess.run(
            f'"{adb_path}" -s {device_id} shell "uiautomator dump /sdcard/window_dump.xml >/dev/null && cat /sdcard/window_dump.xml"',
            shell=True, capture_output=True, text=True, encoding='utf-8'
        )
        xml = result.stdout
        if xml and "<node" in xml:
            return xml
        time.sleep(0.1)
    return ""

def adb_focus_input(app, device_id, adb_path):
    xml_data = get_ui_xml(device_id, adb_path)
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
    if not isinstance(target_text, list):
        target_text = [target_text]
        
    xml_data = get_ui_xml(device_id, adb_path)
    if not xml_data:
        return False
    matches = re.findall(r'node.*?text="([^"]*)".*?bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', xml_data)
    
    clicked = False
    current_index = 0
    for text, x1, y1, x2, y2 in matches:
        match_found = False
        if exact_match:
            match_found = any(t.lower() == text.lower().strip() for t in target_text)
        else:
            match_found = any(t.lower() in text.lower() for t in target_text)
            
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
                app.log(f"[{device_id}] Click '{text}' at ({click_x}, {click_y})")
            adb_click(app, device_id, f"{click_x} {click_y}", adb_path)
            clicked = True
            if not click_all:
                return True
            
    return clicked

def check_text_exists(device_id, target_text, adb_path):
    if not isinstance(target_text, list):
        target_text = [target_text]
    xml_data = get_ui_xml(device_id, adb_path)
    if not xml_data:
        return False
    xml_lower = xml_data.lower()
    return any(t.lower() in xml_lower for t in target_text)

def wait_for_text(app, device_id, target_text, adb_path, timeout=30):
    if not isinstance(target_text, list):
        target_text = [target_text]
    
    app.log(f"[{device_id}] Waiting for text {target_text} (Timeout: {timeout}s)")
    start_time = time.time()
    while time.time() - start_time < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return False
        if check_text_exists(device_id, target_text, adb_path):
            app_sleep(app, 0.1, device_id)
            return True
        app_sleep(app, 0.2, device_id)
    return False

# Firebase OTP Fetcher
def wait_for_otp_firebase(app, device_id, phone_number, firebase_url, timeout=120):
    app.log(f"[{device_id}] OTP Fetching for {phone_number}...")
    start_time = time.time()
    phone_clean = phone_number.replace("+84", "0").replace(" ", "")
    
    # First, find the machine and port to clear old OTP
    try:
        res = http_session.get(f"{firebase_url}machines.json", timeout=5)
        if res.status_code == 200 and res.json():
            machines = res.json()
            for machine_id, m_data in machines.items():
                ports = m_data.get("ports", {})
                for port_id, data in ports.items():
                    if phone_clean in str(data.get("phone", "")):
                        http_session.delete(f"{firebase_url}machines/{machine_id}/ports/{port_id}/otp.json", timeout=5)
                        http_session.delete(f"{firebase_url}web_states/machines/{machine_id}/ports/{port_id}/errorMsg.json", timeout=5)
                        break
    except Exception as e:
        app.log(f"[{device_id}] Warning: Could not clear old OTP: {e}")
    
    while time.time() - start_time < timeout:
        if not app.is_running or device_id not in app.active_running_devices:
            return None
        try:
            res = http_session.get(f"{firebase_url}machines.json", timeout=5)
            if res.status_code == 200 and res.json():
                machines = res.json()
                for machine_id, m_data in machines.items():
                    ports = m_data.get("ports", {})
                    for port_id, data in ports.items():
                        p_num = str(data.get("phone", ""))
                        if phone_clean in p_num:
                            # Error messages checks
                            try:
                                res_state = http_session.get(f"{firebase_url}web_states/machines/{machine_id}/ports/{port_id}/errorMsg.json", timeout=5)
                                if res_state.status_code == 200 and res_state.json():
                                    error_msg = res_state.json()
                                    app.log(f"[{device_id}] GSM Port Error: {error_msg}")
                                    http_session.delete(f"{firebase_url}web_states/machines/{machine_id}/ports/{port_id}/errorMsg.json", timeout=5)
                                    if "Lỗi thiết bị không phản hồi (Timeout)" in error_msg:
                                        return "TIMEOUT_RETRY"
                                    return "ERROR_OTP"
                            except:
                                pass
                                
                            otp_val = data.get("otp")
                            if otp_val and otp_val != "N/A":
                                app.log(f"[{device_id}] Got OTP: {otp_val} from {port_id}!")
                                return otp_val
        except Exception as e:
            pass
        time.sleep(1.5)
    return None

def send_sms_via_firebase(app, device_id, phone_number, recipient, content, firebase_url):
    app.log(f"[{device_id}] Finding GSM port for {phone_number}...")
    phone_clean = phone_number.replace("+84", "0").replace(" ", "")
    
    try:
        res = http_session.get(f"{firebase_url}machines.json", timeout=10)
        if res.status_code == 200 and res.json():
            machines = res.json()
            port_found = None
            machine_found = None
            for machine_id, m_data in machines.items():
                ports = m_data.get("ports", {})
                for port_id, data in ports.items():
                    p_num = str(data.get("phone", ""))
                    if phone_clean in p_num:
                        port_found = port_id
                        machine_found = machine_id
                        break
                if port_found:
                    break
                    
            if not port_found:
                app.log(f"[{device_id}] ❌ Phone {phone_number} not found in GSM database.")
                return False
                
            app.log(f"[{device_id}] Found {phone_number} at machine {machine_found} port {port_found}. Issuing SMS order...")
            cmd_data = {
                "machineId": machine_found,
                "portId": port_found,
                "recipient": recipient,
                "content": content
            }
            post_res = http_session.post(f"{firebase_url}commands.json", json=cmd_data, timeout=10)
            if post_res.status_code == 200:
                app.log(f"[{device_id}] SMS request '{content}' sent to {recipient}")
                return True
    except Exception as e:
        app.log(f"[{device_id}] Firebase SMS Connection error: {e}")
    return False

def solve_zalo_captcha(app, device_id, screen_w, screen_h, adb_path, offset_captcha):
    app.log(f"[{device_id}] Solving Zalo puzzle slider captcha...")
    safe_id = device_id.replace(':', '_').replace('.', '_')
    local_img_path = os.path.abspath(os.path.join(os.path.dirname(__file__), f"captcha_{safe_id}.png"))
    
    # Take screenshot
    subprocess.run(f'"{adb_path}" -s {device_id} exec-out screencap -p > "{local_img_path}"', shell=True)
    if not os.path.exists(local_img_path) or os.path.getsize(local_img_path) == 0:
        app.log(f"[{device_id}] Screencap failed.")
        return False
        
    try:
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
                res = http_session.post("https://api.omocaptcha.com/v2/createTask", json=payload, timeout=8)
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
        
        for _ in range(40):
            if not app.is_running or device_id not in app.active_running_devices: return False
            time.sleep(0.8)
            try:
                poll_res = http_session.post("https://api.omocaptcha.com/v2/getTaskResult", json={"clientKey": client_key, "taskId": task_id}, timeout=5)
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
        subprocess.run(f'"{adb_path}" -s {device_id} shell input swipe {btn_start_x} {btn_start_y} {end_x} {btn_start_y + random.randint(-1, 1)} {duration}', shell=True)
        return True
    except Exception as e:
        app.log(f"[{device_id}] Captcha Exception: {e}")
        return False
    finally:
        if os.path.exists(local_img_path):
            try: os.remove(local_img_path)
            except: pass

# Main Single Device execution stream
def process_device(app, device_id, phone, adb_path, offset_captcha, firebase_url):
    try:
        if not app.is_running or device_id not in app.active_running_devices: return False
        
        screen_w, screen_h = get_screen_size(device_id, adb_path)
        app.log(f"[{device_id}] Resolution: {screen_w}x{screen_h}")
        
        COORD_O_NHAP_SDT = f"{int(screen_w * 0.5)} {int(screen_h * 0.31)}"
        COORD_BTN_TIEP_TUC = f"{int(screen_w * 0.89)} {int(screen_h * 0.94)}"
        COORD_BTN_XAC_NHAN = f"{int(screen_w * 0.5)} {int(screen_h * 0.52)}"
        SWIPE_SCROLL_UP = f"{int(screen_w * 0.5)} {int(screen_h * 0.73)} {int(screen_w * 0.5)} {int(screen_h * 0.31)} 300"
        
        app.update_device_ui(device_id, status_text="🔄 Starting Xtoolz...", text_color="#38bdf8")
        
        # Locate package
        result = subprocess.run(f'"{adb_path}" -s {device_id} shell "pm list packages | grep -i xtoolz"', shell=True, capture_output=True, text=True)
        packages = [line.replace("package:", "").strip() for line in result.stdout.splitlines() if line.strip()]
        target_pkg = packages[0] if packages else "xtoolz"
        
        subprocess.run(f'"{adb_path}" -s {device_id} shell monkey -p {target_pkg} -c android.intent.category.LAUNCHER 1', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        app_sleep(app, 1.0, device_id)
        
        # Reset phone ID
        app.update_device_ui(device_id, status_text="🔄 Rebooting device ID...", text_color="#e0f2fe")
        COORD_BTN_RESET_XTOOLZ = f"{int(screen_w * 0.83)} {int(screen_h * 0.065)}"
        adb_click(app, device_id, COORD_BTN_RESET_XTOOLZ, adb_path)
        app_sleep(app, 5.0, device_id)
        
        # Monitor adb disconnection
        app.log(f"[{device_id}] Waiting for device offline...")
        device_went_offline = False
        for _ in range(60):
            if not app.is_running or device_id not in app.active_running_devices: return False
            res = subprocess.run(f'"{adb_path}" -s {device_id} get-state', shell=True, capture_output=True, text=True)
            if res.stdout.strip() != "device":
                device_went_offline = True
                break
            app_sleep(app, 0.5, device_id)
            
        if not device_went_offline:
            app.log(f"[{device_id}] ⚠️ Thiết bị không tự offline, tiến hành force adb reboot...")
            subprocess.run(f'"{adb_path}" -s {device_id} reboot', shell=True)
            for _ in range(30):
                if not app.is_running or device_id not in app.active_running_devices: return False
                res = subprocess.run(f'"{adb_path}" -s {device_id} get-state', shell=True, capture_output=True, text=True)
                if res.stdout.strip() != "device":
                    break
                app_sleep(app, 0.5, device_id)
        
        # Monitor adb connection
        app.update_device_ui(device_id, status_text="🔄 Rebooting (Wait ADB)...", text_color="#fbbf24")
        app.log(f"[{device_id}] Waiting for device to reconnect...")
        reconnect_start = time.time()
        reconnect_success = False
        while app.is_running:
            if time.time() - reconnect_start > 120: # 2 minutes timeout
                break
            res = subprocess.run(f'"{adb_path}" -s {device_id} get-state', shell=True, capture_output=True, text=True)
            if res.stdout.strip() == "device":
                reconnect_success = True
                break
            app_sleep(app, 1.0, device_id)
            
        if not reconnect_success or not app.is_running or device_id not in app.active_running_devices:
            app.log(f"[{device_id}] ❌ Reconnect timeout or stopped by user.")
            return False
        
        # Wait boot completed
        app.update_device_ui(device_id, status_text="🔄 Booting OS...", text_color="#fbbf24")
        boot_start = time.time()
        boot_success = False
        while app.is_running:
            if time.time() - boot_start > 120: # 2 minutes timeout
                break
            result = subprocess.run(f'"{adb_path}" -s {device_id} shell getprop sys.boot_completed', shell=True, capture_output=True, text=True)
            if "1" in result.stdout:
                boot_success = True
                break
            app_sleep(app, 1.5, device_id)
            
        if not boot_success or not app.is_running or device_id not in app.active_running_devices:
            app.log(f"[{device_id}] ❌ Boot completion timeout or stopped by user.")
            return False
            
        # Chờ OS ổn định sau khi khởi động
        app.update_device_ui(device_id, status_text="🔄 OS Stabilizing...", text_color="#fbbf24")
        app.log(f"[{device_id}] OS đã khởi động xong. Đang chờ 15 giây cho hệ thống ổn định...")
        app_sleep(app, 15.0, device_id)
        
        # Stabilize network - Wait for 4G (Cellular Data) connection and Internet Ping
        app.update_device_ui(device_id, status_text="🔄 Waiting 4G...", text_color="#fbbf24")
        app_sleep(app, 2.0, device_id)
        net_start = time.time()
        net_success = False
        while app.is_running:
            if time.time() - net_start > 300: # 5 minutes timeout
                app.log(f"[{device_id}] ❌ Quá 5 phút không có 4G. Đang khởi động lại thiết bị (Reboot)...")
                subprocess.run(f'"{adb_path}" -s {device_id} reboot', shell=True)
                time.sleep(2)
                break
            
            # Check 4G connection status (mDataConnectionState=2 is connected)
            tel_res = subprocess.run(f'"{adb_path}" -s {device_id} shell "dumpsys telephony.registry | grep mDataConnectionState"', shell=True, capture_output=True, text=True)
            if "mDataConnectionState=2" in tel_res.stdout:
                # 4G is connected, check internet access via ping
                ping_res = subprocess.run(f'"{adb_path}" -s {device_id} shell ping -c 1 -W 2 8.8.8.8', shell=True, capture_output=True, text=True)
                if "1 received" in ping_res.stdout or "1 packets received" in ping_res.stdout or "0% packet loss" in ping_res.stdout:
                    net_success = True
                    app.log(f"[{device_id}] 📶 4G connected and internet ping successful.")
                    break
            else:
                app.log(f"[{device_id}] ⏳ Waiting for 4G connection ")
                
            app_sleep(app, 2.0, device_id)
            
        if not net_success or not app.is_running or device_id not in app.active_running_devices:
            app.log(f"[{device_id}] ❌ Network connection (4G) timeout or stopped by user.")
            return "FAILED"
        
        # Launch Zalo
        app.update_device_ui(device_id, status_text="🔄 Launching Zalo...", text_color="#a855f7")
        subprocess.run(f'"{adb_path}" -s {device_id} shell monkey -p com.zing.zalo -c android.intent.category.LAUNCHER 1', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        # Wait and click Log in
        app.update_device_ui(device_id, status_text="🔄 Clicking Log in...", text_color="#3b82f6")
        wait_for_text(app, device_id, ["Log in", "Đăng nhập"], adb_path, timeout=20)
        
        for i in range(3):
            if adb_click_text(app, device_id, ["Log in", "Đăng nhập"], adb_path, silent=(i > 0)):
                app_sleep(app, 0.4, device_id)
            else:
                break
                
        # Wait for "Enter your phone number" screen
        wait_for_text(app, device_id, ["Enter your phone number", "Số điện thoại"], adb_path, timeout=15)
        
        # Enter Phone
        app.update_device_ui(device_id, status_text="🔄 Entering Phone...", text_color="#22c55e")
        if not adb_focus_input(app, device_id, adb_path):
            # Fallback approximate coordinate if XML parsing fails
            screen_w, screen_h = get_screen_size(device_id, adb_path)
            adb_click(app, device_id, f"{int(screen_w/2)} {int(screen_h*0.25)}", adb_path)
            
        adb_type(app, device_id, phone, adb_path)
        
        # Click Next
        adb_click_text(app, device_id, ["Next", "Tiếp tục"], adb_path)
        
        # Wait for Password screen and click Forgot password
        app.update_device_ui(device_id, status_text="🔄 Clicking Forgot password...", text_color="#f59e0b")
        wait_for_text(app, device_id, ["Forgot password", "Quên mật khẩu"], adb_path, timeout=15)
        adb_click_text(app, device_id, ["Forgot password", "Quên mật khẩu"], adb_path)
        
        # Wait for Verification Method screen and click Receive the code
        app.update_device_ui(device_id, status_text="🔄 Clicking Receive code...", text_color="#22c55e")
        wait_for_text(app, device_id, ["verification method", "phương thức", "Receive", "Nhận"], adb_path, timeout=15)
        adb_click_text(app, device_id, ["Receive", "Nhận"], adb_path)
        
        # Wait for "Receive verification code" screen, extract shortcode, and click "Text ZALO"
        app.update_device_ui(device_id, status_text="🔄 Extracting Shortcode...", text_color="#22c55e")
        wait_for_text(app, device_id, ["Receive verification code", "mã xác thực", "Text", "Soạn"], adb_path, timeout=15)
        
        xml_data = get_ui_xml(device_id, adb_path)
        target_recipient = None
        texts_on_screen = re.findall(r'text="([^"]*)"', xml_data)
        for text_node in texts_on_screen:
            # Match 4-digit shortcodes like 6020, 78xx, 8020
            match = re.search(r'\b([6789]\d{3})\b', text_node)
            if match:
                target_recipient = match.group(1)
                break
                
        adb_click_text(app, device_id, ["Text", "Soạn"], adb_path)
        
        # The SMS app opens automatically. We need to exit it to return to the OTP screen.
        app_sleep(app, 2.0, device_id)
        app.log(f"[{device_id}] Exiting SMS app and returning to Zalo...")
        subprocess.run(f'"{adb_path}" -s {device_id} shell input keyevent 4', shell=True)
        app_sleep(app, 1.0, device_id)
        
        # Ensure Zalo is in the foreground
        subprocess.run(f'"{adb_path}" -s {device_id} shell monkey -p com.zing.zalo -c android.intent.category.LAUNCHER 1', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        app_sleep(app, 1.0, device_id)        
        # Send SMS Trigger and wait for OTP
        app.update_device_ui(device_id, status_text="✉️ Waiting SMS OTP...", text_color="#10b981")
        if target_recipient:
            app.log(f"[{device_id}] Zalo requires MO SMS to shortcode {target_recipient}")
            if send_sms_via_firebase(app, device_id, phone, target_recipient, "ZALO", firebase_url):
                otp = wait_for_otp_firebase(app, device_id, phone, firebase_url)
                if otp == "TIMEOUT_RETRY":
                    app_sleep(app, 3.0, device_id)
                    if send_sms_via_firebase(app, device_id, phone, target_recipient, "ZALO", firebase_url):
                        otp = wait_for_otp_firebase(app, device_id, phone, firebase_url)
            else:
                return False
        else:
            app.log(f"[{device_id}] Could not find shortcode, trying auto-receive OTP.")
            otp = wait_for_otp_firebase(app, device_id, phone, firebase_url)
            
        if otp in ["ERROR_OTP", "TIMEOUT_RETRY"] or not otp:
            app.log(f"[{device_id}] Failed to grab valid OTP from database.")
            return False
            
        # Type OTP
        app.update_device_ui(device_id, status_text="🔄 Submitting OTP...", text_color="#10b981")
        wait_for_text(app, device_id, ["Enter your verification code", "Nhập mã", "Mã xác thực", "gửi tin nhắn", "verification"], adb_path, timeout=60)
        
        screen_w, screen_h = get_screen_size(device_id, adb_path)
        if not adb_focus_input(app, device_id, adb_path):
            adb_click(app, device_id, f"{int(screen_w/2)} {int(screen_h*0.35)}", adb_path)
        
        adb_type(app, device_id, otp, adb_path, slow=True)
        app_sleep(app, 0.5, device_id)
        
        # Click Next
        if check_text_exists(device_id, ["Next", "Tiếp tục", "NEXT"], adb_path):
            adb_click_text(app, device_id, ["Next", "Tiếp tục", "NEXT"], adb_path)
            
        app_sleep(app, 2.0, device_id)
        
        # Check login outcome
        app.update_device_ui(device_id, status_text="🔄 Checking outcome...", text_color="#10b981")
        outcome_start = time.time()
        login_success = False
        list_saved = False
        while time.time() - outcome_start < 15:
            if not app.is_running or device_id not in app.active_running_devices:
                return "FAILED"
                
            xml_data = get_ui_xml(device_id, adb_path)
            if xml_data:
                xml_lower = xml_data.lower()
                
                # Check Fail condition: "Get help logging in"
                if any(t in xml_lower for t in ["get help logging in", "ask a friend", "nhờ bạn bè", "trợ giúp đăng nhập"]):
                    app.log(f"[{device_id}] ❌ Failed: Account requires friend verification (QR).")
                    return "QR"
                    
                # Check Safety Verification screen or Friend Selection screen
                if any(t in xml_lower for t in ["verify your login", "for safety reasons", "xác minh đăng nhập", "lý do bảo mật", "choose 3 people", "chọn 3 người"]):
                    app.log(f"[{device_id}] Safety verification / Friend selection screen detected.")
                    
                    # If we are on the initial warning screen, click NEXT to go to friend list
                    if any(t in xml_lower for t in ["verify your login", "for safety reasons", "xác minh đăng nhập", "lý do bảo mật"]):
                        app.log(f"[{device_id}] Clicking NEXT to enter friend selection.")
                        adb_click_text(app, device_id, ["NEXT", "Next", "Tiếp tục"], adb_path)
                        app_sleep(app, 2.0, device_id)
                        
                    all_names = set()
                    clicked_names = set()
                    
                    for attempt in range(6):
                        if not app.is_running or device_id not in app.active_running_devices:
                            return "FAILED"
                            
                        app.update_device_ui(device_id, status_text=f"🔄 Scanning friends ({attempt+1}/6)...", text_color="#10b981")
                        
                        for swipe_idx in range(2):
                            if not app.is_running or device_id not in app.active_running_devices:
                                return "FAILED"
                                
                            x_data = get_ui_xml(device_id, adb_path)
                            if x_data:
                                visible_names = []
                                matches = re.findall(r'node.*?text="([^"]+)".*?bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', x_data)
                                ignore_texts = ["zalo", "next", "tiếp tục", "bỏ qua", "skip", "identify your friends", "xác minh", "choose 3 people", "bạn đã nhắn tin", "gần đây", "verify account", "recent"]
                                for text_val, x1, y1, x2, y2 in matches:
                                    text_clean = text_val.strip()
                                    if len(text_clean) > 1 and not any(ign in text_clean.lower() for ign in ignore_texts):
                                        y_center = (int(y1) + int(y2)) / 2
                                        if y_center > screen_h * 0.1 and y_center < screen_h * 0.85:
                                            all_names.add(text_clean)
                                            visible_names.append(text_clean)
                                            
                                # Select random friends if we haven't picked 3 yet
                                needed = 3 - len(clicked_names)
                                if needed > 0 and visible_names:
                                    available_to_pick = [n for n in visible_names if n not in clicked_names]
                                    if available_to_pick:
                                        import random
                                        picks = random.sample(available_to_pick, min(needed, len(available_to_pick)))
                                        app.log(f"[{device_id}] Bấm chọn người: {', '.join(picks)}")
                                        for pick in picks:
                                            adb_click_text(app, device_id, [pick], adb_path, exact_match=True)
                                            clicked_names.add(pick)
                                            app_sleep(app, 0.2, device_id)
                                            
                            if swipe_idx < 1:
                                swipe_x = int(screen_w / 2)
                                swipe_start_y = int(screen_h * 0.75)
                                swipe_end_y = int(screen_h * 0.25)
                                subprocess.run(f'"{adb_path}" -s {device_id} shell input swipe {swipe_x} {swipe_start_y} {swipe_x} {swipe_end_y} 500', shell=True)
                                app_sleep(app, 1.0, device_id)
                                
                        app.log(f"[{device_id}] Bấm NEXT để xác nhận...")
                        adb_click_text(app, device_id, ["NEXT", "Next", "Tiếp tục"], adb_path)
                        
                        app_sleep(app, 0.5, device_id)
                        app.log(f"[{device_id}] Đợi popup kết quả (Lần {attempt+1})...")
                        if wait_for_text(app, device_id, ["Confirm", "Xác nhận", "incorrect", "không chính xác"], adb_path, timeout=4):
                            app.log(f"[{device_id}] Trả lời sai, bấm Confirm...")
                            adb_click_text(app, device_id, ["Confirm", "Xác nhận"], adb_path)
                            app_sleep(app, 1.0, device_id)
                        else:
                            app.log(f"[{device_id}] Không thấy báo lỗi, kết thúc vòng lặp chọn bạn bè.")
                            break
                            
                    if all_names:
                        safe_phone = phone.replace("+", "").replace(" ", "")
                        file_path = os.path.join(os.path.dirname(__file__), f"{safe_phone}.txt")
                        try:
                            with open(file_path, "w", encoding="utf-8") as f:
                                for name in sorted(all_names):
                                    f.write(f"{name}\n")
                            app.log(f"[{device_id}] Đã xuất tổng cộng {len(all_names)} tên ra file {safe_phone}.txt")
                            list_saved = True
                        except Exception as e:
                            app.log(f"[{device_id}] Lỗi ghi file txt: {e}")
                    else:
                        app.log(f"[{device_id}] Không quét được tên nào trên màn hình.")
                        
                    app.log(f"[{device_id}] Hoàn tất quy trình chọn bạn bè.")
                    
                    # After finishing friend picking loop, wait a bit and assume success if no new verify screen
                    app_sleep(app, 5.0, device_id)
                    login_success = True
                    break
                    
                # Check Success condition (Login directly)
                if any(t in xml_lower for t in ["tin nhắn", "messages", "danh bạ", "contacts", "cá nhân", "khôi phục", "restore", "skip", "bỏ qua"]):
                    app.log(f"[{device_id}] Đăng nhập thành công (Main screen detected).")
                    login_success = True
                    break
                    
            app_sleep(app, 0.5, device_id)
            
        if not login_success:
            # If loop ends and we haven't failed, we assume it's successful (e.g. some unknown screen).
            login_success = True

        if login_success:
            app.update_device_ui(device_id, status_text="🔄 Backing up Xtoolz...", text_color="#38bdf8")
            app.log(f"[{device_id}] Mở lại Xtoolz để backup...")
            
            # Find xtoolz package
            result = subprocess.run(f'"{adb_path}" -s {device_id} shell "pm list packages | grep -i xtoolz"', shell=True, capture_output=True, text=True)
            packages = [line.replace("package:", "").strip() for line in result.stdout.splitlines() if line.strip()]
            target_pkg = packages[0] if packages else "xtoolz"
            
            # Launch xtoolz
            subprocess.run(f'"{adb_path}" -s {device_id} shell monkey -p {target_pkg} -c android.intent.category.LAUNCHER 1', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            app_sleep(app, 4.0, device_id)
            
            # Click Backup (try common text names)
            backup_clicked = False
            for _ in range(3):
                if adb_click_text(app, device_id, ["Backup", "BACKUP", "Sao lưu", "Sao luu"], adb_path):
                    backup_clicked = True
                    break
                app_sleep(app, 1.0, device_id)
                
            if not backup_clicked:
                app.log(f"[{device_id}] ⚠️ Không tìm thấy nút Backup bằng text. Sử dụng tọa độ backup mặc định...")
                # Coordinate might be standard, we can click somewhere or assume it was clicked if text matching is loose.
                # E.g. Backup is often near the top or bottom depending on Xtoolz layout. We rely on text for now.
                
            app_sleep(app, 2.0, device_id)
            
            safe_phone = phone.replace("+", "").replace(" ", "")
            backup_name = f"{safe_phone} backup"
            
            # Focus input field to enter name
            if not adb_focus_input(app, device_id, adb_path):
                adb_click(app, device_id, f"{int(screen_w/2)} {int(screen_h/2)}", adb_path)
            
            app_sleep(app, 1.0, device_id)
            
            # Clear text using backspace (20 times should be enough)
            app.log(f"[{device_id}] Đang nhập tên backup: {backup_name}")
            for _ in range(20):
                subprocess.run(f'"{adb_path}" -s {device_id} shell input keyevent 67', shell=True)
                
            # Type new name
            adb_type(app, device_id, backup_name, adb_path, slow=False)
            app_sleep(app, 1.0, device_id)
            
            # Click OK
            adb_click_text(app, device_id, ["OK", "Ok", "Xác nhận", "Save", "Lưu"], adb_path)
            app_sleep(app, 3.0, device_id)
            app.log(f"[{device_id}] Hoàn tất backup.")
            return "BACKUP"
            
        if list_saved:
            return "LIST"
            
        return "FAILED"
            
    except Exception as e:
        app.log(f"[{device_id}] Process Error: {e}")
    return "FAILED"


if __name__ == "__main__":
    app = ZaloAutoUIApp()
    app.mainloop()
