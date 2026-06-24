import customtkinter as ctk
from tkinter import filedialog, messagebox
import requests
import json
import os
import glob
import sys
import threading
from datetime import datetime

# Force UTF-8 encoding for stdout on Windows to prevent UnicodeEncodeError
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# CustomTkinter setup
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TokenCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("7UP Token Status Checker & Profiler")
        self.geometry("1200x700")
        self.minsize(1000, 600)
        self.configure(fg_color="#090d16") # Premium dark space background

        self.selected_file_path = ""
        self.token_7up = ""

        # Setup history path and database
        self.history_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "checked_history_7up.json")
        self.history_data = self.load_history()
        self.history_card_widgets = {}

        self.create_widgets()
        self.refresh_history_ui()
        self.auto_load_latest_token()

    def create_widgets(self):
        # Grid layout: 3 main columns (Left controls, Middle history, Right details display)
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=3)
        self.grid_columnconfigure(2, weight=4)
        self.grid_rowconfigure(0, weight=1)

        # ------------------ COLUMN 0: IMPORT & CONTROLS ------------------
        left_panel = ctk.CTkFrame(self, corner_radius=15, fg_color="#131924", border_width=1, border_color="#1e293b")
        left_panel.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        left_panel.grid_columnconfigure(0, weight=1)
        left_panel.grid_rowconfigure(4, weight=1) # raw text gets space

        # Panel Header
        lbl_import = ctk.CTkLabel(left_panel, text="📥 INPUT TOKEN SOURCE", font=("Segoe UI", 16, "bold"), text_color="#38bdf8")
        lbl_import.grid(row=0, column=0, pady=(15, 10), padx=15, sticky="w")

        # Import File button
        self.btn_import = ctk.CTkButton(left_panel, text="📁 Import Token JSON File", height=45, font=("Segoe UI", 15, "bold"),
                                         fg_color="#2563eb", hover_color="#1d4ed8", command=self.import_file)
        self.btn_import.grid(row=1, column=0, padx=15, pady=5, sticky="ew")

        self.lbl_filename = ctk.CTkLabel(left_panel, text="No file selected...", text_color="#64748b", font=("Segoe UI", 14, "italic"), anchor="w")
        self.lbl_filename.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")

        # Manual token input label
        lbl_manual = ctk.CTkLabel(left_panel, text="Or paste 7UP JWT Token manually:", font=("Segoe UI", 14, "bold"), text_color="#94a3b8")
        lbl_manual.grid(row=3, column=0, padx=15, pady=(5, 2), sticky="w")

        # TextBox for manual token
        self.txt_token = ctk.CTkTextbox(left_panel, fg_color="#090d16", border_color="#1e293b", border_width=1, text_color="#cbd5e1", font=("Consolas", 15))
        self.txt_token.grid(row=4, column=0, padx=15, pady=5, sticky="nsew")
        self.txt_token.bind("<KeyRelease>", self.on_token_pasted)

        # Action Button: CHECK
        self.btn_check = ctk.CTkButton(left_panel, text="🔎 CHECK STATUS & PROFILE", height=50, font=("Segoe UI", 16, "bold"),
                                        fg_color="#10b981", hover_color="#059669", command=self.start_check_thread)
        self.btn_check.grid(row=5, column=0, padx=15, pady=(15, 10), sticky="ew")

        # Update Data Section
        update_container = ctk.CTkFrame(left_panel, fg_color="transparent")
        update_container.grid(row=6, column=0, padx=15, pady=(5, 0), sticky="nsew")
        left_panel.grid_rowconfigure(6, weight=2) # Give space to update form

        lbl_update = ctk.CTkLabel(update_container, text="🚚 UPDATE PROFILE DATA", font=("Segoe UI", 15, "bold"), text_color="#f43f5e")
        lbl_update.pack(pady=(5, 5), anchor="w")

        # Scrollable form
        self.update_form_scroll = ctk.CTkScrollableFrame(update_container, fg_color="#0f172a", border_width=1, border_color="#1e293b", corner_radius=8)
        self.update_form_scroll.pack(fill="both", expand=True, pady=(0, 10))

        # Fields
        self.update_entries = {}
        
        # Section headers and their fields
        sections = [
            ("👤 THÔNG TIN CƠ BẢN", [
                ("name", "Name"),
                ("fullName", "FullName"),
                ("phone", "Phone"),
                ("avatar", "Avatar URL"),
                ("gender", "Gender (-1: Unknown, 0: Female, 1: Male)")
            ]),
            ("⚙️ HỆ THỐNG / CHIẾN DỊCH", [
                ("_id", "Account ID (_id)"),
                ("digital_campaign_id", "Digital Campaign ID"),
                ("aId", "Zalo Account ID (aId)"),
                ("uId", "Zalo User ID (uId)"),
                ("channel", "Channel (e.g. zalo)"),
                ("cmpKey", "Campaign Key (cmpKey)")
            ]),
            ("🛡️ TRẠNG THÁI & CHÍNH SÁCH", [
                ("isFollow", "isFollow (true/false)"),
                ("isCheat", "isCheat (true/false)"),
                ("levelBlock", "levelBlock (0, 1, ...)"),
                ("deleted", "deleted (0/1)"),
                ("status", "status (0/1)"),
                ("doneForm", "doneForm (0/1)"),
                ("haveInstantWin", "haveInstantWin (0/1)"),
                ("isApprove", "isApprove (0/1)"),
                ("isAcceptByParent", "isAcceptByParent (0/1)"),
                ("isOver18", "isOver18 (0/1)"),
                ("isPolicy", "isPolicy (true/false)")
            ]),
            ("📊 THÔNG TIN PHỤ (EXTRA DATA)", [
                ("extra_campaign", "extraData.campaign"),
                ("extra_campaign_id", "extraData.campaign_id"),
                ("extra_source", "extraData.source"),
                ("extra_medium", "extraData.medium"),
                ("extra_content", "extraData.content"),
                ("extra_audience", "extraData.audience"),
                ("extra_cid", "extraData.cid"),
                ("extra_uid", "extraData.uid"),
                ("extra_code", "extraData.code"),
                ("extra_collectDone", "extraData.collectDone (true/false)")
            ])
        ]
        
        for section_title, fields in sections:
            lbl_sec = ctk.CTkLabel(self.update_form_scroll, text=section_title, font=("Segoe UI", 12, "bold"), text_color="#38bdf8", anchor="w")
            lbl_sec.pack(fill="x", padx=5, pady=(8, 4))
            
            for key, placeholder in fields:
                lbl_field = ctk.CTkLabel(self.update_form_scroll, text=f"  {placeholder}:", font=("Segoe UI", 11), text_color="#94a3b8", anchor="w")
                lbl_field.pack(fill="x", padx=5, pady=(2, 0))
                
                entry = ctk.CTkEntry(self.update_form_scroll, placeholder_text=placeholder, height=30, font=("Segoe UI", 12), fg_color="#1e293b", border_color="#334155")
                entry.pack(fill="x", padx=5, pady=(0, 4))
                self.update_entries[key] = entry

        self.btn_update = ctk.CTkButton(left_panel, text="✏️ UPDATE DATA", height=45, font=("Segoe UI", 15, "bold"),
                                         fg_color="#e11d48", hover_color="#be123c", command=self.start_update_thread)
        self.btn_update.grid(row=7, column=0, padx=15, pady=(0, 15), sticky="ew")

        # ------------------ COLUMN 1: CHECKED HISTORY LIST ------------------
        middle_panel = ctk.CTkFrame(self, corner_radius=15, fg_color="#131924", border_width=1, border_color="#1e293b")
        middle_panel.grid(row=0, column=1, padx=(0, 15), pady=15, sticky="nsew")
        middle_panel.grid_columnconfigure(0, weight=1)
        middle_panel.grid_rowconfigure(2, weight=1) # Scrollable history gets the vertical space

        # Header
        lbl_history = ctk.CTkLabel(middle_panel, text="📜 CHECKED HISTORY", font=("Segoe UI", 16, "bold"), text_color="#fbbf24")
        lbl_history.grid(row=0, column=0, pady=(15, 5), padx=15, sticky="w")

        # Search box + Clear Button
        search_frame = ctk.CTkFrame(middle_panel, fg_color="transparent")
        search_frame.grid(row=1, column=0, padx=15, pady=(5, 10), sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)

        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="🔍 Search phone or name...", height=40, font=("Segoe UI", 15))
        self.search_entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        self.search_entry.bind("<KeyRelease>", self.on_search_key)

        btn_clear = ctk.CTkButton(search_frame, text="🗑️ Clear", width=80, height=40, font=("Segoe UI", 15, "bold"),
                                   fg_color="#ef4444", hover_color="#dc2626", command=self.clear_history)
        btn_clear.grid(row=0, column=1, sticky="e")

        # Scrollable list container
        self.scrollable_history = ctk.CTkScrollableFrame(middle_panel, fg_color="#090d16", border_width=1, border_color="#1e293b")
        self.scrollable_history.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="nsew")

        # ------------------ COLUMN 2: STATUS & DETAILS ------------------
        right_panel = ctk.CTkFrame(self, corner_radius=15, fg_color="#131924", border_width=1, border_color="#1e293b")
        right_panel.grid(row=0, column=2, padx=(0, 15), pady=15, sticky="nsew")
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(2, weight=1) # Raw JSON response section

        # Header
        lbl_details = ctk.CTkLabel(right_panel, text="🖥️ TOKEN LIVE STATUS", font=("Segoe UI", 16, "bold"), text_color="#10b981")
        lbl_details.grid(row=0, column=0, pady=(15, 10), padx=15, sticky="w")

        # Dynamic Status Badge Card
        self.status_card = ctk.CTkFrame(right_panel, fg_color="#1e293b", border_width=1, border_color="#2b384e", corner_radius=10, height=80)
        self.status_card.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="ew")
        self.status_card.grid_columnconfigure(0, weight=1)
        self.status_card.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(self.status_card, text="AWAITING CHECK", font=("Segoe UI", 22, "bold"), text_color="#94a3b8")
        self.lbl_status.pack(expand=True)

        # TabView for details and raw json
        self.tab_view = ctk.CTkTabview(right_panel, fg_color="#090d16", border_width=1, border_color="#1e293b")
        self.tab_view.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="nsew")

        self.tab_details = self.tab_view.add("Detailed Info")
        self.tab_raw = self.tab_view.add("Raw JSON Response")

        # Scrollable container inside detailed info tab to show all fields
        self.scrollable_details = ctk.CTkScrollableFrame(self.tab_details, fg_color="transparent")
        self.scrollable_details.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.scrollable_details.grid_columnconfigure(0, weight=1)
        self.scrollable_details.grid_columnconfigure(1, weight=1)

        # Fields labels dictionary for updating later
        self.info_labels = {}
        fields = [
            ("Họ tên:", "name", 0, 0),
            ("Số điện thoại:", "phone", 0, 1),
            ("Tài khoản ID:", "_id", 1, 0),
            ("Zalo ID:", "aId", 1, 1),
            ("Zalo UID:", "uId", 2, 0),
            ("Giới tính:", "gender", 2, 1),
            ("Trạng thái Cheat:", "isCheat", 3, 0),
            ("Block Level:", "levelBlock", 3, 1),
            ("Zalo Follow:", "isFollow", 4, 0),
            ("Hoàn thành Form:", "doneForm", 4, 1),
            ("Instant Win:", "haveInstantWin", 5, 0),
            ("Chiến dịch ID:", "digital_campaign_id", 5, 1),
            ("Campaign Key:", "cmpKey", 6, 0),
            ("Kênh đăng nhập:", "channel", 6, 1),
            ("Nguồn tham gia:", "source", 7, 0),
            ("Trên 18 tuổi:", "isOver18", 7, 1),
            ("Đồng ý chính sách:", "isPolicy", 8, 0),
            ("Cha mẹ đồng ý:", "isAcceptByParent", 8, 1),
            ("Trạng thái:", "status", 9, 0),
            ("Trạng thái xóa:", "deleted", 9, 1),
            ("Ngày tham gia:", "createdAt", 10, 0),
            ("Cập nhật cuối:", "updatedAt", 10, 1),
            ("Link Avatar:", "avatar", 11, 0)
        ]

        for label_text, key, row, col in fields:
            container = ctk.CTkFrame(self.scrollable_details, fg_color="#1a2230", corner_radius=8, border_width=1, border_color="#2d3748")
            if key == "avatar":
                container.grid(row=row, column=col, columnspan=2, padx=8, pady=6, sticky="ew")
            else:
                container.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
            
            lbl_title = ctk.CTkLabel(container, text=label_text, font=("Segoe UI", 13, "bold"), text_color="#94a3b8", anchor="w")
            lbl_title.pack(fill="x", padx=10, pady=(6, 2))
            
            lbl_val = ctk.CTkLabel(container, text="-", font=("Segoe UI", 16, "bold"), text_color="#f8fafc", anchor="w")
            lbl_val.pack(fill="x", padx=10, pady=(2, 6))
            self.info_labels[key] = lbl_val

        # Setup Raw JSON text box
        self.raw_json_box = ctk.CTkTextbox(self.tab_raw, fg_color="#090d16", text_color="#38bdf8", font=("Consolas", 12))
        self.raw_json_box.pack(fill="both", expand=True, padx=5, pady=5)
        self.raw_json_box.insert("0.0", "No data loaded yet.")
        self.raw_json_box.configure(state="disabled")

    def load_history(self):
        if os.path.exists(self.history_file_path):
            try:
                with open(self.history_file_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading history: {e}")
        return []

    def save_history(self):
        try:
            with open(self.history_file_path, "w", encoding="utf-8") as f:
                json.dump(self.history_data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving history: {e}")

    def add_to_history(self, token, token_to_use, status_text, bg_color, text_color, data_dict, raw_json):
        phone = data_dict.get("phone") if data_dict else None
        existing_index = -1
        
        if phone and phone != "-":
            for i, entry in enumerate(self.history_data):
                if entry.get("data", {}).get("phone") == phone:
                    existing_index = i
                    break
        else:
            for i, entry in enumerate(self.history_data):
                if entry.get("token") == token or entry.get("exchanged_token") == token:
                    existing_index = i
                    break
                    
        entry_data = {
            "token": token,
            "exchanged_token": token_to_use if token_to_use != token else None,
            "status_text": status_text,
            "bg_color": bg_color,
            "text_color": text_color,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data": data_dict or {},
            "raw_json": raw_json
        }
        
        if existing_index != -1:
            self.history_data.pop(existing_index)
            
        self.history_data.insert(0, entry_data)
        
        if len(self.history_data) > 200:
            self.history_data = self.history_data[:200]
            
        self.save_history()
        self.after(0, self.refresh_history_ui)

    def refresh_history_ui(self, filter_text=""):
        for widget in self.scrollable_history.winfo_children():
            widget.destroy()
            
        self.history_card_widgets.clear()
        filter_text = filter_text.lower().strip()
        
        for entry in self.history_data:
            data = entry.get("data") or {}
            name = data.get("name") or data.get("fullName") or "-"
            phone = data.get("phone") or "-"
            token = entry.get("token", "")
            
            if filter_text:
                if filter_text not in name.lower() and filter_text not in phone.lower():
                    continue
                    
            status_text = entry.get("status_text", "UNKNOWN")
            timestamp = entry.get("timestamp", "")
            
            if "ACTIVE" in status_text or "ALIVE" in status_text:
                status_icon = "🟢"
            elif "BANNED" in status_text or "BLOCKED" in status_text:
                status_icon = "🚫"
            else:
                status_icon = "🔴"
                
            card = ctk.CTkFrame(self.scrollable_history, fg_color="#1a2230", border_width=1, border_color="#2d3748", corner_radius=8, cursor="hand2")
            card.pack(fill="x", padx=5, pady=4)
            
            self.history_card_widgets[token] = card
            
            lbl_name = ctk.CTkLabel(card, text=f"{status_icon} {name}", font=("Segoe UI", 17, "bold"), text_color="#38bdf8", anchor="w")
            lbl_name.pack(fill="x", padx=10, pady=(6, 2))
            
            lbl_phone = ctk.CTkLabel(card, text=f"📞 {phone}", font=("Segoe UI", 15, "bold"), text_color="#fbbf24", anchor="w")
            lbl_phone.pack(fill="x", padx=10, pady=(2, 2))
            
            lbl_time = ctk.CTkLabel(card, text=f"🕒 {timestamp}", font=("Segoe UI", 13, "italic"), text_color="#64748b", anchor="w")
            lbl_time.pack(fill="x", padx=10, pady=(2, 6))
            
            def make_select_handler(selected_entry=entry, current_card=card):
                return lambda event: self.select_history_entry(selected_entry, current_card)
                
            card.bind("<Button-1>", make_select_handler())
            lbl_name.bind("<Button-1>", make_select_handler())
            lbl_phone.bind("<Button-1>", make_select_handler())
            lbl_time.bind("<Button-1>", make_select_handler())

    def select_history_entry(self, entry, selected_card):
        for card in self.history_card_widgets.values():
            try:
                card.configure(border_color="#2d3748", border_width=1)
            except:
                pass
                
        if selected_card:
            try:
                selected_card.configure(border_color="#2563eb", border_width=2)
            except:
                pass
            
        token_to_load = entry.get("exchanged_token") or entry.get("token", "")
        self.token_7up = token_to_load
        self.txt_token.delete("0.0", "end")
        self.txt_token.insert("0.0", token_to_load)
        
        status_text = entry.get("status_text", "UNKNOWN")
        bg_color = entry.get("bg_color", "#1e293b")
        text_color = entry.get("text_color", "#94a3b8")
        self.update_status_card(status_text, bg_color, text_color)
        
        data = entry.get("data")
        raw_json = entry.get("raw_json", "")
        
        self.update_raw_json(raw_json)
        
        if "ACTIVE" in status_text or "ALIVE" in status_text or "BANNED" in status_text or "BLOCKED" in status_text:
            self.update_profile_details(data)
        else:
            self.update_profile_error(status_text)

    def on_search_key(self, event):
        search_txt = self.search_entry.get().strip()
        self.refresh_history_ui(search_txt)

    def clear_history(self):
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa toàn bộ lịch sử đã check không?"):
            self.history_data = []
            if os.path.exists(self.history_file_path):
                try:
                    os.remove(self.history_file_path)
                except Exception as e:
                    print(e)
            self.refresh_history_ui()
            self.update_status_card("AWAITING CHECK", "#1e293b", "#94a3b8")
            self.update_profile_error("Lịch sử trống.")

    def auto_load_latest_token(self):
        """Tự động nạp token mới nhất từ lịch sử, hoặc tìm kiếm từ thư mục server nếu lịch sử trống."""
        if self.history_data:
            self.select_history_entry(self.history_data[0], None)
            return

        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Thử tìm file FINAL của 7up trước
        final_files = glob.glob(os.path.join(script_dir, "token_7up_FINAL_*.json"))
        if final_files:
            latest_file = max(final_files, key=os.path.getmtime)
            self.load_token_from_file(latest_file)
            return

        # Nếu không có, tìm file MiniApp token của 7up
        miniapp_files = glob.glob(os.path.join(script_dir, "token_MiniApp_3151270984274302494_*.json"))
        if miniapp_files:
            latest_file = max(miniapp_files, key=os.path.getmtime)
            self.load_token_from_file(latest_file)

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All Files", "*.*")])
        if file_path:
            self.load_token_from_file(file_path)

    def load_token_from_file(self, file_path):
        self.selected_file_path = file_path
        filename = os.path.basename(file_path)
        self.lbl_filename.configure(text=f"Loaded: {filename}", text_color="#10b981")

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Lấy token từ các key có thể có
            token = data.get("sevenup_token") or data.get("access_token") or data.get("token")
            if token:
                self.token_7up = token.strip()
                self.txt_token.delete("0.0", "end")
                self.txt_token.insert("0.0", self.token_7up)
                self.update_status_card("AWAITING CHECK", "#1e293b", "#94a3b8")
            else:
                messagebox.showerror("Error", "File JSON không chứa trường 'sevenup_token' hoặc 'access_token'!")
        except Exception as e:
            messagebox.showerror("Error", f"Lỗi đọc file JSON: {e}")

    def on_token_pasted(self, event):
        raw_text = self.txt_token.get("0.0", "end").strip()
        self.token_7up = raw_text
        self.selected_file_path = ""
        self.lbl_filename.configure(text="Pasted manually...", text_color="#cbd5e1")
        self.update_status_card("AWAITING CHECK", "#1e293b", "#94a3b8")

    def update_status_card(self, text, bg_color, text_color):
        self.lbl_status.configure(text=text, text_color=text_color)
        self.status_card.configure(fg_color=bg_color)

    def start_check_thread(self):
        if not self.token_7up:
            messagebox.showerror("Lỗi", "Vui lòng import file token hoặc paste token thủ công trước!")
            return

        self.btn_check.configure(state="disabled", text="🔍 CHECKING...")
        self.update_status_card("CONNECTING API...", "#1e293b", "#e2e8f0")
        
        # Reset labels
        for key in self.info_labels.keys():
            self.info_labels[key].configure(text="-")

        threading.Thread(target=self.check_token_status, daemon=True).start()

    def check_token_status(self):
        import hashlib
        import time

        token_to_use = self.token_7up.strip()
        
        # Nếu token không phải dạng JWT (thường bắt đầu bằng eyJ), tự động đổi Zalo Token sang 7UP Token
        if not token_to_use.startswith("eyJ"):
            self.after(0, lambda: self.update_status_card("EXCHANGING ZALO TOKEN...", "#1e293b", "#e2e8f0"))
            
            auth_url = "https://7up-april-utc-public-api.adtimabox.vn/digital-api/digital-user-auth/getTokenDigitalMiniApp"
            campaign_id = "69ddf7b66f593f26c5daf9a5"
            secret_key = "mLzu2c89ZRjkP2bN"
            timestamp_ms = int(time.time() * 1000)
            
            sign_str = f"{campaign_id}{secret_key}{timestamp_ms}"
            mac = hashlib.sha256(sign_str.encode('utf-8')).hexdigest()
            
            payload = {
                "campaignId": campaign_id,
                "tokenUser": token_to_use,
                "time": timestamp_ms,
                "mac": mac,
                "channelProvider": "zalo"
            }
            
            headers_auth = {
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0 (Linux; Android 9; Mi A1 Build/PKQ1.180917.001;) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/138.0.7204.179 Mobile Safari/537.36 Zalo android/260501901 ZaloTheme/light ZaloLanguage/vi",
                "Origin": "https://h5.zdn.vn",
                "Referer": "https://h5.zdn.vn/"
            }
            
            try:
                auth_res = requests.post(auth_url, json=payload, headers=headers_auth, timeout=20)
                # Cập nhật phản hồi đổi token tạm thời
                try:
                    auth_json = auth_res.json()
                    raw_text = json.dumps(auth_json, indent=4, ensure_ascii=False)
                except:
                    raw_text = auth_res.text
                self.update_raw_json(raw_text)

                if auth_res.status_code == 200:
                    res_json = auth_res.json()
                    data_obj = res_json.get("data")
                    exchanged_token = data_obj.get("accessToken") if isinstance(data_obj, dict) else None
                    if exchanged_token:
                        token_to_use = exchanged_token
                        self.after(0, lambda: self.update_status_card("EXCHANGED OK, CHECKING...", "#1e293b", "#e2e8f0"))
                    else:
                        self.after(0, lambda: self.update_status_card("🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444"))
                        self.update_profile_error("Đổi token thất bại: Server không trả về accessToken")
                        self.add_to_history(self.token_7up.strip(), self.token_7up.strip(), "🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444", {}, raw_text)
                        self.after(0, lambda: self.btn_check.configure(state="normal", text="🔎 CHECK STATUS & PROFILE"))
                        return
                else:
                    self.after(0, lambda: self.update_status_card("🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444"))
                    self.update_profile_error(f"Lỗi gọi API đổi token: HTTP {auth_res.status_code}")
                    self.add_to_history(self.token_7up.strip(), self.token_7up.strip(), "🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444", {}, raw_text)
                    self.after(0, lambda: self.btn_check.configure(state="normal", text="🔎 CHECK STATUS & PROFILE"))
                    return
            except Exception as e:
                self.update_raw_json(str(e))
                self.after(0, lambda: self.update_status_card("🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444"))
                self.update_profile_error(f"Lỗi kết nối đổi token: {e}")
                self.add_to_history(self.token_7up.strip(), self.token_7up.strip(), "🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444", {}, str(e))
                self.after(0, lambda: self.btn_check.configure(state="normal", text="🔎 CHECK STATUS & PROFILE"))
                return

        # Gọi API lấy thông tin tài khoản 7UP
        url = "https://7up-april-utc-public-api.adtimabox.vn/digital-api/digital-user/me"
        headers = {
            "Authorization": f"Bearer {token_to_use}",
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0"
        }

        try:
            response = requests.get(url, headers=headers, timeout=20)
            
            raw_text = ""
            try:
                raw_text = json.dumps(response.json(), indent=4, ensure_ascii=False)
            except:
                raw_text = response.text

            self.update_raw_json(raw_text)

            if response.status_code == 200:
                res_json = response.json()
                data = res_json.get("data", {})
                
                is_cheat = data.get("isCheat", False)
                level_block = data.get("levelBlock", 0)

                if is_cheat or level_block > 0:
                    status_txt = "🚫 BANNED / BLOCKED"
                    status_bg = "#7f1d1d"
                    status_fg = "#ef4444"
                else:
                    status_txt = "🟢 ACTIVE / ALIVE"
                    status_bg = "#064e3b"
                    status_fg = "#10b981"

                self.after(0, lambda: self.update_status_card(status_txt, status_bg, status_fg))
                self.after(0, lambda: self.update_profile_details(data))
                self.add_to_history(self.token_7up.strip(), token_to_use, status_txt, status_bg, status_fg, data, raw_text)

            elif response.status_code in (401, 403):
                status_txt = "🔴 DEAD / EXPIRED"
                status_bg = "#7f1d1d"
                status_fg = "#ef4444"
                self.after(0, lambda: self.update_status_card(status_txt, status_bg, status_fg))
                self.update_profile_error(f"Unauthorized (HTTP {response.status_code})")
                self.add_to_history(self.token_7up.strip(), token_to_use, status_txt, status_bg, status_fg, {}, raw_text)
            else:
                status_txt = "🔴 UNKNOWN / SERVER ERROR"
                status_bg = "#78350f"
                status_fg = "#f59e0b"
                self.after(0, lambda: self.update_status_card(status_txt, status_bg, status_fg))
                self.update_profile_error(f"Server response code: {response.status_code}")
                self.add_to_history(self.token_7up.strip(), token_to_use, status_txt, status_bg, status_fg, {}, raw_text)

        except Exception as e:
            self.raw_json_box.configure(state="normal")
            self.raw_json_box.delete("0.0", "end")
            self.raw_json_box.insert("0.0", str(e))
            self.raw_json_box.configure(state="disabled")
            
            status_txt = "🔴 CONNECTION ERROR"
            status_bg = "#7f1d1d"
            status_fg = "#ef4444"
            
            self.after(0, lambda: self.update_status_card(status_txt, status_bg, status_fg))
            self.update_profile_error(f"Error: {e}")
            self.add_to_history(self.token_7up.strip(), token_to_use, status_txt, status_bg, status_fg, {}, str(e))
        finally:
            self.after(0, lambda: self.btn_check.configure(state="normal", text="🔎 CHECK STATUS & PROFILE"))

    def update_raw_json(self, text):
        def _update():
            self.raw_json_box.configure(state="normal")
            self.raw_json_box.delete("0.0", "end")
            self.raw_json_box.insert("0.0", text)
            self.raw_json_box.configure(state="disabled")
        self.after(0, _update)

    def update_profile_details(self, data):
        # Save loaded profile data for auto-fill in update form
        self.last_loaded_data = data

        def format_date(date_str):
            if not date_str:
                return "-"
            try:
                # Format Zalo timestamp standard "2026-06-13T09:49:11.029Z"
                dt = datetime.strptime(date_str.replace("Z", ""), "%Y-%m-%dT%H:%M:%S.%f")
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            except:
                try:
                    dt = datetime.strptime(date_str.replace("Z", ""), "%Y-%m-%dT%H:%M:%S")
                    return dt.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    return date_str

        # Helper to convert truthy values (True, 1) to YES/NO with custom colors
        def get_yes_no(val):
            if val is True or val == 1 or str(val).lower() == "true":
                return "YES", "#10b981"
            return "NO", "#ef4444"

        # Update values
        self.info_labels["name"].configure(text=data.get("name") or data.get("fullName") or "-", text_color="#38bdf8")
        self.info_labels["phone"].configure(text=data.get("phone") or "-", text_color="#fbbf24")
        self.info_labels["_id"].configure(text=data.get("_id") or "-", text_color="#e2e8f0")
        self.info_labels["aId"].configure(text=data.get("aId") or "-", text_color="#e2e8f0")
        self.info_labels["uId"].configure(text=data.get("uId") or "-", text_color="#e2e8f0")
        
        # Auto-fill update form with currently loaded details
        def safe_insert(entry_widget, value):
            entry_widget.delete(0, "end")
            if value is not None:
                if isinstance(value, bool):
                    entry_widget.insert(0, "true" if value else "false")
                else:
                    entry_widget.insert(0, str(value))

        extra = data.get("extraData") or {}
        for key in self.update_entries:
            if key.startswith("extra_"):
                sub_key = key.replace("extra_", "")
                val = extra.get(sub_key)
            elif key == "name":
                val = data.get("name") or data.get("fullName") or ""
            else:
                val = data.get(key)

            # Special default values for empty fields
            if val is None:
                if key == "avatar":
                    val = "https://s120-ava-talk.zadn.vn/default"
                elif key in ["isApprove", "isOver18", "doneForm", "isAcceptByParent"]:
                    val = 1
                elif key in ["isPolicy", "isFollow"]:
                    val = True
                elif key == "gender":
                    val = -1
                elif key in ["levelBlock", "deleted", "haveInstantWin"]:
                    val = 0
                elif key == "isCheat":
                    val = False
                else:
                    val = ""
            
            safe_insert(self.update_entries[key], val)
        
        gender_map = {-1: "Không xác định", 0: "Nữ", 1: "Nam"}
        gender_val = data.get("gender", -1)
        gender_text = gender_map.get(gender_val, "Không xác định")
        gender_color = "#f472b6" if gender_val == 0 else ("#60a5fa" if gender_val == 1 else "#a1a1aa")
        self.info_labels["gender"].configure(text=gender_text, text_color=gender_color)
        
        is_cheat = data.get("isCheat", False)
        self.info_labels["isCheat"].configure(
            text="YES (Phát hiện cheat)" if is_cheat else "NO",
            text_color="#ef4444" if is_cheat else "#10b981"
        )
        
        level_block = data.get("levelBlock", 0)
        self.info_labels["levelBlock"].configure(
            text=str(level_block),
            text_color="#ef4444" if level_block > 0 else "#10b981"
        )
        
        # isFollow
        follow_txt, follow_col = get_yes_no(data.get("isFollow"))
        self.info_labels["isFollow"].configure(text=follow_txt, text_color=follow_col)
        
        # doneForm
        df_txt, df_col = get_yes_no(data.get("doneForm"))
        self.info_labels["doneForm"].configure(text=df_txt, text_color=df_col)
        
        # haveInstantWin
        hiw = data.get("haveInstantWin", 0)
        self.info_labels["haveInstantWin"].configure(text=str(hiw), text_color="#fbbf24" if hiw > 0 else "#e2e8f0")
        
        # digital_campaign_id
        self.info_labels["digital_campaign_id"].configure(text=data.get("digital_campaign_id") or "-", text_color="#e2e8f0")
        
        # cmpKey
        self.info_labels["cmpKey"].configure(text=data.get("cmpKey") or "-", text_color="#a78bfa")
        
        # channel
        self.info_labels["channel"].configure(text=str(data.get("channel") or "-").upper(), text_color="#38bdf8")
        
        # source under extraData
        extra_data = data.get("extraData") or {}
        source_val = extra_data.get("source") or "-"
        self.info_labels["source"].configure(text=str(source_val), text_color="#34d399")
        
        # isOver18
        o18_txt, o18_col = get_yes_no(data.get("isOver18"))
        self.info_labels["isOver18"].configure(text=o18_txt, text_color=o18_col)
        
        # isPolicy
        policy_txt, policy_col = get_yes_no(data.get("isPolicy"))
        self.info_labels["isPolicy"].configure(text=policy_txt, text_color=policy_col)
        
        # isAcceptByParent
        parent_txt, parent_col = get_yes_no(data.get("isAcceptByParent"))
        self.info_labels["isAcceptByParent"].configure(text=parent_txt, text_color=parent_col)
        
        # status and deleted
        status_txt, status_col = get_yes_no(data.get("status"))
        self.info_labels["status"].configure(text=status_txt, text_color=status_col)
        self.info_labels["deleted"].configure(text=str(data.get("deleted", 0)), text_color="#e2e8f0")

        # createdAt, updatedAt
        self.info_labels["createdAt"].configure(text=format_date(data.get("createdAt")), text_color="#cbd5e1")
        self.info_labels["updatedAt"].configure(text=format_date(data.get("updatedAt")), text_color="#cbd5e1")
        self.info_labels["avatar"].configure(text=data.get("avatar") or "-", text_color="#60a5fa")

    def update_profile_error(self, error_message):
        def _update():
            for key in self.info_labels.keys():
                self.info_labels[key].configure(text="-", text_color="#f8fafc")
            self.info_labels["name"].configure(text=error_message, text_color="#ef4444")
            
            # Clear auto-fill inputs on error
            for entry in self.update_entries.values():
                entry.delete(0, "end")
        self.after(0, _update)

    def start_update_thread(self):
        if not self.token_7up:
            messagebox.showerror("Lỗi", "Vui lòng nhập hoặc chọn token trước!")
            return

        form_data = {}
        for key, entry in self.update_entries.items():
            form_data[key] = entry.get().strip()
        
        # Lấy dữ liệu cũ nếu tên/số điện thoại bị trống
        existing_data = getattr(self, "last_loaded_data", {})
        if not form_data["name"]:
            form_data["name"] = existing_data.get("name") or existing_data.get("fullName") or ""
        if not form_data["phone"]:
            form_data["phone"] = existing_data.get("phone") or ""

        if not form_data["name"] or not form_data["phone"]:
            messagebox.showerror("Lỗi", "Cần phải có Name và Phone (có thể nhập hoặc lấy từ thông tin đã check)!")
            return

        self.btn_update.configure(state="disabled", text="✏️ UPDATING...")
        self.update_status_card("UPDATING DATA...", "#1e293b", "#e2e8f0")
        
        threading.Thread(target=self.update_delivery_data, args=(form_data,), daemon=True).start()

    def update_delivery_data(self, form_data):
        import hashlib
        import time

        token_to_use = self.token_7up.strip()
        
        # Đổi Zalo Token sang 7UP Token nếu cần
        if not token_to_use.startswith("eyJ"):
            self.after(0, lambda: self.update_status_card("EXCHANGING ZALO TOKEN...", "#1e293b", "#e2e8f0"))
            
            auth_url = "https://7up-april-utc-public-api.adtimabox.vn/digital-api/digital-user-auth/getTokenDigitalMiniApp"
            campaign_id = "69ddf7b66f593f26c5daf9a5"
            secret_key = "mLzu2c89ZRjkP2bN"
            timestamp_ms = int(time.time() * 1000)
            
            sign_str = f"{campaign_id}{secret_key}{timestamp_ms}"
            mac = hashlib.sha256(sign_str.encode('utf-8')).hexdigest()
            
            payload = {
                "campaignId": campaign_id,
                "tokenUser": token_to_use,
                "time": timestamp_ms,
                "mac": mac,
                "channelProvider": "zalo"
            }
            
            headers_auth = {
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0",
                "Origin": "https://h5.zdn.vn",
                "Referer": "https://h5.zdn.vn/"
            }
            
            try:
                auth_res = requests.post(auth_url, json=payload, headers=headers_auth, timeout=20)
                try:
                    raw_text = json.dumps(auth_res.json(), indent=4, ensure_ascii=False)
                except:
                    raw_text = auth_res.text
                self.update_raw_json(raw_text)

                if auth_res.status_code == 200:
                    res_json = auth_res.json()
                    exchanged_token = res_json.get("data", {}).get("accessToken") if isinstance(res_json.get("data"), dict) else None
                    if exchanged_token:
                        token_to_use = exchanged_token
                    else:
                        self.after(0, lambda: self.update_status_card("🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444"))
                        self.update_profile_error("Đổi token thất bại: Không có accessToken")
                        self.after(0, lambda: self.btn_update.configure(state="normal", text="✏️ UPDATE DATA"))
                        return
                else:
                    self.after(0, lambda: self.update_status_card("🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444"))
                    self.update_profile_error(f"Lỗi API đổi token: {auth_res.status_code}")
                    self.after(0, lambda: self.btn_update.configure(state="normal", text="✏️ UPDATE DATA"))
                    return
            except Exception as e:
                self.update_raw_json(str(e))
                self.after(0, lambda: self.update_status_card("🔴 EXCHANGE FAILED", "#7f1d1d", "#ef4444"))
                self.update_profile_error(f"Lỗi kết nối đổi token: {e}")
                self.after(0, lambda: self.btn_update.configure(state="normal", text="✏️ UPDATE DATA"))
                return

        # Gọi API lấy thông tin hiện tại trước
        try:
            me_res = requests.get("https://7up-april-utc-public-api.adtimabox.vn/digital-api/digital-user/me", headers={
                "Authorization": f"Bearer {token_to_use}",
                "Accept": "application/json"
            }, timeout=10)
            me_data = me_res.json().get("data", {}) if me_res.status_code == 200 else {}
        except:
            me_data = {}

        # Gọi API Update Profile gốc
        url = "https://7up-april-utc-public-api.adtimabox.vn/digital-api/digital-user"
        headers = {
            "Authorization": f"Bearer {token_to_use}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0"
        }
        
        # Dựng payload từ form_data
        def safe_int(val, default):
            try:
                return int(val)
            except:
                return default

        extra_data_payload = {
            "campaign": form_data.get("extra_campaign", ""),
            "campaign_id": form_data.get("extra_campaign_id", ""),
            "source": form_data.get("extra_source", ""),
            "medium": form_data.get("extra_medium", ""),
            "content": form_data.get("extra_content", ""),
            "audience": form_data.get("extra_audience", ""),
            "cid": form_data.get("extra_cid", ""),
            "uid": form_data.get("extra_uid", ""),
            "code": form_data.get("extra_code", ""),
            "collectDone": form_data.get("extra_collectDone", "true").lower() == "true"
        }

        # Override with values from me_data extraData if inputs are empty
        me_extra = me_data.get("extraData", {}) if isinstance(me_data, dict) else {}
        for key in ["campaign", "campaign_id", "source", "medium", "content", "audience", "cid", "uid", "code"]:
            extra_key = f"extra_{key}"
            if not form_data.get(extra_key) and key in me_extra:
                extra_data_payload[key] = me_extra[key]

        payload = {
            "_id": form_data.get("_id", ""),
            "digital_campaign_id": form_data.get("digital_campaign_id", ""),
            "name": form_data.get("name", ""),
            "fullName": form_data.get("fullName", "") or form_data.get("name", ""),
            "aId": form_data.get("aId", ""),
            "avatar": form_data.get("avatar") or "https://s120-ava-talk.zadn.vn/default",
            "gender": safe_int(form_data.get("gender"), -1),
            "channel": form_data.get("channel", "zalo"),
            "isFollow": form_data.get("isFollow", "true").lower() == "true",
            "isCheat": form_data.get("isCheat", "false").lower() == "true",
            "levelBlock": safe_int(form_data.get("levelBlock"), 0),
            "deleted": safe_int(form_data.get("deleted"), 0),
            "status": safe_int(form_data.get("status"), 1),
            "doneForm": safe_int(form_data.get("doneForm"), 1),
            "haveInstantWin": safe_int(form_data.get("haveInstantWin"), 0),
            "cmpKey": form_data.get("cmpKey", ""),
            "cmp_key": form_data.get("cmpKey", ""),
            "phone": form_data.get("phone", ""),
            "uId": form_data.get("uId", ""),
            "extraData": extra_data_payload,
            "extra_data": extra_data_payload,
            "isAcceptByParent": safe_int(form_data.get("isAcceptByParent"), 1),
            "isOver18": safe_int(form_data.get("isOver18"), 1),
            "isPolicy": form_data.get("isPolicy", "true").lower() == "true"
        }

        try:
            # Gửi PUT
            response = requests.put(url, headers=headers, json=payload, timeout=20)
            try:
                raw_text = json.dumps(response.json(), indent=4, ensure_ascii=False)
            except:
                raw_text = response.text
            
            self.update_raw_json(raw_text)

            if response.status_code == 200:
                self.after(0, lambda text=raw_text: messagebox.showinfo("Thành công", f"Đã gửi yêu cầu cập nhật!\n\nPhản hồi từ server:\n{text}"))
                self.after(0, lambda: self.update_status_card("UPDATE SUCCESS", "#064e3b", "#10b981"))
                # Delay một chút rồi mới check lại profile để tránh cache
                self.after(2000, self.start_check_thread)
            else:
                self.after(0, lambda text=raw_text: messagebox.showerror("Lỗi", f"Cập nhật thất bại: HTTP {response.status_code}\n\nChi tiết:\n{text}"))
                self.after(0, lambda: self.update_status_card("UPDATE FAILED", "#7f1d1d", "#ef4444"))

        except Exception as e:
            self.update_raw_json(str(e))
            self.after(0, lambda err=str(e): messagebox.showerror("Lỗi", f"Lỗi kết nối: {err}"))
            self.after(0, lambda: self.update_status_card("UPDATE ERROR", "#7f1d1d", "#ef4444"))
        finally:
            self.after(0, lambda: self.btn_update.configure(state="normal", text="✏️ UPDATE DATA"))

if __name__ == "__main__":
    app = TokenCheckerApp()
    app.mainloop()
