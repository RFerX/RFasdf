import customtkinter as ctk
from tkinter import messagebox, filedialog
import json
import os
import threading
import time
import datetime
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURASI TEMA ---
ctk.set_appearance_mode("dark")

class AutomationBotApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Bot - Ultimate Version")
        
        window_width = 1450 
        window_height = 900
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        self.color_bg = "#0A0A0B"
        self.color_card = "#141417"
        self.color_main = "#A855F7"  
        self.color_accent = "#00F5FF" 
        self.color_error = "#EF4444"
        
        self.configure(fg_color=self.color_bg)
        self.bots = {} 
        self.col_weights = [12, 12, 5, 12, 12, 15, 25] 

        self.setup_ui()
        self.refresh_link_list()
        self.refresh_config_list()
        self.load_session() 
        
        self.add_log("System initialized. Welcome and be Smart.", "SYSTEM", "blue")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    # --- LOG SYSTEM ---
    def add_log(self, message, bot_display_name="SYSTEM", color="green"):
        self.after(0, self._process_log, message, bot_display_name, color)

    def _process_log(self, message, bot_display_name, color):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_message = f" [{timestamp}] [{bot_display_name}] ➜ {message}\n"
        color_map = {"green": "#10B981", "blue": "#3B82F6", "orange": "#F59E0B", "red": "#EF4444"}
        tag_name = f"color_{color}"
        self.log_box.tag_config(tag_name, foreground=color_map.get(color, "#10B981"))
        self.log_box.insert("end", full_message, tag_name)
        self.log_box.see("end")

    # --- VALIDASI INPUT ---
    def check_cfg_inputs(self):
        all_valid = True
        text_cols = ["Name Col", "Nominal Col", "Username Col", "Status Col"]
        num_cols = ["Max", "Timeout (m)", "DupTime (m)"]
        for label in text_cols:
            en = self.cfg_entries[label]
            val = en.get().strip()
            if val and re.match(r"^[a-zA-Z]+$", val):
                en.configure(text_color="white")
            else:
                en.configure(text_color=self.color_error)
                all_valid = False
        for label in num_cols:
            en = self.cfg_entries[label]
            val = en.get().strip()
            if val and val.isdigit():
                en.configure(text_color="white")
            else:
                en.configure(text_color=self.color_error)
                all_valid = False
        self.btn_save_cfg.configure(state="normal" if all_valid else "disabled", fg_color=self.color_main if all_valid else "#52525B")

    # --- UI SETUP ---
    def setup_ui(self):
        self.header = ctk.CTkFrame(self, height=140, fg_color=self.color_card, corner_radius=0, border_width=1, border_color="#27272A")
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)
        
        logo_f = ctk.CTkFrame(self.header, fg_color="transparent")
        logo_f.pack(side="left", padx=50)
        ctk.CTkLabel(logo_f, text="Automation Bot V1", font=("Impact", 56), text_color=self.color_main).pack(side="top", anchor="w")
        ctk.CTkLabel(logo_f, text="By RFer-X Garasi", font=("Consolas", 13, "bold"), text_color=self.color_accent).pack(side="top", anchor="w", padx=5)
        
        domain_container = ctk.CTkFrame(self.header, fg_color=self.color_bg, corner_radius=12, height=60, width=550)
        domain_container.pack(side="right", padx=50, pady=40); domain_container.pack_propagate(False)
        domain_container.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(domain_container, text="🌐", font=("Inter", 20)).grid(row=0, column=0, padx=(20, 5))
        self.global_domain = ctk.CTkEntry(domain_container, width=500, height=55, fg_color="transparent", border_width=0, font=("Consolas", 16, "bold"), text_color=self.color_main, placeholder_text="Input Domain URL.........")
        self.global_domain.grid(row=0, column=1, padx=10, sticky="ew")
        self.global_domain.bind("<KeyRelease>", lambda e: self.update_all_locks())

        self.sidebar = ctk.CTkFrame(self, width=280, fg_color=self.color_card, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        ctk.CTkLabel(self.sidebar, text="BOTS STATUS", font=("Inter", 12, "bold"), text_color=self.color_main).pack(pady=(25, 10), padx=20, anchor="w")
        self.status_container = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.status_container.pack(fill="both", expand=True, padx=5, pady=5)

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(side="right", expand=True, fill="both", padx=15, pady=15)
        self.tabview = ctk.CTkTabview(content, segmented_button_selected_color=self.color_main, fg_color=self.color_card)
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_dash = self.tabview.add("LIVE MONITOR")
        self.tab_link = self.tabview.add("CONNECT SHEETS")
        self.tab_config = self.tabview.add("RULES ENGINE")
        self.tab_run = self.tabview.add("BOT INSTANCES")

        self.setup_dashboard(); self.setup_link_tab(); self.setup_config_tab(); self.setup_running_tab()

    def setup_dashboard(self):
        self.log_box = ctk.CTkTextbox(self.tab_dash, fg_color="#050505", font=("Consolas", 14), border_width=1)
        self.log_box.pack(fill="both", expand=True, padx=15, pady=15)

    def setup_link_tab(self):
        container = ctk.CTkFrame(self.tab_link, fg_color="transparent"); container.pack(fill="both", expand=True, padx=30, pady=20)
        top_bar = ctk.CTkFrame(container, fg_color="transparent"); top_bar.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(top_bar, text="🔄 Refresh", width=120, fg_color="#2E1065", command=self.refresh_link_list).pack(side="right")
        input_card = ctk.CTkFrame(container, fg_color="#1C1C20", border_width=1, border_color="#27272A"); input_card.pack(fill="x", pady=10, padx=20)
        self.link_name = ctk.CTkEntry(input_card, width=250, placeholder_text="Sheet Category", height=40); self.link_name.grid(row=0, column=0, padx=10, pady=20)
        self.link_url = ctk.CTkEntry(input_card, width=500, placeholder_text="Full Google Sheet URL", height=40); self.link_url.grid(row=0, column=1, padx=10, pady=20)
        self.btn_save_link = ctk.CTkButton(input_card, text="SAVE LINK", fg_color=self.color_main, height=40, command=self.save_link_json, state="disabled"); self.btn_save_link.grid(row=0, column=2, padx=10, pady=20)
        self.link_name.bind("<KeyRelease>", lambda e: self.check_link_inputs()); self.link_url.bind("<KeyRelease>", lambda e: self.check_link_inputs())
        self.link_list_frame = ctk.CTkScrollableFrame(container, fg_color="transparent"); self.link_list_frame.pack(fill="both", expand=True, pady=20)

    def setup_config_tab(self):
        container = ctk.CTkFrame(self.tab_config, fg_color="transparent"); container.pack(fill="both", expand=True, padx=30, pady=20)
        top_bar = ctk.CTkFrame(container, fg_color="transparent"); top_bar.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(top_bar, text="🔄 Refresh", width=120, fg_color="#2E1065", command=self.refresh_config_list).pack(side="right")
        input_grid = ctk.CTkFrame(container, fg_color="#1C1C20", border_width=1, border_color="#27272A"); input_grid.pack(fill="x", padx=20, pady=10)
        self.cfg_entries = {}
        labels = ["Name Col", "Nominal Col", "Username Col", "Status Col", "Max", "Timeout (m)", "DupTime (m)"]
        for i, label in enumerate(labels):
            r, c = (0, i) if i < 4 else (2, i-4)
            ctk.CTkLabel(input_grid, text=label.upper(), font=("Inter", 10, "bold"), text_color="gray").grid(row=r, column=c, padx=15, pady=(15,0), sticky="w")
            en = ctk.CTkEntry(input_grid, width=160, height=35)
            en.grid(row=r+1, column=c, padx=15, pady=(0, 15), sticky="w")
            en.bind("<KeyRelease>", lambda e: self.check_cfg_inputs())
            self.cfg_entries[label] = en
        self.btn_save_cfg = ctk.CTkButton(input_grid, text="SAVE RULE", fg_color=self.color_main, width=200, height=40, command=self.save_cfg_json, state="disabled"); self.btn_save_cfg.grid(row=3, column=3, pady=20, padx=15, sticky="e")
        self.cfg_list_frame = ctk.CTkScrollableFrame(container, fg_color="transparent"); self.cfg_list_frame.pack(fill="both", expand=True, pady=10)

    def setup_running_tab(self):
        top = ctk.CTkFrame(self.tab_run, fg_color="transparent"); top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="+ DEPLOY NEW BOT", fg_color=self.color_main, height=45, font=("Inter", 13, "bold"), command=self.add_bot_row).pack(side="left")
        ctk.CTkButton(top, text="Refresh", fg_color="#2E1065", width=150, height=45, command=self.refresh_all_bot_dropdowns).pack(side="right")
        self.h_row = ctk.CTkFrame(self.tab_run, fg_color="#1F1F23", height=45); self.h_row.pack(fill="x", padx=20, pady=(10, 0))
        headers = ["IDENTIFIER", "SHEET", "ROW", "CONFIG", "DATA SOURCE", "JSON KEY", "COMMANDS"]
        for i, txt in enumerate(headers):
            self.h_row.grid_columnconfigure(i, weight=self.col_weights[i])
            ctk.CTkLabel(self.h_row, text=txt, font=("Inter", 10, "bold"), text_color="gray").grid(row=0, column=i, padx=5, sticky="nsew")
        self.run_container = ctk.CTkScrollableFrame(self.tab_run, fg_color="transparent"); self.run_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def add_bot_row(self, saved_info=None):
        rid = f"Bot_{int(time.time() * 1000)}"
        row_card = ctk.CTkFrame(self.run_container, fg_color="#16161A", border_color="#27272A", border_width=1); row_card.pack(fill="x", pady=5)
        for i, w in enumerate(self.col_weights): row_card.grid_columnconfigure(i, weight=w)
        
        n_en = ctk.CTkEntry(row_card, height=35); n_en.grid(row=0, column=0, padx=8, pady=12, sticky="ew")
        s_en = ctk.CTkEntry(row_card, height=35); s_en.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        r_en = ctk.CTkEntry(row_card, height=35); r_en.insert(0, "2"); r_en.grid(row=0, column=2, padx=8, pady=12, sticky="ew")
        cfg_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("cfg_")]), height=35); cfg_dd.set("Select"); cfg_dd.grid(row=0, column=3, padx=8, pady=12, sticky="ew")
        lnk_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("link_")]), height=35); lnk_dd.set("Select"); lnk_dd.grid(row=0, column=4, padx=8, pady=12, sticky="ew")
        j_f = ctk.CTkFrame(row_card, fg_color="transparent"); j_f.grid(row=0, column=5, padx=8, pady=12, sticky="ew")
        j_en = ctk.CTkEntry(j_f, height=35); j_en.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(j_f, text="📂", width=35, height=35, command=lambda r=rid: self.browse_json_path(r)).pack(side="left", padx=(2,0))
        
        btn_c = ctk.CTkFrame(row_card, fg_color="transparent"); btn_c.grid(row=0, column=6, padx=8, pady=12, sticky="ew")
        b_web = ctk.CTkButton(btn_c, text="WEB", width=65, state="disabled", command=lambda r=rid: self.bot_open_ui(r)); b_web.pack(side="left", padx=2)
        b_start = ctk.CTkButton(btn_c, text="START", width=60, state="disabled", fg_color="#166534", command=lambda r=rid: self.bot_start_ui(r)); b_start.pack(side="left", padx=2)
        b_stop = ctk.CTkButton(btn_c, text="STOP", width=60, state="disabled", fg_color="#991B1B", command=lambda r=rid: self.bot_stop_ui(r)); b_stop.pack(side="left", padx=2)
        ctk.CTkButton(btn_c, text="×", width=35, fg_color="transparent", command=lambda r=rid, f=row_card: self.bot_del(r, f)).pack(side="left", padx=2)
        
        status_lbl = ctk.CTkLabel(self.status_container, text=f"IDLE ➜ {rid}", font=("Consolas", 17), text_color="#52525B", anchor="w")
        status_lbl.pack(fill="x", padx=15, pady=8)
        
        self.bots[rid] = {'n_en': n_en, 's_en': s_en, 'r_en': r_en, 'cfg_dd': cfg_dd, 'lnk_dd': lnk_dd, 'j_en': j_en, 'b_web': b_web, 'b_start': b_start, 'b_stop': b_stop, 'status_lbl': status_lbl, 'driver': None, 'is_running': False, 'timeout_tracker': {}, 'last_processed': {}}
        if saved_info:
            n_en.insert(0, saved_info.get('identifier', '')); s_en.insert(0, saved_info.get('sheet', '')); r_en.delete(0, 'end'); r_en.insert(0, saved_info.get('row', '2'))
            if saved_info.get('config') in cfg_dd.cget("values"): cfg_dd.set(saved_info['config'])
            if saved_info.get('link') in lnk_dd.cget("values"): lnk_dd.set(saved_info['link'])
            j_en.insert(0, saved_info.get('json_path', ''))

        n_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); s_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); j_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); self.lock_logic(rid)

    # --- BOT CORE ---
    def main_logic(self, rid):
        b = self.bots[rid]
        bot_name = b['n_en'].get()
        try:
            self.add_log("------------ MENGAMBIL DATA DARI SHEET...", bot_name, "blue")
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(b['j_en'].get(), scope)
            client = gspread.authorize(creds)
            with open(b['lnk_dd'].get(), "r") as f: lnk = json.load(f)
            sheet = client.open_by_url(lnk['url']).worksheet(b['s_en'].get())

            while b['is_running']:
                with open(b['cfg_dd'].get(), "r") as f: cfg = json.load(f)
                idx_map = {k: self.col_to_idx(cfg[k]) for k in ['Name Col', 'Nominal Col', 'Username Col', 'Status Col']}
                max_nom = int(re.sub(r'[^\d]', '', cfg.get('Max', '0')))
                dup_min = int(cfg.get('DupTime (m)', 2))
                to_min = int(cfg.get('Timeout (m)', 10))
                
                all_rows = sheet.get_all_values()
                start_row = int(b['r_en'].get())
                now = time.time()
                pending_queue = []; updates = []

                for i, row in enumerate(all_rows[start_row-1:], start=start_row):
                    if not b['is_running']: break
                    if len(row) <= max(idx_map.values()): continue
                    
                    nama_gs = row[idx_map['Name Col']].strip()
                    nom_raw = row[idx_map['Nominal Col']].strip()
                    user_gs = row[idx_map['Username Col']].strip()
                    status_gs = row[idx_map['Status Col']].strip()

                    if nama_gs and nom_raw and not user_gs and not status_gs:
                        nom_clean = "".join(filter(str.isdigit, re.split(r'[.,]\d{2}$', nom_raw)[0]))
                        if not nom_clean: continue
                        if max_nom > 0 and int(nom_clean) > max_nom:
                            self.add_log(f"--- MAX REACHED: {nama_gs}", bot_name, "orange")
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["❌"]]})
                            continue
                        dup_key = f"{nama_gs.lower()}_{nom_clean}"
                        if dup_key in b['last_processed'] and (now - b['last_processed'][dup_key])/60 < dup_min:
                            self.add_log(f"--- DUP SKIP: {nama_gs}", bot_name, "orange")
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["⚠️"]]})
                            continue
                        r_key = f"row_{i}_{nama_gs}"
                        if r_key not in b['timeout_tracker']: b['timeout_tracker'][r_key] = now
                        if (now - b['timeout_tracker'][r_key])/60 > to_min:
                            self.add_log(f"--- TIMEOUT: {nama_gs}", bot_name, "orange")
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["❌"]]})
                            continue
                        pending_queue.append({"row": i, "nama": nama_gs, "nominal": nom_clean, "dup_key": dup_key, "r_key": r_key})

                if pending_queue:
                    self.add_log(f"--- {len(pending_queue)} NEW DATA", bot_name, "blue")
                    try: b['driver'].find_element(By.ID, "btnRefresh").click(); time.sleep(1.5)
                    except: b['driver'].refresh(); time.sleep(3)
                    for item in pending_queue:
                        if not b['is_running']: break
                        res_user = self.cari_dan_klik_web(b['driver'], item["nama"], item["nominal"])
                        if res_user:
                            self.add_log(f"--- SUCCESS: {item['nama']}", bot_name, "green")
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], idx_map['Status Col'] + 1), 'values': [["✅"]]})
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], idx_map['Username Col'] + 1), 'values': [[res_user]]})
                            b['last_processed'][item["dup_key"]] = time.time()
                            if item["r_key"] in b['timeout_tracker']: del b['timeout_tracker'][item["r_key"]]

                if updates: sheet.batch_update(updates)
                time.sleep(3)
        except Exception as e:
            self.add_log(f"--- ERROR: {str(e)[:50]}", bot_name, "red")
            self.bot_stop_ui(rid)

    def cari_dan_klik_web(self, driver, nama_gs, nominal_gs_string):
        try:
            nama_gs_bersih = re.sub(r'[^a-z0-9]', '', nama_gs.lower())
            rows = driver.find_elements(By.CSS_SELECTOR, "table#report tbody tr[class^='Grid']")
            for row in rows:
                try:
                    name_web = row.find_element(By.CLASS_NAME, "fromAccountName").text.strip()
                    amount_raw = row.find_element(By.CLASS_NAME, "amount").text.strip()
                    amount_web_clean = "".join(filter(str.isdigit, amount_raw.split('.')[0]))
                    if nominal_gs_string == amount_web_clean and nama_gs_bersih == re.sub(r'[^a-z0-9]', '', name_web.lower()):
                        username_web = row.find_element(By.CLASS_NAME, "username").text.strip()
                        btn = row.find_element(By.CLASS_NAME, "confirm").find_element(By.TAG_NAME, "input")
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn); time.sleep(0.5); btn.click()
                        WebDriverWait(driver, 5).until(EC.alert_is_present()); driver.switch_to.alert.accept()
                        return username_web
                except: continue
        except: pass
        return None

    # --- CONTROLS ---
    def bot_open_ui(self, rid):
        b = self.bots[rid]; b['status_lbl'].configure(text="LAUNCHING...", text_color=self.color_accent)
        threading.Thread(target=self.open_browser_task, args=(rid,), daemon=True).start()

    def open_browser_task(self, rid):
        b = self.bots[rid]
        try:
            url = self.global_domain.get().strip()
            if not url.startswith("http"): url = "https://" + url
            b['driver'] = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            b['driver'].get(url)
            self.after(0, lambda: [b['status_lbl'].configure(text=f"READY ➜ {b['n_en'].get()}", text_color=self.color_accent), b['b_start'].configure(state="normal")])
            while b['driver']: 
                time.sleep(1)
                _ = b['driver'].window_handles
        except: b['driver'] = None
        finally:
            self.after(0, lambda: [b['status_lbl'].configure(text=f"IDLE ➜ {rid}", text_color="#52525B"), b['b_web'].configure(state="normal"), b['b_start'].configure(state="disabled")])

    def bot_start_ui(self, rid):
        b = self.bots[rid]
        # Pindah ke Link Khusus Deposit
        try:
            base_url = self.global_domain.get().strip()
            if base_url.endswith('/'): base_url = base_url[:-1]
            if not base_url.startswith("http"): base_url = "https://" + base_url
            target_url = f"{base_url}/_SubAg_Sub/DepositRequest.aspx?role=sa&userName=al3"
            
            if b['driver']:
                b['driver'].get(target_url)
                self.add_log(f"Navigasi ke: {target_url}", b['n_en'].get(), "blue")
        except Exception as e:
            self.add_log(f"Gagal navigasi: {str(e)}", b['n_en'].get(), "red")
            return

        b['is_running'] = True
        b['timeout_tracker'] = {}; b['last_processed'] = {}
        
        # Update Tombol: WEB MATI, START MATI, STOP NYALA
        b['b_web'].configure(state="disabled", fg_color="#52525B")
        b['b_start'].configure(state="disabled")
        b['b_stop'].configure(state="normal", fg_color="#991B1B")
        
        b['status_lbl'].configure(text=f"RUNNING ➜ {b['n_en'].get()}", text_color="#10B981")
        self.add_log("------------ BOT MULAI BERJALAN", b['n_en'].get(), "blue")
        threading.Thread(target=self.main_logic, args=(rid,), daemon=True).start()

    def bot_stop_ui(self, rid):
        b = self.bots[rid]; b['is_running'] = False
        
        # Update Tombol: WEB NYALA, START NYALA, STOP MATI
        b['b_stop'].configure(state="disabled", fg_color="#52525B")
        b['b_start'].configure(state="normal", fg_color="#166534")
        b['b_web'].configure(state="normal", fg_color=self.color_main)
        
        b['status_lbl'].configure(text=f"STOP ➜ {b['n_en'].get()}", text_color="#F59E0B")
        self.add_log("------------ BOT DIHENTIKAN (STOP)", b['n_en'].get(), "orange")

    def lock_logic(self, rid):
        b = self.bots[rid]; domain = self.global_domain.get().strip()
        ready = all([domain, b['n_en'].get().strip(), b['s_en'].get().strip(), b['j_en'].get().strip(), b['cfg_dd'].get() != "Select", b['lnk_dd'].get() != "Select"])
        b['b_web'].configure(state="normal" if ready else "disabled", fg_color=self.color_main if ready else "#52525B")

    def update_all_locks(self):
        for rid in self.bots: self.lock_logic(rid)

    def col_to_idx(self, letter):
        idx = 0
        for c in letter.upper().strip(): idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1

    # --- UI HELPERS ---
    def browse_json_path(self, r):
        fn = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if fn: self.bots[r]['j_en'].delete(0, 'end'); self.bots[r]['j_en'].insert(0, fn); self.lock_logic(r)

    # --- KEMBALI KE LIST DETAIL RULES ENGINE ---
    def refresh_config_list(self):
        for w in self.cfg_list_frame.winfo_children(): w.destroy()
        h = ctk.CTkFrame(self.cfg_list_frame, fg_color="transparent")
        h.pack(fill="x", pady=(0,5))
        ctk.CTkLabel(h, text="DETAILED RULE CONFIGURATIONS", font=("Inter", 11, "bold"), text_color=self.color_main).pack(side="left", padx=10)
        
        for fn in sorted(os.listdir()):
            if fn.startswith("cfg_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.cfg_list_frame, fg_color="#1C1C20", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=5, padx=5)
                    
                    detail_text = (
                        f"📁 FILE: {fn}\n"
                        f"------------------------------------------------------------------------------------------------------------------------\n"
                        f"📍 KOLOM NAMA: {d.get('Name Col')}  |  📍 KOLOM NOMINAL: {d.get('Nominal Col')}  |  📍 KOLOM USERNAME: {d.get('Username Col')}  |  📍 KOLOM STATUS: {d.get('Status Col')}\n"
                        f"⚙️ BATAS MAKSIMAL: {d.get('Max')} Baris  |  ⏳ TIMEOUT: {d.get('Timeout (m)')} Menit  |  🔄 CEK DUPLIKAT: {d.get('DupTime (m)')} Menit"
                    )
                    
                    ctk.CTkLabel(c, text=detail_text, font=("Consolas", 11), justify="left", text_color="#E2E8F0").pack(side="left", padx=20, pady=15)
                    ctk.CTkButton(c, text="HAPUS", width=100, height=40, fg_color="#450a0a", hover_color="#EF4444", command=lambda f=fn: [os.remove(f), self.refresh_config_list()]).pack(side="right", padx=20)
                except: pass

    def save_cfg_json(self):
        d = {k: v.get().strip() for k, v in self.cfg_entries.items()}
        raw_name = f"{d['Name Col']}_{d['Nominal Col']}_{d['Username Col']}_{d['Status Col']}"
        clean_name = re.sub(r'[^a-zA-Z0-9]', '', raw_name)
        file_name = f"cfg_{clean_name}.json"
        
        with open(file_name, "w") as j: json.dump(d, j, indent=4)
        for e in self.cfg_entries.values(): e.delete(0, 'end')
        self.refresh_config_list(); self.refresh_all_bot_dropdowns(); self.check_cfg_inputs()

    # --- KEMBALI KE LIST DETAIL CONNECT SHEETS ---
    def refresh_link_list(self):
        for w in self.link_list_frame.winfo_children(): w.destroy()
        h = ctk.CTkFrame(self.link_list_frame, fg_color="transparent")
        h.pack(fill="x", pady=(0,5))
        ctk.CTkLabel(h, text="REGISTERED DATA SOURCES", font=("Inter", 11, "bold"), text_color=self.color_main).pack(side="left", padx=10)
        
        for fn in sorted(os.listdir()):
            if fn.startswith("link_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.link_list_frame, fg_color="#1C1C20", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=5, padx=5)
                    
                    detail_text = (
                        f"🏷️ NAMA KATEGORI : {d.get('name')}\n"
                        f"🔗 URL GOOGLE SHEET: {d.get('url')}"
                    )
                    
                    ctk.CTkLabel(c, text=detail_text, font=("Consolas", 11), justify="left", text_color="#E2E8F0").pack(side="left", padx=20, pady=15)
                    ctk.CTkButton(c, text="HAPUS", width=100, height=40, fg_color="#450a0a", hover_color="#EF4444", command=lambda f=fn: [os.remove(f), self.refresh_link_list()]).pack(side="right", padx=20)
                except: pass

    def save_link_json(self):
        d = {"name": self.link_name.get().strip(), "url": self.link_url.get().strip()}
        with open(f"link_{d['name']}.json", "w") as j: json.dump(d, j, indent=4)
        self.link_name.delete(0, 'end'); self.link_url.delete(0, 'end'); self.refresh_link_list(); self.refresh_all_bot_dropdowns()

    def check_link_inputs(self): self.btn_save_link.configure(state="normal" if (self.link_name.get().strip() and self.link_url.get().strip()) else "disabled")
    
    def refresh_all_bot_dropdowns(self):
        cf = ["Select"] + sorted([f for f in os.listdir() if f.startswith("cfg_")])
        ln = ["Select"] + sorted([f for f in os.listdir() if f.startswith("link_")])
        for b in self.bots.values(): b['cfg_dd'].configure(values=cf); b['lnk_dd'].configure(values=ln)

    def save_session(self):
        data = {rid: {"identifier": b['n_en'].get(), "sheet": b['s_en'].get(), "row": b['r_en'].get(), "config": b['cfg_dd'].get(), "link": b['lnk_dd'].get(), "json_path": b['j_en'].get()} for rid, b in self.bots.items()}
        with open("session_bots.json", "w") as f: json.dump(data, f, indent=4)

    def load_session(self):
        if os.path.exists("session_bots.json"):
            try:
                with open("session_bots.json", "r") as f:
                    for rid, info in json.load(f).items(): self.add_bot_row(saved_info=info)
            except: pass

    def bot_del(self, rid, frame):
        if rid in self.bots:
            if self.bots[rid]['driver']:
                try: self.bots[rid]['driver'].quit()
                except: pass
            self.bots[rid]['status_lbl'].destroy(); del self.bots[rid]
        frame.destroy()

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Tutup aplikasi?"):
            for b in self.bots.values():
                if b['driver']: 
                    try: b['driver'].quit()
                    except: pass
            self.save_session(); self.destroy(); os._exit(0)

if __name__ == "__main__":
    app = AutomationBotApp(); app.mainloop()
