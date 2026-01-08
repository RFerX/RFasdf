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

# --- CONFIGURASI TEMA V1 STYLE ---
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
        
        # Color Palette V1 Neon
        self.color_bg = "#0A0A0B"
        self.color_card = "#141417"
        self.color_main = "#A855F7"  
        self.color_main_hover = "#9333EA"
        self.color_dark_purple = "#2E1065"
        self.color_accent = "#00F5FF" 
        
        self.configure(fg_color=self.color_bg)
        self.bots = {} 
        self.col_weights = [12, 12, 5, 12, 12, 15, 25] 

        self.setup_ui()
        self.refresh_link_list()
        self.refresh_config_list()
        
        self.after(500, self.animate_placeholder_loop)
        self.add_log("System initialized. Welcome and be smart.", "SYSTEM", "blue")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def animate_placeholder_loop(self):
        full_text = "Input Domain URL ........."
        def type_text(index):
            if not self.global_domain.get():
                if index <= len(full_text):
                    self.global_domain.configure(placeholder_text=full_text[:index])
                    self.after(100, lambda: type_text(index + 1))
                else:
                    self.after(2000, lambda: erase_text(len(full_text)))
        def erase_text(index):
            if not self.global_domain.get():
                if index >= 0:
                    self.global_domain.configure(placeholder_text=full_text[:index])
                    self.after(50, lambda: erase_text(index - 1))
                else:
                    self.after(500, lambda: type_text(0))
            else:
                self.after(3000, lambda: type_text(0))
        type_text(0)

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

    def setup_ui(self):
        self.header = ctk.CTkFrame(self, height=140, fg_color=self.color_card, corner_radius=0, border_width=1, border_color="#27272A")
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)
        logo_f = ctk.CTkFrame(self.header, fg_color="transparent")
        logo_f.pack(side="left", padx=50)
        ctk.CTkLabel(logo_f, text="Automation Bot V1", font=("Impact", 56), text_color=self.color_main).pack(side="top", anchor="w")
        ctk.CTkLabel(logo_f, text="By RFer-X Garasi", font=("Consolas", 13, "bold"), text_color=self.color_accent).pack(side="top", anchor="w", padx=5)
        domain_container = ctk.CTkFrame(self.header, fg_color=self.color_bg, corner_radius=12, border_width=0, height=60, width=550)
        domain_container.pack(side="right", padx=50, pady=40)
        domain_container.pack_propagate(False)
        domain_container.grid_columnconfigure(1, weight=1)
        domain_container.grid_rowconfigure(0, weight=1)
        ctk.CTkLabel(domain_container, text="🌐", font=("Inter", 20)).grid(row=0, column=0, padx=(20, 5))
        self.global_domain = ctk.CTkEntry(domain_container, width=500, height=55, fg_color="transparent", border_width=0, font=("Consolas", 16, "bold"), text_color=self.color_main, placeholder_text="")
        self.global_domain.grid(row=0, column=1, padx=10, sticky="ew")

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

        self.setup_dashboard()
        self.setup_link_tab()
        self.setup_config_tab()
        self.setup_running_tab()

    def setup_dashboard(self):
        self.log_box = ctk.CTkTextbox(self.tab_dash, fg_color="#050505", font=("Consolas", 14), border_width=1)
        self.log_box.pack(fill="both", expand=True, padx=15, pady=15)

    def setup_link_tab(self):
        container = ctk.CTkFrame(self.tab_link, fg_color="transparent"); container.pack(fill="both", expand=True, padx=30, pady=20)
        input_card = ctk.CTkFrame(container, fg_color="#1C1C20", border_width=1, border_color="#27272A"); input_card.pack(fill="x", pady=20, padx=20)
        self.link_name = ctk.CTkEntry(input_card, width=250, placeholder_text="Bank Name (e.g. BCA_01)", height=40); self.link_name.grid(row=0, column=0, padx=10, pady=20)
        self.link_url = ctk.CTkEntry(input_card, width=500, placeholder_text="Full Google Sheet URL", height=40); self.link_url.grid(row=0, column=1, padx=10, pady=20)
        self.btn_save_link = ctk.CTkButton(input_card, text="ADD SOURCE", fg_color=self.color_main, height=40, command=self.save_link_json, state="disabled"); self.btn_save_link.grid(row=0, column=2, padx=10, pady=20)
        
        self.link_name.bind("<KeyRelease>", lambda e: self.check_link_inputs())
        self.link_url.bind("<KeyRelease>", lambda e: self.check_link_inputs())
        
        self.link_list_frame = ctk.CTkScrollableFrame(container, fg_color="transparent"); self.link_list_frame.pack(fill="both", expand=True, pady=20)

    def setup_config_tab(self):
        container = ctk.CTkFrame(self.tab_config, fg_color="transparent"); container.pack(fill="both", expand=True, padx=30, pady=20)
        input_grid = ctk.CTkFrame(container, fg_color="#1C1C20", border_width=1, border_color="#27272A"); input_grid.pack(fill="x", padx=20, pady=10)
        self.cfg_entries = {}
        labels = ["Name Col", "Nominal Col", "Username Col", "Status Col", "Max", "Timeout (m)", "DupTime (m)"]
        for i, label in enumerate(labels):
            r, c = (0, i) if i < 4 else (2, i-4)
            ctk.CTkLabel(input_grid, text=label.upper(), font=("Inter", 10, "bold"), text_color="gray").grid(row=r, column=c, padx=15, pady=(15,0), sticky="w")
            en = ctk.CTkEntry(input_grid, width=160, height=35); en.grid(row=r+1, column=c, padx=15, pady=(0, 15), sticky="w")
            en.bind("<KeyRelease>", lambda e: self.check_cfg_inputs())
            self.cfg_entries[label] = en
        self.btn_save_cfg = ctk.CTkButton(input_grid, text="SAVE ENGINE RULES", fg_color=self.color_main, width=200, height=40, command=self.save_cfg_json, state="disabled"); self.btn_save_cfg.grid(row=3, column=3, pady=20, padx=15, sticky="e")
        self.cfg_list_frame = ctk.CTkScrollableFrame(container, fg_color="transparent"); self.cfg_list_frame.pack(fill="both", expand=True, pady=10)

    def setup_running_tab(self):
        top = ctk.CTkFrame(self.tab_run, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="+ DEPLOY NEW BOT", fg_color=self.color_main, hover_color=self.color_main_hover, height=45, font=("Inter", 13, "bold"), command=self.add_bot_row).pack(side="left")
        ctk.CTkButton(top, text="Refresh All Assets", fg_color=self.color_dark_purple, width=150, height=45, command=self.refresh_all_bot_dropdowns).pack(side="right")
        self.h_row = ctk.CTkFrame(self.tab_run, fg_color="#1F1F23", height=45)
        self.h_row.pack(fill="x", padx=20, pady=(10, 0))
        headers = ["IDENTIFIER", "SHEET", "ROW", "CONFIG", "DATA SOURCE", "JSON KEY", "COMMANDS"]
        for i, txt in enumerate(headers):
            self.h_row.grid_columnconfigure(i, weight=self.col_weights[i])
            lbl = ctk.CTkLabel(self.h_row, text=txt, font=("Inter", 10, "bold"), text_color="gray")
            lbl.grid(row=0, column=i, padx=5, sticky="nsew")
        self.run_container = ctk.CTkScrollableFrame(self.tab_run, fg_color="transparent")
        self.run_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def add_bot_row(self):
        rid = f"Bot_{int(time.time() * 1000)}"
        row_card = ctk.CTkFrame(self.run_container, fg_color="#16161A", border_color="#27272A", border_width=1)
        row_card.pack(fill="x", pady=5)
        for i, w in enumerate(self.col_weights): row_card.grid_columnconfigure(i, weight=w)
        
        n_en = ctk.CTkEntry(row_card, height=35, fg_color="#0A0A0B"); n_en.grid(row=0, column=0, padx=8, pady=12, sticky="ew")
        s_en = ctk.CTkEntry(row_card, height=35, fg_color="#0A0A0B"); s_en.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        r_en = ctk.CTkEntry(row_card, height=35, fg_color="#0A0A0B"); r_en.insert(0, "2"); r_en.grid(row=0, column=2, padx=8, pady=12, sticky="ew")
        
        cfg_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("cfg_")]), height=35, fg_color="#1F1F23", command=lambda x, r=rid: self.lock_logic(r))
        cfg_dd.set("Select"); cfg_dd.grid(row=0, column=3, padx=8, pady=12, sticky="ew")
        lnk_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("link_")]), height=35, fg_color="#1F1F23", command=lambda x, r=rid: self.lock_logic(r))
        lnk_dd.set("Select"); lnk_dd.grid(row=0, column=4, padx=8, pady=12, sticky="ew")
        
        j_f = ctk.CTkFrame(row_card, fg_color="transparent")
        j_f.grid(row=0, column=5, padx=8, pady=12, sticky="ew")
        j_en = ctk.CTkEntry(j_f, height=35, fg_color="#0A0A0B")
        j_en.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(j_f, text="📂", width=35, height=35, fg_color="#27272A", command=lambda r=rid: self.browse_json_path(r)).pack(side="left", padx=(2,0))
        
        btn_c = ctk.CTkFrame(row_card, fg_color="transparent")
        btn_c.grid(row=0, column=6, padx=8, pady=12, sticky="ew")
        b_web = ctk.CTkButton(btn_c, text="WEB", width=65, height=35, fg_color="#52525B", font=("Inter", 10, "bold"), state="disabled", command=lambda r=rid: self.bot_open_ui(r))
        b_web.pack(side="left", padx=2)
        b_start = ctk.CTkButton(btn_c, text="START", width=60, height=35, fg_color="#166534", text_color="#BBF7D0", font=("Inter", 10, "bold"), state="disabled", command=lambda r=rid: self.bot_start_ui(r))
        b_start.pack(side="left", padx=2)
        b_stop = ctk.CTkButton(btn_c, text="STOP", width=60, height=35, fg_color="#991B1B", text_color="#FECACA", font=("Inter", 10, "bold"), state="disabled", command=lambda r=rid: self.bot_stop_ui(r))
        b_stop.pack(side="left", padx=2)
        ctk.CTkButton(btn_c, text="×", width=35, height=35, fg_color="transparent", hover_color="#EF4444", text_color="gray", font=("Inter", 16), command=lambda r=rid, f=row_card: self.bot_del(r, f)).pack(side="left", padx=2)
        
        status_lbl = ctk.CTkLabel(self.status_container, text=f"IDLE ➜ {rid}", font=("Consolas", 11), text_color="#52525B", anchor="w")
        status_lbl.pack(fill="x", padx=15, pady=8)
        
        self.bots[rid] = {
            'n_en': n_en, 's_en': s_en, 'r_en': r_en, 'cfg_dd': cfg_dd, 'lnk_dd': lnk_dd, 'j_en': j_en, 
            'b_web': b_web, 'b_start': b_start, 'b_stop': b_stop, 'status_lbl': status_lbl, 
            'driver': None, 'is_running': False, 'browser_ready': False,
            'timeout_tracker': {}, 'last_processed': {}
        }
        n_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r))
        s_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r))
        j_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r))

    # --- LOGIKA LIST DENGAN DETAIL LENGKAP ---
    def refresh_config_list(self):
        for w in self.cfg_list_frame.winfo_children(): w.destroy()
        for fn in sorted(os.listdir()):
            if fn.startswith("cfg_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.cfg_list_frame, fg_color="#1C1C20", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=4, padx=10)
                    detail_txt = (
                        f"📄 File: {fn}\n"
                        f"• Cols: [Name: {d.get('Name Col')}] [Nominal: {d.get('Nominal Col')}] "
                        f"[User: {d.get('Username Col')}] [Stat: {d.get('Status Col')}]\n"
                        f"• Rules: Max: {d.get('Max')} | Timeout: {d.get('Timeout (m)')}m | Dup: {d.get('DupTime (m)')}m"
                    )
                    ctk.CTkLabel(c, text=detail_txt, font=("Arial", 11), justify="left").pack(side="left", padx=15, pady=10)
                    ctk.CTkButton(c, text="🗑", width=40, height=40, fg_color="#450a0a", text_color="#f87171", command=lambda f=fn: [os.remove(f), self.refresh_config_list(), self.refresh_all_bot_dropdowns()]).pack(side="right", padx=15)
                except: continue

    def refresh_link_list(self):
        for w in self.link_list_frame.winfo_children(): w.destroy()
        for fn in sorted(os.listdir()):
            if fn.startswith("link_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.link_list_frame, fg_color="#1C1C20", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=2, padx=10)
                    detail_txt = f"🔗 Nama: {d.get('name')} | URL: {d.get('url')[:80]}..."
                    ctk.CTkLabel(c, text=detail_txt, font=("Arial", 11)).pack(side="left", padx=15, pady=8)
                    ctk.CTkButton(c, text="🗑", width=35, height=35, fg_color="#450a0a", command=lambda f=fn: [os.remove(f), self.refresh_link_list(), self.refresh_all_bot_dropdowns()]).pack(side="right", padx=10)
                except: continue

    # --- LOCK LOGIC BUTTONS ---
    def check_link_inputs(self): 
        ready = self.link_name.get().strip() and self.link_url.get().strip()
        self.btn_save_link.configure(state="normal" if ready else "disabled")

    def check_cfg_inputs(self): 
        ready = all(v.get().strip() for v in self.cfg_entries.values())
        self.btn_save_cfg.configure(state="normal" if ready else "disabled")

    def lock_logic(self, rid):
        b = self.bots[rid]
        if b['browser_ready']: return 
        ready = all([b['n_en'].get().strip(), b['s_en'].get().strip(), b['j_en'].get().strip(), 
                     b['cfg_dd'].get() != "Select", b['lnk_dd'].get() != "Select"])
        b['b_web'].configure(state="normal" if ready else "disabled", fg_color=self.color_main if ready else "#52525B")

    # --- CORE BOT LOGIC ---
    def main_logic(self, rid):
        b = self.bots[rid]; bot_name = b['n_en'].get()
        try:
            if b['driver']:
                url = self.global_domain.get().strip()
                if url:
                    target = f"{url.rstrip('/')}/_SubAg_Sub/DepositRequest.aspx?role=sa&userName=al3"
                    b['driver'].get(target if target.startswith("http") else "https://"+target); time.sleep(3)
            
            creds = ServiceAccountCredentials.from_json_keyfile_name(b['j_en'].get(), ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
            client = gspread.authorize(creds)
            with open(b['lnk_dd'].get(), "r") as f: lnk = json.load(f)
            sheet = client.open_by_url(lnk['url']).worksheet(b['s_en'].get())
            self.add_log("Bot Monitoring Aktif.", bot_name, "green")
            
            while b['is_running']:
                try:
                    with open(b['cfg_dd'].get(), "r") as f: cfg = json.load(f)
                    idx_map = {k: self.col_to_idx(cfg[k]) for k in ['Name Col', 'Nominal Col', 'Username Col', 'Status Col']}
                    max_nom = int(re.sub(r'[^\d]', '', cfg.get('Max', '0')))
                    dup_min, to_min = int(cfg.get('DupTime (m)', 2)), int(cfg.get('Timeout (m)', 10))
                    
                    all_rows = sheet.get_all_values(); start_row = int(b['r_en'].get())
                    now = time.time(); pending_queue = []; updates = []
                    
                    self.add_log(f"Mencari data pending (Baris {start_row}+)", bot_name, "blue")
                    for i, row in enumerate(all_rows[start_row-1:], start=start_row):
                        if not b['is_running']: break
                        if len(row) <= max(idx_map.values()): continue
                        n, nom, u, s = row[idx_map['Name Col']].strip(), row[idx_map['Nominal Col']].strip(), row[idx_map['Username Col']].strip(), row[idx_map['Status Col']].strip()
                        
                        if n and nom and not u and not s:
                            nc = "".join(filter(str.isdigit, re.split(r'[.,]\d{2}$', nom)[0]))
                            if not nc: continue
                            if max_nom > 0 and int(nc) > max_nom:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["❌"]]}); continue
                            dk = f"{n.lower()}_{nc}"
                            if dk in b['last_processed'] and (now - b['last_processed'][dk])/60 < dup_min:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["⚠️"]]}); continue
                            rk = f"row_{i}_{n}"
                            if rk not in b['timeout_tracker']: b['timeout_tracker'][rk] = now
                            if (now - b['timeout_tracker'][rk])/60 > to_min:
                                self.add_log(f"EXPIRED: {n}", bot_name, "red")
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["❌"]]}); continue
                            pending_queue.append({"row": i, "nama": n, "nominal": nc, "dk": dk, "rk": rk})
                    
                    if not pending_queue: self.add_log("Scan Selesai: Tidak ada pendingan.", bot_name, "blue")
                    else:
                        self.add_log(f"Mendeteksi {len(pending_queue)} data...", bot_name, "orange")
                        try: b['driver'].find_element(By.ID, "btnRefresh").click(); time.sleep(1.5)
                        except: b['driver'].refresh(); time.sleep(3)
                        for item in pending_queue:
                            if not b['is_running']: break
                            res = self.cari_dan_klik_web(b['driver'], item["nama"], item["nominal"])
                            if res:
                                self.add_log(f"MATCH: {item['nama']} Berhasil!", bot_name, "green")
                                updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], idx_map['Status Col'] + 1), 'values': [["✅"]]})
                                updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], idx_map['Username Col'] + 1), 'values': [[res]]})
                                b['last_processed'][item["dk"]] = time.time()
                            else: self.add_log(f"NOT FOUND: {item['nama']}", bot_name, "red")
                    if updates: sheet.batch_update(updates)
                    time.sleep(3)
                except Exception as e: self.add_log(f"Err: {str(e)[:30]}", bot_name, "red"); time.sleep(5)
        except Exception as e: self.add_log(f"Fatal: {str(e)[:30]}", bot_name, "red"); self.bot_stop_ui(rid)

    def cari_dan_klik_web(self, driver, nama_gs, nominal_gs_string):
        try:
            nama_gs_bersih = re.sub(r'[^a-z0-9]', '', nama_gs.lower())
            rows = driver.find_elements(By.CSS_SELECTOR, "table#report tbody tr[class^='Grid']")
            for row in rows:
                try:
                    name_web = row.find_element(By.CLASS_NAME, "fromAccountName").text.strip()
                    amount_web_clean = "".join(filter(str.isdigit, row.find_element(By.CLASS_NAME, "amount").text.strip().split('.')[0]))
                    if nominal_gs_string == amount_web_clean and nama_gs_bersih == re.sub(r'[^a-z0-9]', '', name_web.lower()):
                        username_web = row.find_element(By.CLASS_NAME, "username").text.strip()
                        btn = row.find_element(By.CLASS_NAME, "confirm").find_element(By.TAG_NAME, "input")
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn); time.sleep(0.5); btn.click()
                        WebDriverWait(driver, 5).until(EC.alert_is_present()); driver.switch_to.alert.accept(); return username_web
                except: continue
        except: pass
        return None

    def save_cfg_json(self):
        d = {k: v.get().strip() for k, v in self.cfg_entries.items()}
        # GABUNGAN 4 ISI TEXTBOX UNTUK NAMA FILE
        raw_name = f"{d['Name Col']}_{d['Nominal Col']}_{d['Username Col']}_{d['Status Col']}"
        safe_name = re.sub(r'[^a-zA-Z0-9]', '', raw_name)
        filename = f"cfg_{safe_name}.json"
        
        with open(filename, "w") as j: json.dump(d, j, indent=4)
        for e in self.cfg_entries.values(): e.delete(0, 'end')
        self.btn_save_cfg.configure(state="disabled")
        self.refresh_config_list()
        self.refresh_all_bot_dropdowns()

    def save_link_json(self):
        link_data = {"name": self.link_name.get().strip(), "url": self.link_url.get().strip()}
        with open(f"link_{link_data['name']}.json", "w") as j: json.dump(link_data, j, indent=4)
        self.link_name.delete(0, 'end'); self.link_url.delete(0, 'end'); self.btn_save_link.configure(state="disabled")
        self.refresh_link_list(); self.refresh_all_bot_dropdowns()

    def browse_json_path(self, rid):
        fn = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if fn: self.bots[rid]['j_en'].delete(0, 'end'); self.bots[rid]['j_en'].insert(0, fn); self.lock_logic(rid)

    def bot_open_ui(self, rid):
        b = self.bots[rid]; b['b_web'].configure(state="disabled")
        b['status_lbl'].configure(text=f"LAUNCHING ➜ {b['n_en'].get()}", text_color=self.color_accent)
        threading.Thread(target=self.open_browser_task, args=(rid,), daemon=True).start()

    def open_browser_task(self, rid):
        b = self.bots[rid]
        try:
            url = self.global_domain.get().strip()
            if not url: return
            b['driver'] = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            b['driver'].get(url if url.startswith("http") else "https://"+url)
            b['browser_ready'] = True
            self.after(0, lambda: [b['status_lbl'].configure(text=f"READY ➜ {b['n_en'].get()}", text_color=self.color_accent), b['b_start'].configure(state="normal")])
            while True: time.sleep(1); b['driver'].window_handles
        except:
            b['driver'] = None; b['is_running'] = False; b['browser_ready'] = False
            self.after(0, lambda: [b['status_lbl'].configure(text=f"IDLE ➜ {rid}", text_color="#52525B"), b['b_web'].configure(state="normal"), b['b_start'].configure(state="disabled")])

    def bot_start_ui(self, rid):
        b = self.bots[rid]; b['is_running'] = True; b['timeout_tracker'] = {}; b['last_processed'] = {}
        b['b_start'].configure(state="disabled"); b['b_stop'].configure(state="normal")
        b['status_lbl'].configure(text=f"RUNNING ➜ {b['n_en'].get()}", text_color="#10B981")
        threading.Thread(target=self.main_logic, args=(rid,), daemon=True).start()

    def bot_stop_ui(self, rid):
        b = self.bots[rid]; b['is_running'] = False; b['b_stop'].configure(state="disabled"); b['b_start'].configure(state="normal")
        b['status_lbl'].configure(text=f"HALTED ➜ {b['n_en'].get()}", text_color="#F59E0B")

    def bot_del(self, rid, frame):
        if self.bots[rid]['driver']: self.bots[rid]['driver'].quit()
        self.bots[rid]['status_lbl'].destroy(); del self.bots[rid]; frame.destroy()

    def refresh_all_bot_dropdowns(self):
        cfgs = ["Select"] + sorted([f for f in os.listdir() if f.startswith("cfg_")])
        lnks = ["Select"] + sorted([f for f in os.listdir() if f.startswith("link_")])
        for b in self.bots.values(): b['cfg_dd'].configure(values=cfgs); b['lnk_dd'].configure(values=lnks)

    def col_to_idx(self, letter):
        idx = 0
        for c in letter.upper().strip(): idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1
    def on_closing(self): self.destroy(); os._exit(0)

if __name__ == "__main__":
    app = AutomationBotApp()
    app.mainloop()
