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

# --- KONFIGURASI TEMA ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class PurpleBotApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Bot - Pro Version")
        
        window_width = 1250
        window_height = 850
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)
        
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        self.bots = {} 
        self.color_main = "#8e44ad" 
        self.color_dark = "#1e002a" 
        
        self.col_weights = [12, 15, 5, 12, 12, 18, 26]

        self.setup_ui()
        self.refresh_link_list()
        self.refresh_config_list()
        self.add_log("Aplikasi dijalankan. Semua sistem siap!", "SYSTEM", "blue")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def add_log(self, message, bot_display_name="SYSTEM", color="green"):
        self.after(0, self._process_log, message, bot_display_name, color)

    def _process_log(self, message, bot_display_name, color):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] [{bot_display_name}] --- {message}\n"
        color_map = {"green": "#00FFCC", "blue": "#00D4FF", "orange": "#FFCC00", "red": "#FF4B4B"}
        tag_name = f"color_{color}"
        self.log_box.tag_config(tag_name, foreground=color_map.get(color, "#00FF00"))
        self.log_box.insert("end", full_message, tag_name)
        self.log_box.see("end")

    def setup_ui(self):
        header_frame = ctk.CTkFrame(self, fg_color=self.color_dark, height=50, border_color=self.color_main, border_width=1)
        header_frame.pack(fill="x", padx=10, pady=(5, 0))
        
        ctk.CTkLabel(header_frame, text="🌐 URL UTAMA:", font=("Arial", 14, "bold"), text_color="#a29bfe").pack(side="left", padx=15)
        self.global_domain = ctk.CTkEntry(header_frame, width=500, height=30, border_color=self.color_main, placeholder_text="")
        self.global_domain.pack(side="left", padx=5, pady=10)

        self.tabview = ctk.CTkTabview(self, segmented_button_selected_color=self.color_main, segmented_button_unselected_hover_color="#6c5ce7")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        
        self.tab_dash = self.tabview.add("Live Status")
        self.tab_link = self.tabview.add("Connect Sheet")
        self.tab_config = self.tabview.add("Set Rules")
        self.tab_run = self.tabview.add("Launch Bot")

        self.setup_dashboard()
        self.setup_link_tab()
        self.setup_config_tab()
        self.setup_running_tab()

    def setup_dashboard(self):
        monitor_wrapper = ctk.CTkFrame(self.tab_dash, fg_color="transparent")
        monitor_wrapper.pack(fill="x", padx=10, pady=2)
        mon_header = ctk.CTkFrame(monitor_wrapper, fg_color="transparent")
        mon_header.pack(fill="x")
        ctk.CTkLabel(mon_header, text="Bot Status", font=("Arial", 14, "bold"), text_color="#a29bfe").pack(side="left")
        ctk.CTkButton(mon_header, text="Refresh", font=("Arial", 10), fg_color=self.color_main, width=70, height=25, command=self.refresh_all_bot_dropdowns).pack(side="right")
        
        self.status_container = ctk.CTkScrollableFrame(monitor_wrapper, height=200, fg_color="#0a0a0a", border_width=1, border_color=self.color_main)
        self.status_container.pack(fill="x", pady=2)
        for i in range(3): self.status_container.grid_columnconfigure(i, weight=1)
        
        ctk.CTkLabel(self.tab_dash, text="📋 Log Activity", font=("Arial", 14, "bold"), text_color="#a29bfe").pack(anchor="w", padx=15, pady=(10, 0))
        self.log_box = ctk.CTkTextbox(self.tab_dash, fg_color="black", font=("Consolas", 13), border_color=self.color_main, border_width=1)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    def setup_link_tab(self):
        wrapper = ctk.CTkFrame(self.tab_link, fg_color="transparent")
        wrapper.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(wrapper, text="Source Configuration", font=("Arial", 14, "bold"), text_color="#a29bfe").pack(side="left")
        ctk.CTkButton(wrapper, text="Refresh", fg_color="#34495e", font=("Arial", 10), width=70, height=25, command=self.refresh_link_list).pack(side="right", padx=5)

        input_f = ctk.CTkFrame(self.tab_link, fg_color="transparent")
        input_f.pack(fill="x", padx=20)
        self.link_name = ctk.CTkEntry(input_f, width=200, placeholder_text="Nama Bank/Sheet", border_color=self.color_main)
        self.link_name.pack(side="left", padx=5)
        self.link_url = ctk.CTkEntry(input_f, width=500, placeholder_text="URL Google Sheet", border_color=self.color_main)
        self.link_url.pack(side="left", padx=5)
        self.btn_save_link = ctk.CTkButton(input_f, text="SIMPAN LINK", fg_color=self.color_main, width=120, command=self.save_link_json, state="disabled")
        self.btn_save_link.pack(side="left", padx=5)
        self.link_name.bind("<KeyRelease>", lambda e: self.check_link_inputs())
        self.link_url.bind("<KeyRelease>", lambda e: self.check_link_inputs())
        
        self.link_list_frame = ctk.CTkScrollableFrame(self.tab_link, fg_color="#1a1a1a", border_width=1, border_color=self.color_main)
        self.link_list_frame.pack(fill="both", expand=True, padx=20, pady=10)

    def setup_config_tab(self):
        wrapper = ctk.CTkFrame(self.tab_config, fg_color="transparent")
        wrapper.pack(fill="x", padx=20, pady=10)
        header_area = ctk.CTkFrame(wrapper, fg_color="transparent")
        header_area.pack(fill="x")
        ctk.CTkLabel(header_area, text="System Parameters & Validation", font=("Arial", 14, "bold"), text_color="#a29bfe").pack(side="left")
        ctk.CTkButton(header_area, text="Refresh", fg_color="#34495e", font=("Arial", 10), width=70, height=25, command=self.refresh_config_list).pack(side="right")

        input_grid = ctk.CTkFrame(wrapper, fg_color="transparent")
        input_grid.pack(fill="x", pady=10)
        
        self.cfg_entries = {}
        labels = ["Name Col", "Nominal Col", "Username Col", "Status Col", "Max", "Timeout (m)", "DupTime (m)"]
        for i, label in enumerate(labels):
            r, c = (0, i) if i < 4 else (2, i-4)
            ctk.CTkLabel(input_grid, text=label+":", font=("Arial", 11, "bold")).grid(row=r, column=c, padx=10, pady=(5,0), sticky="w")
            en = ctk.CTkEntry(input_grid, width=150, border_color=self.color_main)
            en.grid(row=r+1, column=c, padx=10, pady=(0, 5), sticky="w")
            en.bind("<KeyRelease>", lambda e: self.check_cfg_inputs())
            self.cfg_entries[label] = en
        
        self.btn_save_cfg = ctk.CTkButton(input_grid, text="SIMPAN ATURAN", fg_color=self.color_main, width=150, height=35, font=("Arial", 11, "bold"), command=self.save_cfg_json, state="disabled")
        self.btn_save_cfg.grid(row=3, column=3, padx=10, pady=5, sticky="e")

        self.cfg_list_frame = ctk.CTkScrollableFrame(self.tab_config, fg_color="#1a1a1a", border_width=1, border_color=self.color_main)
        self.cfg_list_frame.pack(fill="both", expand=True, padx=20, pady=5)

    def setup_running_tab(self):
        top = ctk.CTkFrame(self.tab_run, fg_color="transparent"); top.pack(fill="x", padx=10, pady=5)
        ctk.CTkButton(top, text="🚀 TAMBAH BOT BARU", fg_color=self.color_main, height=30, command=self.add_bot_row).pack(side="left", padx=5)
        ctk.CTkButton(top, text="Refresh All", fg_color="#34495e", font=("Arial", 10), width=80, height=30, command=self.refresh_all_bot_dropdowns).pack(side="right", padx=10)
        
        self.h_row = ctk.CTkFrame(self.tab_run, fg_color=self.color_dark, height=40, border_color=self.color_main, border_width=1)
        self.h_row.pack(fill="x", padx=10, pady=(10, 0))
        
        headers = ["BOT NAME", "SHEET NAME", "ROW", "CONFIG", "LINK", "JSON API PATH", "ACTION"]
        for i, txt in enumerate(headers):
            self.h_row.grid_columnconfigure(i, weight=self.col_weights[i])
            lbl = ctk.CTkLabel(self.h_row, text=txt, font=("Arial", 10, "bold"), text_color="#a29bfe")
            lbl.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")
            
        self.run_container = ctk.CTkScrollableFrame(self.tab_run, fg_color="#1a1a1a", border_width=1, border_color=self.color_main)
        self.run_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def add_bot_row(self):
        rid = f"Bot_{int(time.time() * 1000)}"
        row_card = ctk.CTkFrame(self.run_container, fg_color="#2b2b2b", border_color=self.color_main, border_width=1); row_card.pack(fill="x", pady=2, padx=2)
        for i, w in enumerate(self.col_weights): row_card.grid_columnconfigure(i, weight=w)
        
        n_en = ctk.CTkEntry(row_card, height=28); n_en.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        s_en = ctk.CTkEntry(row_card, height=28); s_en.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        r_en = ctk.CTkEntry(row_card, height=28); r_en.insert(0, "2"); r_en.grid(row=0, column=2, padx=5, pady=10, sticky="ew")
        cfg_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("cfg_")]), height=28, command=lambda x, r=rid: self.lock_logic(r)); cfg_dd.set("Select"); cfg_dd.grid(row=0, column=3, padx=5, pady=10, sticky="ew")
        lnk_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("link_")]), height=28, command=lambda x, r=rid: self.lock_logic(r)); lnk_dd.set("Select"); lnk_dd.grid(row=0, column=4, padx=5, pady=10, sticky="ew")
        j_f = ctk.CTkFrame(row_card, fg_color="transparent"); j_f.grid(row=0, column=5, padx=5, pady=10, sticky="ew")
        j_en = ctk.CTkEntry(j_f, height=28); j_en.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(j_f, text="📂", width=30, height=28, command=lambda r=rid: self.browse_json_path(r)).pack(side="left", padx=(2,0))
        btn_c = ctk.CTkFrame(row_card, fg_color="transparent"); btn_c.grid(row=0, column=6, padx=5, pady=10, sticky="ew")
        b_web = ctk.CTkButton(btn_c, text="WEB", width=42, height=30, fg_color="#9b59b6", font=("Arial", 10, "bold"), state="disabled", command=lambda r=rid: self.bot_open_ui(r))
        b_web.pack(side="left", padx=1, expand=True, fill="x")
        b_start = ctk.CTkButton(btn_c, text="Start", width=42, height=30, fg_color="#2ecc71", font=("Arial", 10, "bold"), state="disabled", command=lambda r=rid: self.bot_start_ui(r))
        b_start.pack(side="left", padx=1, expand=True, fill="x")
        b_stop = ctk.CTkButton(btn_c, text="Stop", width=42, height=30, fg_color="#e67e22", font=("Arial", 10, "bold"), state="disabled", command=lambda r=rid: self.bot_stop_ui(r))
        b_stop.pack(side="left", padx=1, expand=True, fill="x")
        b_del = ctk.CTkButton(btn_c, text="Del", width=40, height=30, fg_color="#e74c3c", font=("Arial", 10, "bold"), command=lambda r=rid, f=row_card: self.bot_del(r, f))
        b_del.pack(side="left", padx=1, expand=True, fill="x")
        
        bot_count = len(self.bots); status_lbl = ctk.CTkLabel(self.status_container, text=f"⚪ {rid}: IDLE", font=("Arial", 13, "bold"), text_color="gray")
        status_lbl.grid(row=bot_count // 3, column=bot_count % 3, padx=15, pady=10, sticky="w")
        
        # Inisialisasi Bot dengan timeout_tracker dan last_processed
        self.bots[rid] = {
            'n_en': n_en, 's_en': s_en, 'r_en': r_en, 'cfg_dd': cfg_dd, 'lnk_dd': lnk_dd, 'j_en': j_en, 
            'b_web': b_web, 'b_start': b_start, 'b_stop': b_stop, 'b_del': b_del, 'status_lbl': status_lbl, 
            'driver': None, 'is_running': False, 'browser_ready': False,
            'timeout_tracker': {}, 'last_processed': {}
        }
        n_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); s_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); j_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r))

    def main_logic(self, rid):
        b = self.bots[rid]
        bot_name = b['n_en'].get()
        
        try:
            if b['driver']:
                base_url = self.global_domain.get().strip()
                if base_url:
                    base_url = base_url.rstrip('/')
                    target_url = f"{base_url}/_SubAg_Sub/DepositRequest.aspx?role=sa&userName=al3"
                    if not target_url.startswith("http"): target_url = "https://" + target_url
                    self.add_log(f"Membuka halaman mutasi...", bot_name, "blue")
                    b['driver'].get(target_url)
                    time.sleep(3)

            self.add_log("Koneksi Google Sheets...", bot_name, "blue")
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            json_path = b['j_en'].get()
            creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
            client = gspread.authorize(creds)
            
            with open(b['lnk_dd'].get(), "r") as f: lnk = json.load(f)
            sheet_file = client.open_by_url(lnk['url'])
            sheet = sheet_file.worksheet(b['s_en'].get())

            self.add_log("Bot Monitoring Aktif.", bot_name, "green")
            
            while b['is_running']:
                try:
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
                            
                            # Logika Max Nominal
                            if max_nom > 0 and int(nom_clean) > max_nom:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["❌"]]})
                                continue

                            dup_key = f"{nama_gs.lower()}_{nom_clean}"
                            
                            # --- LOGIKA DUPLIKAT (NYONTEK) ---
                            if dup_key in b['last_processed'] and (now - b['last_processed'][dup_key])/60 < dup_min:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["⚠️"]]})
                                continue 
                            
                            # --- LOGIKA TIMEOUT (NYONTEK) ---
                            r_key = f"row_{i}_{nama_gs}"
                            if r_key not in b['timeout_tracker']: 
                                b['timeout_tracker'][r_key] = now
                            
                            if (now - b['timeout_tracker'][r_key])/60 > to_min:
                                self.add_log(f"TIMEOUT: {nama_gs}", bot_name, "red")
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, idx_map['Status Col'] + 1), 'values': [["❌"]]})
                                continue

                            pending_queue.append({"row": i, "nama": nama_gs, "nominal": nom_clean, "dup_key": dup_key, "r_key": r_key})

                    if pending_queue:
                        self.add_log(f"SCAN: {len(pending_queue)} data ditemukan...", bot_name, "green")
                        try:
                            b['driver'].find_element(By.ID, "btnRefresh").click()
                            time.sleep(1.5)
                        except:
                            try: b['driver'].refresh(); time.sleep(3)
                            except: pass

                        for item in pending_queue:
                            if not b['is_running']: break
                            res_user = self.cari_dan_klik_web(b['driver'], item["nama"], item["nominal"])
                            if res_user:
                                self.add_log(f"SUCCESS: {item['nama']} selesai!", bot_name, "green")
                                updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], idx_map['Status Col'] + 1), 'values': [["✅"]]})
                                updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], idx_map['Username Col'] + 1), 'values': [[res_user]]})
                                
                                # Simpan waktu sukses untuk duplikat
                                b['last_processed'][item["dup_key"]] = time.time()
                                # Hapus dari tracker timeout karena sudah selesai
                                if item["r_key"] in b['timeout_tracker']: del b['timeout_tracker'][item["r_key"]]

                    if updates: 
                        try: sheet.batch_update(updates)
                        except: self.add_log("ERR: Gagal Update Sheet", bot_name, "red")
                    
                    time.sleep(5)
                except Exception as e:
                    self.add_log(f"LOOP ERR: {str(e)[:40]}", bot_name, "red")
                    time.sleep(5)
        except Exception as e:
            self.add_log(f"FATAL: {str(e)[:50]}", bot_name, "red")
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

    def refresh_config_list(self):
        for w in self.cfg_list_frame.winfo_children(): w.destroy()
        for fn in sorted(os.listdir()):
            if fn.startswith("cfg_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.cfg_list_frame, fg_color="#2b2b2b", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=4, padx=10)
                    detail_txt = (
                        f"📄 File: {fn}\n"
                        f"• Nama Col: {d.get('Name Col')} | Nom Col: {d.get('Nominal Col')} | "
                        f"User Col: {d.get('Username Col')} | Stat Col: {d.get('Status Col')}\n"
                        f"• Max: {d.get('Max')} | Timeout: {d.get('Timeout (m)')}m | Jeda: {d.get('DupTime (m)')}m"
                    )
                    ctk.CTkLabel(c, text=detail_txt, font=("Arial", 11), justify="left").pack(side="left", padx=15, pady=10)
                    ctk.CTkButton(c, text="🗑", width=40, height=40, fg_color="#e74c3c", command=lambda f=fn: [os.remove(f), self.refresh_config_list()]).pack(side="right", padx=15)
                except: continue

    def save_cfg_json(self):
        d = {k: v.get().strip() for k, v in self.cfg_entries.items()}
        nama_gabungan = f"{d['Name Col']}_{d['Nominal Col']}_{d['Username Col']}_{d['Status Col']}"
        safe_name = re.sub(r'[^\w]', '', nama_gabungan)
        filename = f"cfg_{safe_name}.json"
        with open(filename, "w") as j: json.dump(d, j, indent=4)
        for e in self.cfg_entries.values(): e.delete(0, 'end')
        self.btn_save_cfg.configure(state="disabled")
        self.refresh_config_list()
        self.add_log(f"Aturan disimpan: {filename}", "SYSTEM", "green")

    def save_link_json(self):
        name = self.link_name.get().strip()
        url = self.link_url.get().strip()
        if "docs.google.com/spreadsheets" not in url:
            messagebox.showwarning("URL Salah", "URL Google Sheet tidak valid!")
            return
        try:
            link_data = {"name": name, "url": url}
            filename = f"link_{name}.json"
            with open(filename, "w") as j: json.dump(link_data, j, indent=4)
            self.link_name.delete(0, 'end'); self.link_url.delete(0, 'end'); self.btn_save_link.configure(state="disabled")
            self.refresh_link_list()
        except Exception as e:
            messagebox.showerror("Error", f"Gagal: {str(e)}")

    def refresh_link_list(self):
        for w in self.link_list_frame.winfo_children(): w.destroy()
        for fn in sorted(os.listdir()):
            if fn.startswith("link_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.link_list_frame, fg_color="#2b2b2b", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=2, padx=10)
                    detail_txt = f"📄 Nama: {d.get('name')} | URL: {d.get('url')[:60]}..."
                    ctk.CTkLabel(c, text=detail_txt, font=("Arial", 11)).pack(side="left", padx=15, pady=8)
                    ctk.CTkButton(c, text="🗑", width=35, height=35, fg_color="#e74c3c", command=lambda f=fn: [os.remove(f), self.refresh_link_list()]).pack(side="right", padx=10)
                except: continue

    def lock_logic(self, rid):
        b = self.bots[rid]
        if b['browser_ready']: return 
        ready = all([b['n_en'].get().strip(), b['s_en'].get().strip(), b['j_en'].get().strip(), b['cfg_dd'].get() != "Select", b['lnk_dd'].get() != "Select"])
        b['b_web'].configure(state="normal" if ready else "disabled")

    def bot_open_ui(self, rid):
        b = self.bots[rid]; b['b_web'].configure(state="disabled")
        b['status_lbl'].configure(text=f"🔵 {b['n_en'].get()}: OPENING...", text_color="#3498db")
        threading.Thread(target=self.open_browser_task, args=(rid,), daemon=True).start()

    def open_browser_task(self, rid):
        b = self.bots[rid]
        try:
            url = self.global_domain.get().strip()
            if not url: messagebox.showerror("Error", "URL Kosong!"); b['b_web'].configure(state="normal"); return
            b['driver'] = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            b['driver'].get(url if url.startswith("http") else "https://"+url)
            b['browser_ready'] = True
            self.after(0, lambda: b['status_lbl'].configure(text=f"🔵 {b['n_en'].get()}: READY", text_color="#3498db"))
            self.after(0, lambda: b['b_start'].configure(state="normal"))
            while True:
                time.sleep(1)
                try: _ = b['driver'].window_handles
                except:
                    b['driver'] = None; b['is_running'] = False; b['browser_ready'] = False
                    self.after(0, lambda: b['status_lbl'].configure(text=f"⚪ {rid}: IDLE", text_color="gray"))
                    self.after(0, lambda: [b['b_web'].configure(state="normal"), b['b_start'].configure(state="disabled")])
                    break
        except: b['b_web'].configure(state="normal")

    def bot_start_ui(self, rid):
        b = self.bots[rid]
        b['is_running'] = True
        # Reset tracker setiap kali start
        b['timeout_tracker'] = {}
        b['last_processed'] = {}
        
        b['b_start'].configure(state="disabled")
        b['b_stop'].configure(state="normal")
        b['status_lbl'].configure(text=f"🟢 {b['n_en'].get()}: RUNNING", text_color="#2ecc71")
        threading.Thread(target=self.main_logic, args=(rid,), daemon=True).start()

    def bot_stop_ui(self, rid):
        b = self.bots[rid]; b['is_running'] = False; b['b_stop'].configure(state="disabled"); b['b_start'].configure(state="normal")
        b['status_lbl'].configure(text=f"🟠 {b['n_en'].get()}: STOPPED", text_color="#e67e22")

    def bot_del(self, rid, frame):
        if self.bots[rid]['driver']:
            try: self.bots[rid]['driver'].quit()
            except: pass
        self.bots[rid]['status_lbl'].destroy(); del self.bots[rid]; frame.destroy()

    def col_to_idx(self, letter):
        idx = 0
        for c in letter.upper().strip(): idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1

    def check_link_inputs(self): self.btn_save_link.configure(state="normal" if self.link_name.get().strip() and self.link_url.get().strip() else "disabled")
    def check_cfg_inputs(self): self.btn_save_cfg.configure(state="normal" if all(v.get().strip() for v in self.cfg_entries.values()) else "disabled")
    def browse_json_path(self, rid):
        fn = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if fn: self.bots[rid]['j_en'].delete(0, 'end'); self.bots[rid]['j_en'].insert(0, fn); self.lock_logic(rid)
    def refresh_all_bot_dropdowns(self):
        cfgs, lnks = ["Select"] + sorted([f for f in os.listdir() if f.startswith("cfg_")]), ["Select"] + sorted([f for f in os.listdir() if f.startswith("link_")])
        for b in self.bots.values(): b['cfg_dd'].configure(values=cfgs); b['lnk_dd'].configure(values=lnks)
    def on_closing(self):
        self.destroy()
        os._exit(0)

if __name__ == "__main__":
    app = PurpleBotApp()
    app.mainloop()
