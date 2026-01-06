import customtkinter as ctk
from tkinter import filedialog
import datetime
import threading
import time
import gspread
import re
import json
import os
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set Tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue") 

CONFIG_FILE = "configuration.json"

class AppBotUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Bot - Purple Edition")
        
        # Lebar 580px adalah angka paling pas untuk label "Username" & "Status"
        self.geometry("580x850") 
        self.resizable(False, False)
        
        self.entries = {}
        self.driver = None 
        self.is_running = False
        self.tracking_timeout = {} 
        self.last_processed = {}   
        
        self.setup_ui()
        self.load_config()
        self.update_button_states()

        # Konfigurasi Tag Warna Log
        self.log_box.tag_config("error", foreground="#FF4B4B")
        self.log_box.tag_config("success", foreground="#00D4FF")
        self.log_box.tag_config("warning", foreground="#FFA500")

    def setup_ui(self):
        # Frame Konfigurasi Utama
        self.config_frame = ctk.CTkFrame(self)
        self.config_frame.pack(fill="x", padx=10, pady=10)
        
        fields = ["Domain Website:", "Link Sheet:", "JSON Path:", "Sheet Name:"]
        for i, field in enumerate(fields):
            ctk.CTkLabel(self.config_frame, text=field).grid(row=i, column=0, padx=(10, 5), pady=3, sticky="e")
            if "JSON" in field:
                ctk.CTkButton(self.config_frame, text="Browse", width=60, height=24, command=self.browse_json, fg_color="#8e44ad").grid(row=i, column=1, padx=2)
                entry = ctk.CTkEntry(self.config_frame, width=330) 
                entry.grid(row=i, column=2, padx=(2, 10), pady=3, sticky="w")
            else:
                entry = ctk.CTkEntry(self.config_frame, width=400) 
                entry.grid(row=i, column=1, columnspan=2, padx=(2, 10), pady=3, sticky="w")
            
            entry.bind("<KeyRelease>", lambda e: self.update_button_states())
            self.entries[field] = entry

        # Frame Advanced Settings (Label Sesuai Request)
        self.adv_frame = ctk.CTkFrame(self)
        self.adv_frame.pack(fill="x", padx=10, pady=5)
        
        settings = [
            ("Name", "A"), ("Nominal", "B"), ("Username", "C"), ("Status", "D"),
            ("Start Row", "2"), ("Max", "500000"), ("Time Out(m)", "10"), ("Dup Time(m)", "2")
        ]        
        for i, (label, val) in enumerate(settings):
            row_idx = i // 4
            col_idx = (i % 4) * 2
            ctk.CTkLabel(self.adv_frame, text=label, font=("Arial", 10, "bold")).grid(row=row_idx, column=col_idx, padx=(8, 2), pady=10, sticky="w")
            
            # Textbox disamakan ukurannya (65px) agar grid terlihat simetris
            entry_adv = ctk.CTkEntry(self.adv_frame, width=65)
            entry_adv.insert(0, val)
            entry_adv.grid(row=row_idx, column=col_idx + 1, padx=(2, 8), pady=10)
            entry_adv.bind("<KeyRelease>", lambda e: self.update_button_states())
            
            # Mapping internal agar logika script tidak berubah
            key_map = {
                "Name": "Name Col", "Nominal": "Nominal Col", "Username": "Username Col", 
                "Status": "Status Col", "Max": "Max Nominal"
            }
            self.entries[key_map.get(label, label)] = entry_adv 

        # Frame Tombol Kontrol (Simetris 130 px)
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=10, pady=15)
        
        btn_configs = [
            ("Open Browser", "#9b59b6", self.btn_open_browser),
            ("Start Bot", "#6c5ce7", self.btn_start),
            ("Stop", "#5f27cd", self.btn_stop),
            ("Save", "#a29bfe", self.save_config)
        ]

        for i, (txt, clr, cmd) in enumerate(btn_configs):
            btn = ctk.CTkButton(self.btn_frame, text=txt, width=130, fg_color=clr, command=cmd)
            btn.grid(row=0, column=i, padx=4)
            if txt == "Open Browser": self.btn_open = btn
            elif txt == "Start Bot": self.btn_run = btn
            elif txt == "Stop": self.btn_stop_ui = btn

        self.status_label = ctk.CTkLabel(self, text="Status: ● Idle", text_color="yellow", font=("Arial", 11, "bold"))
        self.status_label.pack(anchor="w", padx=15)

        self.log_box = ctk.CTkTextbox(self, height=350, fg_color="black", text_color="#00FF00", font=("Consolas", 11))
        self.log_box.pack(fill="both", padx=10, pady=5, expand=True)

    # --- LOGIKA BOT (100% SAMA) ---

    def update_button_states(self):
        domain = self.entries["Domain Website:"].get().strip()
        sheet_link = self.entries["Link Sheet:"].get().strip()
        json_path = self.entries["JSON Path:"].get().strip()
        if not self.driver:
            self.btn_open.configure(state="normal" if domain else "disabled")
            self.btn_run.configure(state="disabled")
            self.btn_stop_ui.configure(state="disabled")
        else:
            self.btn_open.configure(state="disabled")
            if self.is_running:
                self.btn_run.configure(state="disabled")
                self.btn_stop_ui.configure(state="normal")
            else:
                is_ready = all([domain, sheet_link, json_path])
                self.btn_run.configure(state="normal" if is_ready else "disabled")
                self.btn_stop_ui.configure(state="disabled")

    def btn_open_browser(self):
        def open_logic():
            try:
                domain = self.entries["Domain Website:"].get().strip()
                if not domain.startswith("http"): domain = "https://" + domain
                options = webdriver.ChromeOptions()
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                self.driver.get(domain)
                self.after(0, lambda: self.status_label.configure(text="Status: ● Browser Terbuka", text_color="#3498db"))
                self.add_log("Browser Berhasil Terbuka.")
                self.after(0, self.update_button_states)
                def monitor():
                    while True:
                        time.sleep(1)
                        try: _ = self.driver.window_handles
                        except:
                            self.driver = None; self.is_running = False
                            self.after(0, lambda: self.status_label.configure(text="Status: ● Idle", text_color="yellow"))
                            self.after(0, self.update_button_states)
                            break
                threading.Thread(target=monitor, daemon=True).start()
            except Exception as e: self.add_log(f"Error: {str(e)}", "error")
        threading.Thread(target=open_logic, daemon=True).start()

    def main_loop(self):
        self.is_running = True
        self.after(0, self.update_button_states)
        self.after(0, lambda: self.status_label.configure(text="Status: ● Bot Berjalan", text_color="#2ecc71"))
        try:
            domain = self.entries["Domain Website:"].get().strip()
            if not domain.startswith("http"): domain = "https://" + domain
            self.driver.get(domain.rstrip('/') + "/_SubAg_Sub/DepositRequest.aspx?")
            time.sleep(3)
        except: pass
        while self.is_running:
            try:
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_name(self.entries["JSON Path:"].get(), scope)
                client = gspread.authorize(creds)
                sheet = client.open_by_url(self.entries["Link Sheet:"].get()).worksheet(self.entries["Sheet Name:"].get())
                all_rows = sheet.get_all_values()
                start_row = int(self.entries["Start Row"].get())
                n_idx, u_idx, s_idx, name_idx = self.col_to_idx(self.entries["Nominal Col"].get()), self.col_to_idx(self.entries["Username Col"].get()), self.col_to_idx(self.entries["Status Col"].get()), self.col_to_idx(self.entries["Name Col"].get())
                pending_queue, updates = [], []
                for i, row in enumerate(all_rows[start_row-1:], start=start_row):
                    if not self.is_running: break
                    try:
                        nama, nom_raw = row[name_idx].strip(), row[n_idx].strip()
                        u_gs, stat_gs = row[u_idx].strip(), row[s_idx].strip()
                        if nama and nom_raw and not u_gs and not stat_gs:
                            now = time.time()
                            nom_clean = "".join(filter(str.isdigit, re.split(r'[.,]\d{2}$', nom_raw)[0]))
                            max_nom = int(self.entries["Max Nominal"].get()) if self.entries["Max Nominal"].get() else None
                            if max_nom and int(nom_clean) > max_nom:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["❌"]]}); continue
                            dup_min = int(self.entries["Dup Time(m)"].get()) if self.entries["Dup Time(m)"].get() else 2
                            dup_key = f"{nama.lower()}_{nom_clean}"
                            if dup_key in self.last_processed and (now - self.last_processed[dup_key])/60 < dup_min:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["⚠️"]]}); continue
                            timeout_min = int(self.entries["Time Out(m)"].get()) if self.entries["Time Out(m)"].get() else 10
                            r_key = f"row_{i}_{nama}"
                            if r_key not in self.tracking_timeout: self.tracking_timeout[r_key] = now
                            if (now - self.tracking_timeout[r_key])/60 > timeout_min:
                                updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["❌"]]}); continue
                            pending_queue.append({"row": i, "nama": nama, "nominal": nom_clean, "u_col": u_idx+1, "s_col": s_idx+1, "key": r_key, "dup_key": dup_key})
                    except: continue
                if pending_queue:
                    self.add_log(f"SCAN: {len(pending_queue)} data pending ditemukan...", "warning")
                    try: self.driver.find_element(By.ID, "btnRefresh").click(); time.sleep(1.5)
                    except: self.driver.refresh(); time.sleep(3)
                    self.handle_alerts()
                    for item in pending_queue:
                        if not self.is_running: break
                        res_user = self.cari_dan_klik_web(item["nama"], item["nominal"])
                        if res_user:
                            self.add_log(f"SUKSES! {item['nama']} diproses.", "success")
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], item["s_col"]), 'values': [["✅"]]})
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], item["u_col"]), 'values': [[res_user]]})
                            self.last_processed[item["dup_key"]] = time.time()
                if updates: sheet.batch_update(updates)
                time.sleep(3)
            except Exception as e: self.add_log(f"Error Loop: {str(e)}", "error"); time.sleep(5)
        self.after(0, self.update_button_states)

    def cari_dan_klik_web(self, nama_gs, nominal_gs_string):
        try:
            nama_gs_bersih = re.sub(r'[^a-z0-9]', '', nama_gs.lower())
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table#report tbody tr[class^='Grid']")
            for row in rows:
                try:
                    name_web = row.find_element(By.CLASS_NAME, "fromAccountName").text.strip()
                    amount_raw = row.find_element(By.CLASS_NAME, "amount").text.strip()
                    amount_web_clean = "".join(filter(str.isdigit, amount_raw.split('.')[0]))
                    name_web_bersih = re.sub(r'[^a-z0-9]', '', name_web.lower())
                    username_web = row.find_element(By.CLASS_NAME, "username").text.strip()
                    if nama_gs_bersih == name_web_bersih and nominal_gs_string == amount_web_clean:
                        btn_confirm = row.find_element(By.CLASS_NAME, "confirm").find_element(By.TAG_NAME, "input")
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_confirm)
                        time.sleep(0.5); btn_confirm.click()
                        WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                        alert = self.driver.switch_to.alert; alert.accept()
                        return username_web
                except: continue
        except: pass
        return None

    def handle_alerts(self):
        try:
            WebDriverWait(self.driver, 1).until(EC.alert_is_present())
            self.driver.switch_to.alert.accept()
        except: pass

    def browse_json(self):
        f = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if f:
            self.entries["JSON Path:"].delete(0, "end"); self.entries["JSON Path:"].insert(0, f)
            self.update_button_states()

    def add_log(self, message, tag=None):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n", tag)
        self.log_box.see("end")

    def save_config(self):
        d = {k: v.get() for k, v in self.entries.items()}
        with open(CONFIG_FILE, "w") as f: json.dump(d, f, indent=4)
        self.add_log("Konfigurasi disimpan.")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f: d = json.load(f)
                for k, v in d.items():
                    if k in self.entries: self.entries[k].delete(0, "end"); self.entries[k].insert(0, v)
            except: pass

    def col_to_idx(self, letter):
        idx = 0
        for c in letter.upper().strip(): idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1

    def btn_start(self):
        if not self.driver: return
        self.save_config(); self.tracking_timeout, self.last_processed = {}, {}
        threading.Thread(target=self.main_loop, daemon=True).start()

    def btn_stop(self):
        self.is_running = False
        self.add_log("Bot Berhenti.", "error")

if __name__ == "__main__":
    app = AppBotUI()
    app.mainloop()
