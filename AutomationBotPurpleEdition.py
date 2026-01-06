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

# Konfigurasi Tema UI - Pakai "blue" agar tidak error, tapi tombol tetap ungu
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue") 

CONFIG_FILE = "configuration.json"

class AppBotUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Bot - Purple Edition")
        
        # UKURAN PAS & RAPI
        self.geometry("680x820") 
        self.resizable(False, False)
        
        self.entries = {}
        self.driver = None 
        self.is_running = False
        self.tracking_timeout = {} 
        self.last_processed = {}   
        
        self.setup_ui()
        self.load_config()
        self.update_button_states()

        # KONFIGURASI WARNA LOG
        self.log_box.tag_config("error", foreground="#FF4B4B")
        self.log_box.tag_config("success", foreground="#00D4FF")
        self.log_box.tag_config("warning", foreground="#FFA500")

    def setup_ui(self):
        self.config_frame = ctk.CTkFrame(self)
        self.config_frame.pack(fill="x", padx=10, pady=10)
        
        fields = ["Domain Website:", "Link Sheet:", "JSON Path:", "Sheet Name:"]
        for i, field in enumerate(fields):
            ctk.CTkLabel(self.config_frame, text=field).grid(row=i, column=0, padx=10, pady=3, sticky="e")
            if "JSON" in field:
                ctk.CTkButton(self.config_frame, text="Browse", width=60, height=24, command=self.browse_json, fg_color="#8e44ad").grid(row=i, column=1, padx=5)
                entry = ctk.CTkEntry(self.config_frame, width=380)
                entry.grid(row=i, column=2, padx=5, pady=3, sticky="w")
            else:
                entry = ctk.CTkEntry(self.config_frame, width=450)
                entry.grid(row=i, column=1, columnspan=2, padx=5, pady=3, sticky="w")
            entry.bind("<KeyRelease>", lambda e: self.update_button_states())
            self.entries[field] = entry

        self.adv_frame = ctk.CTkFrame(self)
        self.adv_frame.pack(fill="x", padx=10, pady=5)
        
        settings = [
            ("Name Col", "A"), ("Nominal Col", "B"), ("Username Col", "C"), 
            ("Status Col", "D"), ("Start Row", "2"), ("Max Nominal", "500000"), 
            ("Time Out(m)", "10"), ("Dup Time(m)", "2")
        ]        
        for i, (label, val) in enumerate(settings):
            row_idx = i // 4
            col_idx = (i % 4) * 2
            ctk.CTkLabel(self.adv_frame, text=label, font=("Arial", 11)).grid(row=row_idx, column=col_idx, padx=(5, 2), pady=5, sticky="w")
            entry_adv = ctk.CTkEntry(self.adv_frame, width=55)
            entry_adv.insert(0, val)
            entry_adv.grid(row=row_idx, column=col_idx + 1, padx=(2, 8), pady=5)
            entry_adv.bind("<KeyRelease>", lambda e: self.update_button_states())
            self.entries[label] = entry_adv 

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=10, pady=10)
        
        self.btn_open = ctk.CTkButton(self.btn_frame, text="Open Browser", width=150, fg_color="#9b59b6", hover_color="#8e44ad", command=self.btn_open_browser)
        self.btn_open.pack(side="left", padx=5)
        self.btn_run = ctk.CTkButton(self.btn_frame, text="Start Bot", width=150, fg_color="#6c5ce7", hover_color="#4834d4", command=self.btn_start)
        self.btn_run.pack(side="left", padx=5)
        self.btn_stop_ui = ctk.CTkButton(self.btn_frame, text="Stop", width=100, fg_color="#5f27cd", hover_color="#341f97", command=self.btn_stop)
        self.btn_stop_ui.pack(side="left", padx=5)
        self.btn_save = ctk.CTkButton(self.btn_frame, text="Save", width=100, fg_color="#a29bfe", hover_color="#6c5ce7", text_color="black", command=self.save_config)
        self.btn_save.pack(side="left", padx=5)

        self.status_label = ctk.CTkLabel(self, text="Status: ● Idle", text_color="yellow", font=("Arial", 12, "bold"))
        self.status_label.pack(anchor="w", padx=15)

        self.log_box = ctk.CTkTextbox(self, height=350, fg_color="black", text_color="#00FF00", font=("Consolas", 12))
        self.log_box.pack(fill="both", padx=10, pady=5, expand=True)

    def update_button_states(self):
        try:
            domain = self.entries["Domain Website:"].get().strip()
            sheet_link = self.entries["Link Sheet:"].get().strip()
            json_path = self.entries["JSON Path:"].get().strip()
            sheet_name = self.entries["Sheet Name:"].get().strip()
            fields_filled = all([domain, sheet_link, json_path, sheet_name])

            if not self.driver:
                self.btn_open.configure(state="normal")
                self.btn_run.configure(state="disabled")
                self.btn_stop_ui.configure(state="disabled")
            else:
                self.btn_open.configure(state="disabled")
                if self.is_running:
                    self.btn_run.configure(state="disabled")
                    self.btn_stop_ui.configure(state="normal")
                else:
                    self.btn_run.configure(state="normal" if fields_filled else "disabled")
                    self.btn_stop_ui.configure(state="disabled")
        except: pass

    def add_log(self, message, tag=None):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n", tag)
        self.log_box.see("end")

    def handle_alerts(self):
        try:
            WebDriverWait(self.driver, 1).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            self.add_log(f"Alert: {alert.text}", "warning")
            alert.accept()
        except: pass

    def btn_open_browser(self):
        def open_logic():
            try:
                domain = self.entries["Domain Website:"].get().strip()
                if not domain: return self.add_log("Error: Isi Domain!", "error")
                if not domain.startswith("http"): domain = "https://" + domain
                options = webdriver.ChromeOptions()
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                self.driver.get(domain)
                self.after(0, lambda: self.status_label.configure(text="Status: ● Browser Terbuka", text_color="#3498db"))
                self.add_log(f"Browser Terbuka. Silakan login.")
                self.after(0, self.update_button_states)
                def monitor():
                    while True:
                        time.sleep(1)
                        try: _ = self.driver.window_handles
                        except:
                            self.driver = None
                            self.is_running = False
                            self.after(0, self.update_button_states)
                            self.after(0, lambda: self.status_label.configure(text="Status: ● Idle", text_color="yellow"))
                            break
                threading.Thread(target=monitor, daemon=True).start()
            except Exception as e: self.add_log(f"Error: {str(e)}", "error")
        threading.Thread(target=open_logic, daemon=True).start()

    def cari_dan_klik_web(self, nama_gs, nominal_gs_string):
        self.handle_alerts()
        try:
            nama_gs_bersih = re.sub(r'[^a-z0-9]', '', nama_gs.lower())
            rows = self.driver.find_elements(By.CSS_SELECTOR, "table#report tbody tr[class^='Grid']")
            for row in rows:
                try:
                    name_web = row.find_element(By.CLASS_NAME, "fromAccountName").text.strip()
                    amount_raw = row.find_element(By.CLASS_NAME, "amount").text.strip()
                    amount_web_clean = "".join(filter(str.isdigit, amount_raw.split('.')[0]))
                    username_web = row.find_element(By.CLASS_NAME, "username").text.strip()
                    name_web_bersih = re.sub(r'[^a-z0-9]', '', name_web.lower())

                    if nama_gs_bersih == name_web_bersih and nominal_gs_string == amount_web_clean:
                        btn_confirm = row.find_element(By.CLASS_NAME, "confirm").find_element(By.TAG_NAME, "input")
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_confirm)
                        time.sleep(0.5)
                        btn_confirm.click()
                        try:
                            WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                            alert = self.driver.switch_to.alert
                            alert_text = alert.text
                            alert.accept()
                            self.add_log(f"Berhasil Proses Web: {alert_text}", "success")
                            return username_web
                        except: return None
                except: continue
        except: pass
        return None

    def main_loop(self):
        self.is_running = True
        self.after(0, self.update_button_states)
        self.after(0, lambda: self.status_label.configure(text="Status: ● Bot Berjalan", text_color="#2ecc71"))
        
        try:
            domain = self.entries["Domain Website:"].get().strip()
            if not domain.startswith("http"): domain = "https://" + domain
            deposit_url = domain.rstrip('/') + "/_SubAg_Sub/DepositRequest.aspx?"
            self.driver.get(deposit_url)
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
                
                timeout_limit_min = int(self.entries["Time Out(m)"].get()) if self.entries["Time Out(m)"].get() else 10
                dup_limit_min = int(self.entries["Dup Time(m)"].get()) if self.entries["Dup Time(m)"].get() else 2
                max_nom_limit = int(self.entries["Max Nominal"].get()) if self.entries["Max Nominal"].get() else None

                pending_queue, updates = [], []

                for i, row in enumerate(all_rows[start_row-1:], start=start_row):
                    if not self.is_running: break
                    try:
                        nama, nominal_raw = row[name_idx].strip(), row[n_idx].strip()
                        username, status = row[u_idx].strip(), row[s_idx].strip()
                    except: continue

                    if nama and nominal_raw and not username and not status:
                        now = time.time()
                        nominal_clean = "".join(filter(str.isdigit, re.split(r'[.,]\d{2}$', nominal_raw)[0]))
                        
                        if max_nom_limit and int(nominal_clean) > max_nom_limit:
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["❌"]]})
                            continue

                        dup_key = f"{nama.lower()}_{nominal_clean}"
                        if dup_key in self.last_processed and (now - self.last_processed[dup_key])/60 < dup_limit_min:
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["⚠️"]]})
                            continue

                        row_key = f"row_{i}_{nama}"
                        if row_key not in self.tracking_timeout: self.tracking_timeout[row_key] = now
                        if (now - self.tracking_timeout[row_key])/60 > timeout_limit_min:
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["❌"]]})
                            continue

                        pending_queue.append({"row": i, "nama": nama, "nominal": nominal_clean, "u_col": u_idx + 1, "s_col": s_idx + 1, "key": row_key, "dup_key": dup_key})

                if pending_queue:
                    self.add_log(f"Scan {len(pending_queue)} data pending...")
                    try: self.driver.find_element(By.ID, "btnRefresh").click(); time.sleep(1.5)
                    except: self.driver.refresh(); time.sleep(3)
                    
                    self.handle_alerts()
                    for item in pending_queue:
                        if not self.is_running: break
                        res_user = self.cari_dan_klik_web(item["nama"], item["nominal"])
                        if res_user:
                            self.add_log(f"SUKSES! {item['nama']}", "success")
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], item["s_col"]), 'values': [["✅"]]})
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], item["u_col"]), 'values': [[res_user]]})
                            self.last_processed[item["dup_key"]] = time.time()
                            if item["key"] in self.tracking_timeout: del self.tracking_timeout[item["key"]]

                if updates: sheet.batch_update(updates)
                time.sleep(3)
            except Exception as e:
                self.add_log(f"Error: {str(e)}", "error")
                time.sleep(5)
        self.after(0, self.update_button_states)

    def btn_start(self):
        if not self.driver: return
        self.save_config()
        self.tracking_timeout, self.last_processed = {}, {}
        threading.Thread(target=self.main_loop, daemon=True).start()

    def btn_stop(self):
        self.is_running = False
        self.add_log("Bot Berhenti.", "error")
        self.after(0, self.update_button_states)

    def col_to_idx(self, letter):
        idx = 0
        for c in letter.upper().strip(): idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1

    def browse_json(self):
        f = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if f:
            self.entries["JSON Path:"].delete(0, "end"); self.entries["JSON Path:"].insert(0, f)
            self.update_button_states()

    def save_config(self):
        d = {k: v.get() for k, v in self.entries.items()}
        with open(CONFIG_FILE, "w") as f: json.dump(d, f, indent=4)
        self.add_log("Simpan Berhasil.")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f: d = json.load(f)
                for k, v in d.items():
                    if k in self.entries: self.entries[k].delete(0, "end"); self.entries[k].insert(0, v)
            except: pass

if __name__ == "__main__":
    app = AppBotUI()
    app.mainloop()
