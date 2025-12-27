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

# Konfigurasi Tema UI
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = "configuration.json"

class AppBotUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automation Bot - High Performance")
        self.geometry("950x850")
        self.entries = {}
        self.driver = None 
        self.is_running = False
        self.tracking_timeout = {} # Mencatat waktu pertama kali baris ditemukan di memori
        self.setup_ui()
        self.load_config()

    def setup_ui(self):
        self.config_frame = ctk.CTkFrame(self)
        self.config_frame.pack(fill="x", padx=10, pady=5)
        
        fields = ["Login URL:", "Deposit URL:", "Link Sheet:", "JSON Path:", "Sheet Name:"]
        for i, field in enumerate(fields):
            ctk.CTkLabel(self.config_frame, text=field).grid(row=i, column=0, padx=5, pady=2, sticky="e")
            entry = ctk.CTkEntry(self.config_frame, width=550)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky="w")
            self.entries[field] = entry
            if "JSON" in field:
                ctk.CTkButton(self.config_frame, text="Browse", width=80, command=self.browse_json).grid(row=i, column=2, padx=5)

        self.adv_frame = ctk.CTkFrame(self)
        self.adv_frame.pack(fill="x", padx=10, pady=5)
        
        settings = [
            ("Name Col", "A"), ("Nominal Col", "B"), ("Username Col", "C"), 
            ("Status Col", "D"), ("Start Row", "2"), ("Max Nominal", ""), ("Timeout (m)", "10")
        ]        
        for i, (label, val) in enumerate(settings):
            row_idx = i // 4
            col_idx = (i % 4) * 2
            ctk.CTkLabel(self.adv_frame, text=label).grid(
                row=row_idx, 
                column=col_idx, 
                padx=(10, 2), 
                pady=10, 
                sticky="w"
            )
            entry_adv = ctk.CTkEntry(self.adv_frame, width=65)
            entry_adv.insert(0, val)
            entry_adv.grid(
                row=row_idx, 
                column=col_idx + 1, 
                padx=(2, 10), 
                pady=10
            )
            self.entries[label] = entry_adv 

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=10, pady=10)
        
        self.btn_open = ctk.CTkButton(self.btn_frame, text="Open Browser", fg_color="#3498db", command=self.btn_open_browser)
        self.btn_open.pack(side="left", padx=5)
        self.btn_run = ctk.CTkButton(self.btn_frame, text="Start Bot", fg_color="#2ecc71", command=self.btn_start)
        self.btn_run.pack(side="left", padx=5)
        self.btn_stop_ui = ctk.CTkButton(self.btn_frame, text="Stop Bot", fg_color="#e74c3c", command=self.btn_stop)
        self.btn_stop_ui.pack(side="left", padx=5)
        ctk.CTkButton(self.btn_frame, text="Save Config", fg_color="#9b59b6", command=self.save_config).pack(side="left", padx=5)

        self.status_label = ctk.CTkLabel(self, text="Status: ● Idle", text_color="yellow", font=("Arial", 12, "bold"))
        self.status_label.pack(anchor="w", padx=15)

        self.log_box = ctk.CTkTextbox(self, height=400, fg_color="black", text_color="#00FF00", font=("Consolas", 12))
        self.log_box.pack(fill="both", padx=10, pady=5, expand=True)

    def add_log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n")
        self.log_box.see("end")

    def handle_alerts(self):
        """Menangani alert yang muncul secara tidak terduga"""
        try:
            WebDriverWait(self.driver, 1).until(EC.alert_is_present())
            alert = self.driver.switch_to.alert
            self.add_log(f"Alert Otomatis: {alert.text}")
            alert.accept()
        except: pass

    def btn_open_browser(self):
        def open_logic():
            options = webdriver.ChromeOptions()
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            self.driver.get(self.entries["Login URL:"].get())
            self.status_label.configure(text="Status: ● Browser Terbuka", text_color="#3498db")
        threading.Thread(target=open_logic, daemon=True).start()

    def cari_dan_klik_web(self, nama_gs, nominal_gs_string):
        """Mencari data di tabel web dan melakukan konfirmasi"""
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
                        
                        # Scroll agar tombol terlihat di layar
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_confirm)
                        time.sleep(0.5)
                        
                        # Klik tombol Konfirmasi
                        btn_confirm.click()
                        self.add_log(f"Mencoba Proses: {nama_gs}...")

                        # --- MENANGANI POP-UP OK/CANCEL ---
                        try:
                            # Tunggu alert muncul max 5 detik
                            WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                            alert = self.driver.switch_to.alert
                            alert_text = alert.text
                            alert.accept() # Klik OK
                            self.add_log(f"Berhasil Proses: {alert_text}")
                            
                            # Jeda agar sistem web memproses refresh status
                            time.sleep(1)
                            return username_web
                        except:
                            self.add_log("Gagal: Gagal Diproses.")
                            return None

                except: continue
            return None
        except: return None

    def main_loop(self):
        self.is_running = True
        self.btn_run.configure(state="disabled")
        self.status_label.configure(text="Status: ● Bot Berjalan", text_color="#2ecc71")
        
        try:
            self.driver.get(self.entries["Deposit URL:"].get())
            time.sleep(3)
        except: pass

        while self.is_running:
            try:
                # 1. SCRAPE SPREADSHEET
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_name(self.entries["JSON Path:"].get(), scope)
                client = gspread.authorize(creds)
                sheet = client.open_by_url(self.entries["Link Sheet:"].get()).worksheet(self.entries["Sheet Name:"].get())
                
                all_rows = sheet.get_all_values()
                start_row = int(self.entries["Start Row"].get())
                
                n_idx = self.col_to_idx(self.entries["Nominal Col"].get())
                u_idx = self.col_to_idx(self.entries["Username Col"].get())
                s_idx = self.col_to_idx(self.entries["Status Col"].get())
                name_idx = self.col_to_idx(self.entries["Name Col"].get())
                
                try: timeout_limit_min = int(self.entries["Timeout (m)"].get())
                except: timeout_limit_min = 10
                
                pending_queue = []
                updates = []

                # --- KRITERIA FILTER ---
                for i, row in enumerate(all_rows[start_row-1:], start=start_row):
                    nama = row[name_idx].strip()
                    nominal_raw = row[n_idx].strip()
                    username = row[u_idx].strip()
                    status = row[s_idx].strip()

                    if nama and nominal_raw and not username and not status:
                        row_key = f"row_{i}_{nama}"
                        
                        if row_key not in self.tracking_timeout:
                            self.tracking_timeout[row_key] = time.time()
                        
                        elapsed_min = (time.time() - self.tracking_timeout[row_key]) / 60
                        
                        if elapsed_min > timeout_limit_min:
                            self.add_log(f"TIMEOUT: {nama} (Baris {i}) ditandai ❌.")
                            updates.append({'range': gspread.utils.rowcol_to_a1(i, s_idx + 1), 'values': [["❌"]]})
                            if row_key in self.tracking_timeout: del self.tracking_timeout[row_key]
                            continue

                        nominal_clean = "".join(filter(str.isdigit, re.split(r'[.,]\d{2}$', nominal_raw)[0]))
                        pending_queue.append({
                            "row": i, "nama": nama, "nominal": nominal_clean,
                            "u_col": u_idx + 1, "s_col": s_idx + 1, "key": row_key
                        })

                # 2. PROSES KE WEB
                if pending_queue:
                    self.add_log(f"SCAN: {len(pending_queue)} data pending ditemukan.")
                    
                    try:
                        btn_refresh = self.driver.find_element(By.ID, "btnRefresh")
                        self.driver.execute_script("arguments[0].click();", btn_refresh)
                        time.sleep(1.5)
                    except:
                        self.driver.refresh()
                        time.sleep(3)
                    
                    self.handle_alerts()

                    for item in pending_queue:
                        if not self.is_running: break
                        
                        res_user = self.cari_dan_klik_web(item["nama"], item["nominal"])
                        if res_user:
                            self.add_log(f"COCOK! {item['nama']} Berhasil Diproses.")
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], item["s_col"]), 'values': [["✅"]]})
                            updates.append({'range': gspread.utils.rowcol_to_a1(item["row"], item["u_col"]), 'values': [[res_user]]})
                            if item["key"] in self.tracking_timeout: del self.tracking_timeout[item["key"]]

                # 3. BATCH UPDATE KE SHEETS
                if updates:
                    sheet.batch_update(updates)
                    self.add_log(f"Username & Status Berhasil diUpdate.")
                
                time.sleep(3)

            except Exception as e:
                self.add_log(f"Error Loop: {str(e)}")
                time.sleep(3)

        self.btn_run.configure(state="normal")

    def col_to_idx(self, letter):
        letter = letter.upper().strip()
        index = 0
        for char in letter: index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1

    def browse_json(self):
        filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filename:
            self.entries["JSON Path:"].delete(0, "end")
            self.entries["JSON Path:"].insert(0, filename)

    def save_config(self):
        config_data = {key: entry.get() for key, entry in self.entries.items()}
        with open(CONFIG_FILE, "w") as f:
            json.dump(config_data, f, indent=4)
        self.add_log("Konfigurasi disimpan.")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    config_data = json.load(f)
                for key, val in config_data.items():
                    if key in self.entries:
                        self.entries[key].delete(0, "end")
                        self.entries[key].insert(0, val)
            except: pass

    def btn_start(self):
        if not self.driver:
            self.add_log("Buka Browser dulu!")
            return
        self.save_config()
        if not self.is_running:
            self.tracking_timeout = {} 
            threading.Thread(target=self.main_loop, daemon=True).start()

    def btn_stop(self):
        self.is_running = False
        self.status_label.configure(text="Status: ● Bot Berhenti", text_color="red")

if __name__ == "__main__":
    app = AppBotUI()
    app.mainloop()

