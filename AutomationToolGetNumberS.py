import customtkinter as ctk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import threading
import os
import time

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class BotDataApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automation Tool - Get Number SGA")
        self.geometry("600x620")
        
        self.driver = None 
        self.file_path = None 

        # --- UI Layout ---
        self.label_url = ctk.CTkLabel(self, text="Base URL :", font=("Arial", 12, "bold"))
        self.label_url.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.entry_url = ctk.CTkEntry(self, placeholder_text="Contoh: https://domain.com", width=350)
        self.entry_url.grid(row=0, column=1, padx=20, pady=(20, 10), sticky="ew")
        self.entry_url.bind("<KeyRelease>", self.check_url_input)

        self.btn_browser = ctk.CTkButton(self, text="1. Buka Browser", fg_color="#8e44ad", 
                                        state="disabled", command=self.run_open_browser_thread)
        self.btn_browser.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

        self.btn_upload = ctk.CTkButton(self, text="2. Upload File .xlsx", fg_color="#2c3e50", 
                                       state="disabled", command=self.upload_file)
        self.btn_upload.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

        self.label_progress = ctk.CTkLabel(self, text="Progress: 0%", font=("Arial", 11))
        self.label_progress.grid(row=3, column=0, columnspan=2, padx=20, pady=(10, 0), sticky="w")

        self.progress_bar = ctk.CTkProgressBar(self, height=12)
        self.progress_bar.grid(row=4, column=0, columnspan=2, padx=20, pady=(5, 10), sticky="ew")
        self.progress_bar.set(0)

        self.btn_start = ctk.CTkButton(self, text="3. Start", fg_color="#27ae60", height=45, 
                                      font=("Arial", 14, "bold"), state="disabled", command=self.run_start_thread)
        self.btn_start.grid(row=5, column=0, columnspan=2, padx=20, pady=10, sticky="s")

        self.info_box = ctk.CTkLabel(self, text="Masukkan Base URL untuk memulai....", fg_color="#34495e", 
                                   text_color="white", corner_radius=8, height=100, width=560, wraplength=500)
        self.info_box.grid(row=7, column=0, columnspan=2, padx=20, pady=(5, 20), sticky="ew")

    def check_url_input(self, event=None):
        if len(self.entry_url.get().strip()) > 10:
            self.btn_browser.configure(state="normal")
        else:
            self.btn_browser.configure(state="disabled")

    # Fungsi Update Info dengan dukungan WARNA
    def update_info(self, text, color="white"):
        self.info_box.configure(text=text, text_color=color)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path = path
            self.update_info(f"File Terpilih: {os.path.basename(path)}. Klik START.", "#2ecc71")
            self.btn_upload.configure(fg_color="#16a085")
            self.btn_start.configure(state="normal")

    def run_open_browser_thread(self):
        self.btn_browser.configure(state="disabled")
        threading.Thread(target=self.open_browser_logic, daemon=True).start()

    def open_browser_logic(self):
        url = self.entry_url.get().strip()
        try:
            self.update_info("Membuka browser... mohon tunggu.")
            self.driver = webdriver.Chrome()
            self.driver.get(url)
            self.update_info("Browser Terbuka. Silakan login, lalu Upload File Excel.")
            self.btn_upload.configure(state="normal")
        except Exception as e:
            self.update_info(f"Gagal buka browser: {str(e)}", "red")
            self.btn_browser.configure(state="normal")

    def run_start_thread(self):
        self.btn_start.configure(state="disabled")
        threading.Thread(target=self.start_logic, daemon=True).start()

    def start_logic(self):
        try:
            base_url = self.entry_url.get().strip().rstrip('/')
            self.driver.get(f"{base_url}/SPanel/Player/MemberStat")
            wait = WebDriverWait(self.driver, 10)

            # Dropdown 500 (Opsional)
            try:
                drop_xpath = "/html/body/div[1]/div/div/div/div/div[2]/div/div[1]/div[1]/select"
                drop_el = wait.until(EC.element_to_be_clickable((By.XPATH, drop_xpath)))
                drop_el.click()
                time.sleep(1)
                opt_500 = self.driver.find_element(By.XPATH, f"{drop_xpath}/option[@value='500'] | {drop_xpath}/option[text()='500']")
                opt_500.click()
                time.sleep(3)
            except: pass

            df = pd.read_excel(self.file_path, header=None)
            if df.shape[1] < 2: df[1] = ""
            total_rows = len(df)

            for index, row in df.iterrows():
                username_target = str(df.iloc[index, 0]).strip()
                if username_target == "nan" or not username_target: continue

                # Update UI Progress
                calc_prog = (index + 1) / total_rows
                pct = int(calc_prog * 100)
                self.label_progress.configure(text=f"Progress: {pct}% ({index+1}/{total_rows})")
                self.progress_bar.set(calc_prog)
                self.update_info(f"Mencari: {username_target}")

                # Proses Input Search
                in_xpath = "/html/body/div[1]/div/div/div/div/div[1]/div/input"
                btn_xpath = "/html/body/div[1]/div/div/div/div/div[1]/div/button[3]"
                
                input_f = wait.until(EC.presence_of_element_located((By.XPATH, in_xpath)))
                input_f.clear()
                input_f.send_keys(username_target)
                self.driver.find_element(By.XPATH, btn_xpath).click()
                
                time.sleep(2) # Tunggu hasil search muncul

                phone_found = "Tidak Ditemukan"
                match_found = False
                
                # --- LOGIKA ANTI MACET (Cek Next Halaman) ---
                for _ in range(10): 
                    try:
                        time.sleep(1)
                        name_xpath = "/html/body/div[1]/div/div/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[2]/div[1]/span[1]"
                        name_elements = self.driver.find_elements(By.XPATH, name_xpath)

                        if name_elements:
                            name_web = name_elements[0].text.strip()
                            if name_web.lower() == username_target.lower():
                                phone_xpath = "/html/body/div[1]/div/div/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[2]/div[2]/span"
                                phone_found = self.driver.find_element(By.XPATH, phone_xpath).text.strip()
                                match_found = True
                                break
                        
                        # Jika tidak ketemu di baris ini, cek tombol Next
                        nx_xpath = "/html/body/div[1]/div/div/div/div/div[2]/div/div[1]/div[2]/button[3]"
                        btn_nx = self.driver.find_element(By.XPATH, nx_xpath)
                        if btn_nx.is_enabled() and "disabled" not in btn_nx.get_attribute("class"):
                            btn_nx.click()
                        else:
                            break # Halaman habis
                    except:
                        break

                df.iloc[index, 1] = phone_found
                self.update_info(f"ID: {username_target} -> {phone_found}")

            # --- SELESAI & SIMPAN ---
            try:
                df.to_excel(self.file_path, index=False, header=False)
                self.label_progress.configure(text="Progress: 100% (Selesai)")
                self.update_info("PROSES SELESAI! Data berhasil disimpan ke Excel.", "#2ecc71") # HIJAU
            except Exception as e:
                self.update_info(f"Error Simpan: Tutup file Excel Anda! ({str(e)})", "red") # MERAH

        except Exception as e:
            self.update_info(f"Error Sistem: {str(e)}", "red") # MERAH
        finally:
            self.btn_start.configure(state="normal")

if __name__ == "__main__":
    app = BotDataApp()
    app.mainloop()
