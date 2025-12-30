import customtkinter as ctk
from tkinter import filedialog
import pandas as pd
import threading
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Setup Tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class BotDataApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automation Tool - Get Number")
        self.geometry("600x550")
        self.file_path = None
        self.driver = None

        # --- UI Layout ---
        self.label_url = ctk.CTkLabel(self, text="Base URL :", font=("Arial", 12, "bold"))
        self.label_url.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.entry_url = ctk.CTkEntry(self, placeholder_text="Contoh: https://domain.com", width=350)
        self.entry_url.grid(row=0, column=1, padx=20, pady=(20, 10), sticky="ew")
        self.entry_url.bind("<KeyRelease>", self.check_url_input)

        self.btn_browser = ctk.CTkButton(self, text="Buka Browser", command=self.open_browser, 
                                        fg_color="#8e44ad", height=35, state="disabled")
        self.btn_browser.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

        self.btn_upload = ctk.CTkButton(self, text="Upload File .xlsx", command=self.upload_file, 
                                        fg_color="#2c3e50", height=35, state="disabled")
        self.btn_upload.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

        self.label_progress = ctk.CTkLabel(self, text="Progress: 0%", font=("Arial", 12))
        self.label_progress.grid(row=3, column=0, padx=20, pady=(20, 5), sticky="w")

        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=4, column=0, columnspan=2, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)

        self.btn_start = ctk.CTkButton(self, text="Start", command=self.run_bot_thread, 
                                        fg_color="#27ae60", height=45, font=("Arial", 14, "bold"), state="disabled")
        self.btn_start.grid(row=5, column=0, columnspan=2, padx=20, pady=20, sticky="s")

        # --- PANEL KETERANGAN ---
        self.label_info_title = ctk.CTkLabel(self, text="Keterangan Sistem:", font=("Arial", 11, "bold"))
        self.label_info_title.grid(row=6, column=0, padx=20, sticky="w")

        self.info_box = ctk.CTkLabel(self, text="Masukkan Base URL untuk memulai...", fg_color="#34495e", 
                                     text_color="white", corner_radius=8, height=80, width=560, wraplength=500)
        self.info_box.grid(row=7, column=0, columnspan=2, padx=20, pady=(5, 20), sticky="ew")

    def update_info(self, text, color="white"):
        self.info_box.configure(text=text, text_color=color)

    def check_url_input(self, event=None):
        if not self.driver:
            url = self.entry_url.get().strip()
            if len(url) > 5:
                self.btn_browser.configure(state="normal")
            else:
                self.btn_browser.configure(state="disabled")

    def open_browser(self):
        base_url = self.entry_url.get().strip()
        self.btn_browser.configure(state="disabled")
        try:
            self.update_info("Membuka browser... mohon tunggu.")
            self.driver = webdriver.Chrome()
            self.driver.get(base_url)
            self.update_info("Browser Terbuka. Silakan Upload File Excel.")
            self.btn_upload.configure(state="normal")
        except Exception as e:
            self.update_info(f"Error Browser: Pastikan Chrome sudah terinstall. ({str(e)})", "red")
            self.btn_browser.configure(state="normal")

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            fname = self.file_path.split('/')[-1]
            self.update_info(f"File Terpilih: {fname}. Klik START.", "#2ecc71")
            self.btn_start.configure(state="normal")
        else:
            self.update_info("File batal diupload.", "yellow")

    def run_bot_thread(self):
        self.btn_start.configure(state="disabled")
        threading.Thread(target=self.start_processing, daemon=True).start()

    def start_processing(self):
        try:
            base_url = self.entry_url.get().strip().rstrip('/')
            
            # Error Handling jika file tidak bisa dibaca
            try:
                df = pd.read_excel(self.file_path, header=None)
            except Exception as e:
                self.update_info(f"Error: File Excel tidak bisa dibaca. ({str(e)})", "red")
                self.btn_start.configure(state="normal")
                return

            # Mengatasi error 'iloc cannot enlarge' secara otomatis
            if df.shape[1] < 2:
                df[1] = ""

            total_rows = len(df)
            target_xpath = "/html/body/form/div[3]/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/input"

            for index, row in df.iterrows():
                name = str(df.iloc[index, 0]) 
                full_link = f"{base_url}/_SubAg_Sub/PopUpAccInfo.aspx?userName={name}&type=member&role=sa"
                
                try:
                    self.driver.get(full_link)
                    element = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, target_xpath))
                    )
                    data_diambil = element.get_attribute('value')
                except:
                    data_diambil = "Gagal Ambil Data"

                df.iloc[index, 1] = data_diambil
                
                prog = (index + 1) / total_rows
                self.progress_bar.set(prog)
                self.label_progress.configure(text=f"Progress: {int(prog * 100)}% ({index+1}/{total_rows})")
                self.update_info(f"ID: {name} -> {data_diambil}")
                
                time.sleep(0.3)

            # Error Handling saat simpan (Biasanya karena file sedang dibuka di Excel)
            try:
                df.to_excel(self.file_path, index=False, header=False)
                self.update_info("SELESAI! Data berhasil disimpan ke Excel.", "#2ecc71")
            except Exception as e:
                self.update_info(f"Error Simpan: Tutup file Excel Anda dulu! ({str(e)})", "red")
            
        except Exception as e:
            self.update_info(f"Error Sistem: {str(e)}", "red")
        finally:
            self.btn_start.configure(state="normal")

if __name__ == "__main__":
    app = BotDataApp()
    app.mainloop()


