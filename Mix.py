import customtkinter as ctk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import threading
import os
import time

# --- PALET WARNA ---
COLOR_SIDEBAR = "#3D4654"
COLOR_MAIN_BG = "#2C333D"
COLOR_ACCENT_YELLOW = "#FAB131"
COLOR_INPUT_BG = "#4A5568"
COLOR_TEXT_WHITE = "#EDF2F7"
COLOR_LOG_BG = "#FFFFFF"

ctk.set_appearance_mode("Dark")

class BotToolFrame(ctk.CTkFrame):
    def __init__(self, master, tool_type, back_command):
        super().__init__(master, fg_color=COLOR_MAIN_BG)
        self.pack(fill="both", expand=True)
        self.tool_type = tool_type 
        
        self.grid_columnconfigure(0, weight=1)
        self.url_var = ctk.StringVar()
        self.url_var.trace_add("write", self.check_url_input)
        self.driver = None
        self.file_path = None

        self.btn_back = ctk.CTkButton(self, text="⬅ KEMBALI", width=80, fg_color="#E74C3C", command=self.safe_back(back_command))
        self.btn_back.grid(row=0, column=0, sticky="nw", padx=10, pady=10)

        self.header = ctk.CTkLabel(self, text=f"GET NUMBER {tool_type} 🧀", font=("Arial", 22, "bold"), text_color=COLOR_ACCENT_YELLOW)
        self.header.grid(row=1, column=0, pady=(0, 10))

        self.entry_url = ctk.CTkEntry(self, textvariable=self.url_var, placeholder_text="Masukkan URL Utama...", height=35, fg_color=COLOR_INPUT_BG, border_width=0)
        self.entry_url.grid(row=2, column=0, padx=30, pady=10, sticky="ew")

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.grid(row=3, column=0, padx=30, pady=5, sticky="ew")
        self.btn_frame.grid_columnconfigure((0,1), weight=1)

        self.btn_browser = ctk.CTkButton(self.btn_frame, text="1. Buka Browser", state="disabled", command=self.run_open_browser_thread)
        self.btn_browser.grid(row=0, column=0, padx=5, sticky="ew")

        self.btn_upload = ctk.CTkButton(self.btn_frame, text="2. Upload Excel", state="disabled", command=self.upload_file)
        self.btn_upload.grid(row=0, column=1, padx=5, sticky="ew")

        self.label_status = ctk.CTkLabel(self, text="Status: Menunggu...", font=("Arial", 12))
        self.label_status.grid(row=4, column=0, pady=(10, 0))
        
        self.progress_bar = ctk.CTkProgressBar(self, height=12, progress_color=COLOR_ACCENT_YELLOW)
        self.progress_bar.grid(row=5, column=0, padx=30, pady=10, sticky="ew")
        self.progress_bar.set(0)

        self.btn_start = ctk.CTkButton(self, text="JALANKAN PROSES! 🐭", fg_color=COLOR_ACCENT_YELLOW, text_color="black", height=45, font=("Arial", 14, "bold"), state="disabled", command=self.run_start_thread)
        self.btn_start.grid(row=6, column=0, padx=30, pady=15, sticky="ew")

        self.info_box = ctk.CTkTextbox(self, fg_color=COLOR_LOG_BG, text_color="#2C333D", corner_radius=8, height=120, font=("Consolas", 11))
        self.info_box.grid(row=7, column=0, padx=30, pady=(0, 20), sticky="nsew")
        self.grid_rowconfigure(7, weight=1)

    def safe_back(self, back_cmd):
        return lambda: [self.driver.quit() if self.driver else None, back_cmd()]

    def check_url_input(self, *args):
        val = self.url_var.get().strip()
        self.btn_browser.configure(state="normal" if len(val) >= 8 else "disabled")

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.file_path = path
            self.update_log(f"[FILE]: {os.path.basename(path)} dimuat!")
            self.btn_start.configure(state="normal")

    def run_open_browser_thread(self):
        threading.Thread(target=self.open_browser_logic, daemon=True).start()

    def open_browser_logic(self):
        try:
            self.update_log("[SISTEM]: Membuka Chrome...")
            self.driver = webdriver.Chrome()
            self.driver.get(self.url_var.get().strip())
            self.btn_upload.configure(state="normal")
            self.update_log("[SISTEM]: Browser siap. Silakan Login dulu.")
        except Exception as e:
            self.update_log(f"[ERROR]: {str(e)}")

    def run_start_thread(self):
        self.btn_start.configure(state="disabled")
        threading.Thread(target=self.start_processing_logic, daemon=True).start()

    def update_log(self, message):
        self.info_box.insert("end", message + "\n")
        self.info_box.see("end")

    def start_processing_logic(self):
        try:
            base_url = self.url_var.get().strip().rstrip('/')
            df = pd.read_excel(self.file_path, header=None)
            if df.shape[1] < 2: df[1] = ""
            total_rows = len(df)

            if self.tool_type == "SGA":
                self.update_log("[NAVIGASI]: Menuju Member Statistik SGA...")
                self.driver.get(f"{base_url}/SPanel/Player/MemberStat")
                time.sleep(3)

                try:
                    dropdown = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div[1]/div[1]/select")))
                    select = Select(dropdown)
                    select.select_by_visible_text("500")
                    self.update_log("[SGA]: Dropdown set ke 500.")
                    time.sleep(2)
                except:
                    self.update_log("[SGA]: Gagal set dropdown 500, lanjut...")

                for index, row in df.iterrows():
                    name = str(df.iloc[index, 0]).strip()
                    if name == "nan" or name == "": continue
                    found = False
                    self.update_log(f"[SGA]: Mencari ID {name}...")
                    try:
                        search_input = self.driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div/input")
                        search_input.clear()
                        search_input.send_keys(name)
                        self.driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div/button[3]/i").click()
                        time.sleep(2)

                        while not found:
                            try:
                                xpath_nama_tabel = "/html/body/div[1]/div/div/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[2]/div[1]/span[1]"
                                nama_di_tabel = self.driver.find_element(By.XPATH, xpath_nama_tabel).text.strip()
                                if nama_di_tabel == name:
                                    data_diambil = self.driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[2]/div[2]/span").text.strip()
                                    df.iloc[index, 1] = data_diambil
                                    self.update_log(f"[HASIL]: {name} -> {data_diambil}")
                                    found = True
                                else:
                                    btn_next = self.driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[2]/div/div[1]/div[2]/button[3]")
                                    if btn_next.is_enabled():
                                        btn_next.click()
                                        time.sleep(1.5)
                                    else:
                                        df.iloc[index, 1] = "Invalid"
                                        self.update_log(f"[SGA]: {name} Tidak ditemukan")
                                        break
                            except:
                                df.iloc[index, 1] = "Invalid"
                                break
                    except:
                        df.iloc[index, 1] = "Error"
                    
                    prog = (index + 1) / total_rows
                    self.progress_bar.set(prog)
                    self.label_status.configure(text=f"Progress: {int(prog * 100)}%")

            else:
                target_xpath = "/html/body/form/div[3]/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/input"
                for index, row in df.iterrows():
                    name = str(df.iloc[index, 0]).strip()
                    if name == "nan" or name == "": continue
                    full_link = f"{base_url}/_SubAg_Sub/PopUpAccInfo.aspx?userName={name}&type=member&role=sa"
                    self.driver.get(full_link)
                    try:
                        element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, target_xpath)))
                        data_diambil = element.get_attribute('value')
                        df.iloc[index, 1] = data_diambil
                        self.update_log(f"[PKV]: {name} -> {data_diambil}")
                    except:
                        df.iloc[index, 1] = "Gagal"
                    
                    prog = (index + 1) / total_rows
                    self.progress_bar.set(prog)
                    self.label_status.configure(text=f"Progress: {int(prog * 100)}%")

            df.to_excel(self.file_path, index=False, header=False)
            self.update_log("[SISTEM]: Selesai & Disimpan.")
            self.label_status.configure(text="Selesai 100% ✅", text_color="#2ECC71")

        except Exception as e:
            self.update_log(f"[CRITICAL ERROR]: {str(e)}")
        finally:
            # AUTO CLOSE BROWSER
            if self.driver:
                try:
                    self.driver.quit()
                    self.driver = None
                    self.update_log("[SISTEM]: Browser ditutup otomatis.")
                except:
                    pass
            self.btn_start.configure(state="normal")

class MainLauncher(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RF-- Garasi Tool")
        self.geometry("450x550")
        self.minsize(400, 550)
        self.current_frame = None
        self.show_main_menu()

    def show_main_menu(self):
        if self.current_frame: self.current_frame.destroy()
        self.current_frame = ctk.CTkFrame(self, fg_color=COLOR_MAIN_BG)
        self.current_frame.pack(fill="both", expand=True)
        self.current_frame.grid_columnconfigure(0, weight=1)
        self.current_frame.grid_rowconfigure((1,2), weight=1)
        ctk.CTkLabel(self.current_frame, text="🐱 TOOL BOT", font=("Comic Sans MS", 28, "bold"), text_color=COLOR_ACCENT_YELLOW).grid(row=0, column=0, pady=(40, 20))
        ctk.CTkButton(self.current_frame, text="SGA TOOL 🐭", font=("Arial", 16, "bold"), fg_color=COLOR_SIDEBAR, hover_color=COLOR_ACCENT_YELLOW, command=lambda: self.open_tool("SGA")).grid(row=1, column=0, padx=40, pady=10, sticky="nsew")
        ctk.CTkButton(self.current_frame, text="PKV TOOL 🧀", font=("Arial", 16, "bold"), fg_color=COLOR_SIDEBAR, hover_color=COLOR_ACCENT_YELLOW, command=lambda: self.open_tool("PKV")).grid(row=2, column=0, padx=40, pady=10, sticky="nsew")

    def open_tool(self, tool_type):
        self.current_frame.destroy()
        self.current_frame = BotToolFrame(self, tool_type, self.show_main_menu)

if __name__ == "__main__":
    app = MainLauncher()
    app.mainloop()