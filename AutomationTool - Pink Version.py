import customtkinter as ctk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from tkinter import messagebox
import threading
import time

# Konfigurasi Tema Kawaii
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class AutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Palette Warna Hello Kitty ---
        self.pink_bg = "#FFFBFC"       # Putih kemerahan (Soft)
        self.pink_main = "#FFB6C1"     # Light Pink
        self.pink_dark = "#FF69B4"     # Hot Pink
        self.white = "#FFFFFF"
        self.text_color = "#5D5D5D"

        self.title("Automation Tool - Limited Edition ")
        self.geometry("800x700")
        self.configure(fg_color=self.pink_bg)
        
        self.driver = None 
        self.is_browser_opened = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- HEADER (Pita Pink) ---
        self.header_frame = ctk.CTkFrame(self, height=100, corner_radius=0, fg_color=self.pink_dark)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        
        self.header_label = ctk.CTkLabel(
            self.header_frame, 
            text="✩ 🎀 𝒟𝑒𝓅𝑜𝓈𝒾𝓉 𝑀𝒶𝓃𝓊𝒶𝓁 🎀 ✩", 
            font=("Comic Sans MS", 32, "bold"), 
            text_color=self.white
        )
        self.header_label.pack(pady=25)

        # --- MAIN CONTAINER ---
        self.main_container = ctk.CTkFrame(self, fg_color=self.white, corner_radius=20, border_color=self.pink_main, border_width=2)
        self.main_container.grid(row=1, column=0, padx=30, pady=20, sticky="nsew")
        self.main_container.grid_columnconfigure(0, weight=1)

        # 1. Konfigurasi Link
        self.config_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.config_frame.grid(row=0, column=0, sticky="ew", pady=(20, 10), padx=20)
        self.config_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self.config_frame, 
            text="✧˚ · . 𝖢𝗈𝗇𝖿𝗂𝗀𝗎𝗋𝖺𝗍𝗂𝗈𝗇 𝖲𝖾𝗍𝗍𝗂𝗇𝗀𝗌 . · ˚✧", 
            font=("Comic Sans MS", 18, "bold"), 
            text_color=self.pink_dark
        ).grid(row=0, column=0, columnspan=2, pady=(0, 15))

        # Login Input
        ctk.CTkLabel(self.config_frame, text="💌 Login URL :", font=("Comic Sans MS", 13, "bold"), text_color=self.text_color).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.login_var = ctk.StringVar()
        self.login_var.trace_add("write", self.validate_inputs)
        self.entry_login = ctk.CTkEntry(self.config_frame, textvariable=self.login_var, border_color=self.pink_main, fg_color="#FFF9FB", height=35)
        self.entry_login.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.entry_login.bind("<FocusOut>", self.autofill_login)

        # Deposit Input
        ctk.CTkLabel(self.config_frame, text="💰 Depo URL  :", font=("Comic Sans MS", 13, "bold"), text_color=self.text_color).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.deposit_var = ctk.StringVar()
        self.deposit_var.trace_add("write", self.validate_inputs)
        self.entry_deposit = ctk.CTkEntry(self.config_frame, textvariable=self.deposit_var, border_color=self.pink_main, fg_color="#FFF9FB", height=35)
        self.entry_deposit.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.entry_deposit.bind("<FocusOut>", self.autofill_deposit)

        # 2. Input Data
        self.data_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.data_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        self.data_frame.grid_columnconfigure(0, weight=1)
        self.data_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self.data_frame, 
            text="⸜(｡˃ ᵕ ˂ )⸝♡ Drop your data below! ♡", 
            font=("Comic Sans MS", 14, "italic"), 
            text_color=self.pink_dark
        ).grid(row=0, column=0, pady=5)
        
        self.textbox_data = ctk.CTkTextbox(
            self.data_frame, 
            border_color=self.pink_main, 
            border_width=2, 
            fg_color="#FFF9FB", 
            font=("Consolas", 12),
            corner_radius=15
        )
        self.textbox_data.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.textbox_data.bind("<KeyRelease>", self.validate_inputs)

        # --- BUTTONS ---
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=2, column=0, pady=25)

        self.btn_open_browser = ctk.CTkButton(
            self.action_frame, text="🌸 OPEN BROWSER", 
            command=self.start_browser_thread, 
            fg_color=self.pink_dark, hover_color="#DB7093",
            width=220, height=50, font=("Comic Sans MS", 14, "bold"), corner_radius=30
        )
        self.btn_open_browser.pack(side="left", padx=15)

        self.btn_start = ctk.CTkButton(
            self.action_frame, text="🐾 START PROCESS", 
            command=self.start_logic, 
            fg_color="#FF82AB", hover_color="#FF1493",
            width=220, height=50, font=("Comic Sans MS", 14, "bold"), corner_radius=30, state="disabled"
        )
        self.btn_start.pack(side="left", padx=15)

        # --- STATUS BAR ---
        self.status_bar = ctk.CTkLabel(
            self, text="Ready to sparkle! ✨", 
            anchor="center", fg_color=self.pink_dark, 
            text_color="white", height=35, font=("Comic Sans MS", 12)
        )
        self.status_bar.grid(row=3, column=0, sticky="ew")

    # --- Logika Fungsi ---
    def update_status(self, text):
        self.status_bar.configure(text=f"Status: {text}")

    def autofill_login(self, event):
        url = self.login_var.get().strip()
        if url and not any(x in url for x in ["Default1.aspx", "Main.aspx"]):
            if not url.endswith("/"): url += "/"
            self.login_var.set(f"{url}Public/Default1.aspx")

    def autofill_deposit(self, event):
        url = self.deposit_var.get().strip()
        if url and "DepositManual.aspx" not in url:
            if not url.endswith("/"): url += "/"
            new_url = f"{url}_SubAg_Sub/DepositManual.aspx?role=sa&userName={{USER}}&pg=addCreditRequest&search={{ID}}"
            self.deposit_var.set(new_url)

    def validate_inputs(self, *args):
        login_fill = self.login_var.get().strip()
        depo_fill = self.deposit_var.get().strip()
        table_fill = self.textbox_data.get("1.0", "end-1c").strip()

        if login_fill and not self.is_browser_opened:
            self.btn_open_browser.configure(state="normal")
        else:
            self.btn_open_browser.configure(state="disabled")

        if self.is_browser_opened and depo_fill and table_fill:
            self.btn_start.configure(state="normal")
        else:
            self.btn_start.configure(state="disabled")

    def start_browser_thread(self):
        link = self.entry_login.get()
        self.update_status("Magical gates opening... 🪄💖")
        threading.Thread(target=self.open_browser, args=(link,), daemon=True).start()

    def open_browser(self, url):
        try:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("detach", True)
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            self.driver.get(url)
            self.is_browser_opened = True
            self.after(0, self.finalize_browser_state)
        except Exception as e:
            messagebox.showerror("Oh No! 😿", str(e))
            self.after(0, lambda: self.btn_open_browser.configure(text="🌸 OPEN BROWSER", state="normal"))

    def finalize_browser_state(self):
        self.btn_open_browser.configure(text="💖 BROWSER ON", state="disabled")
        self.update_status("Browser is ready! Let's work! 🎀")
        self.validate_inputs()

    def start_logic(self):
        raw_data = self.textbox_data.get("1.0", "end-1c").strip()
        lines = raw_data.split('\n')
        template_url = self.deposit_var.get()
        self.btn_start.configure(state="disabled", text="🎀 WORKING...")
        threading.Thread(target=self.run_automation_loop, args=(lines, template_url), daemon=True).start()

    def run_automation_loop(self, lines, template):
        total = len([l for l in lines if l.strip()])
        count = 0
        for i, line in enumerate(lines):
            line = line.strip()
            if not line: continue
            
            count += 1
            self.update_status(f"Cooking magic for data {count}/{total}... 🍬")
            
            try:
                parts = line.split()
                user_full = parts[0]
                nominal = "".join(filter(str.isdigit, parts[1])) if len(parts) > 1 else "0"
                user_filtered = user_full[3:] 
                
                final_link = template.replace("{USER}", user_full).replace("{ID}", user_filtered)
                self.driver.get(final_link)
                time.sleep(1.5)

                wait = WebDriverWait(self.driver, 7)
                xpath_textbox = "/html/body/form/div[3]/table/tbody/tr[3]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input"
                input_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_textbox)))
                input_element.clear()
                input_element.send_keys(nominal)
                
                xpath_submit = "/html/body/form/div[3]/table/tbody/tr[3]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[10]/td/a[1]"
                btn_submit = self.driver.find_element(By.XPATH, xpath_submit)
                btn_submit.click()
                time.sleep(1.5)
            except Exception:
                continue
        
        self.update_status("Mission accomplished! You're amazing! 🎀✨")
        self.after(0, lambda: self.btn_start.configure(state="normal", text="🐾 START PROCESS"))

if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()