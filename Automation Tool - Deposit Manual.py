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
import sys

# Konfigurasi Tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class AutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automation Tool - Deposit Manual")
        self.geometry("650x550")
        self.driver = None 
        self.is_browser_opened = False

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Login Section ---
        self.label_login = ctk.CTkLabel(self, text="Login:", font=("Arial", 13, "bold"))
        self.label_login.grid(row=0, column=0, padx=20, pady=(25, 10), sticky="w")
        
        self.login_var = ctk.StringVar()
        self.login_var.trace_add("write", self.validate_inputs) 
        self.entry_login = ctk.CTkEntry(self, textvariable=self.login_var, placeholder_text="")
        self.entry_login.grid(row=0, column=1, padx=20, pady=(25, 10), sticky="ew")
        self.entry_login.bind("<FocusOut>", self.autofill_login)

        # --- Deposit Section ---
        self.label_deposit = ctk.CTkLabel(self, text="Deposit Manual:", font=("Arial", 13, "bold"))
        self.label_deposit.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        
        self.deposit_var = ctk.StringVar()
        self.deposit_var.trace_add("write", self.validate_inputs)
        self.entry_deposit = ctk.CTkEntry(self, textvariable=self.deposit_var, placeholder_text="")
        self.entry_deposit.grid(row=1, column=1, padx=20, pady=10, sticky="ew")
        self.entry_deposit.bind("<FocusOut>", self.autofill_deposit)

        # --- Data Table (Text Area) ---
        self.textbox_data = ctk.CTkTextbox(self, height=200, border_width=1)
        self.textbox_data.grid(row=2, column=0, columnspan=2, padx=20, pady=15, sticky="nsew")
        self.textbox_data.bind("<KeyRelease>", self.validate_inputs)

        # --- Button Frame ---
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=3, column=0, columnspan=2, pady=(0, 20))

        self.btn_open_browser = ctk.CTkButton(self.button_frame, text="Open browser", 
                                             command=self.start_browser_thread, width=250, state="disabled")
        self.btn_open_browser.pack(side="left", padx=10)

        self.btn_start = ctk.CTkButton(self.button_frame, text="Start", command=self.start_logic, 
                                       fg_color="#2fa572", state="disabled")
        self.btn_start.pack(side="left", padx=10)

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
        try:
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
        except:
            pass

    def start_browser_thread(self):
        link = self.entry_login.get()
        self.btn_open_browser.configure(state="disabled", text="Opening Browser...")
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
            messagebox.showerror("Error Browser", str(e))
            self.after(0, lambda: self.btn_open_browser.configure(text="Open browser", state="normal"))

    def finalize_browser_state(self):
        self.btn_open_browser.configure(text="Browser Opened", state="disabled")
        self.validate_inputs()

    def start_logic(self):
        raw_data = self.textbox_data.get("1.0", "end-1c").strip()
        lines = raw_data.split('\n')
        template_url = self.deposit_var.get()
        threading.Thread(target=self.run_automation_loop, args=(lines, template_url), daemon=True).start()

    def run_automation_loop(self, lines, template):
        for line in lines:
            line = line.strip()
            if not line: continue
            try:
                parts = line.split()
                user_full = parts[0]
                nominal = "".join(filter(str.isdigit, parts[1])) if len(parts) > 1 else "0"
                user_filtered = user_full[3:] 
                
                final_link = template.replace("{USER}", user_full).replace("{ID}", user_filtered)
                self.driver.get(final_link)
                time.sleep(2)

                wait = WebDriverWait(self.driver, 10)
                xpath_textbox = "/html/body/form/div[3]/table/tbody/tr[3]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/input"
                input_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_textbox)))
                input_element.clear()
                input_element.send_keys(nominal)
                
                xpath_submit = "/html/body/form/div[3]/table/tbody/tr[3]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[10]/td/a[1]"
                btn_submit = self.driver.find_element(By.XPATH, xpath_submit)
                btn_submit.click()
                time.sleep(2)
            except Exception as e:
                print(f"Gagal di baris {line}: {e}")
                continue

if __name__ == "__main__":
    try:
        app = AutomationApp()
        app.mainloop()
    except Exception as e:
        # Jika ada error saat startup, munculkan di CMD
        print(f"CRITICAL ERROR: {e}")