import customtkinter as ctk
from tkinter import messagebox, filedialog
import json
import os
import threading
import time
import datetime
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURASI TEMA ---
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
        
        # Color Palette
        self.color_bg = "#0A0A0B"
        self.color_card = "#141417"
        self.color_main = "#A855F7"  
        self.color_main_hover = "#9333EA"
        self.color_dark_purple = "#2E1065"
        self.color_accent = "#00F5FF" 
        self.color_error = "#EF4444"
        
        self.configure(fg_color=self.color_bg)
        self.bots = {} 
        self.col_weights = [12, 12, 5, 12, 12, 15, 25] 

        self.setup_ui()
        self.refresh_link_list()
        self.refresh_config_list()
        self.load_session() 
        
        self.add_log("System initialized. Welcome and be Smart.", "SYSTEM", "blue")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    # --- LOGIKA PENYELAMAT (CLEANUP & THREAD SAFE) ---
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

    def stop_all_drivers(self):
        for rid, b in self.bots.items():
            if b.get('driver'):
                try: b['driver'].quit()
                except: pass

    # --- VALIDASI INPUT ---
    def check_cfg_inputs(self):
        all_valid = True
        text_cols = ["Name Col", "Nominal Col", "Username Col", "Status Col"]
        num_cols = ["Max", "Timeout (m)", "DupTime (m)"]

        for label in text_cols:
            en = self.cfg_entries[label]
            val = en.get().strip()
            if val and re.match(r"^[a-zA-Z]+$", val):
                en.configure(text_color="white")
            else:
                en.configure(text_color=self.color_error)
                all_valid = False

        for label in num_cols:
            en = self.cfg_entries[label]
            val = en.get().strip()
            if val and val.isdigit():
                en.configure(text_color="white")
            else:
                en.configure(text_color=self.color_error)
                all_valid = False

        if any(not v.get().strip() for v in self.cfg_entries.values()):
            all_valid = False

        if all_valid:
            self.btn_save_cfg.configure(state="normal", fg_color=self.color_main)
        else:
            self.btn_save_cfg.configure(state="disabled", fg_color="#52525B")

    def validate_only_numbers(self, P):
        if str.isdigit(P) or P == "": return True
        return False

    # --- UI SETUP ---
    def setup_ui(self):
        # HEADER
        self.header = ctk.CTkFrame(self, height=140, fg_color=self.color_card, corner_radius=0, border_width=1, border_color="#27272A")
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)
        
        logo_f = ctk.CTkFrame(self.header, fg_color="transparent")
        logo_f.pack(side="left", padx=50)
        ctk.CTkLabel(logo_f, text="Automation Bot V1", font=("Impact", 56), text_color=self.color_main).pack(side="top", anchor="w")
        ctk.CTkLabel(logo_f, text="By RFer-X Garasi", font=("Consolas", 13, "bold"), text_color=self.color_accent).pack(side="top", anchor="w", padx=5)
        
        domain_container = ctk.CTkFrame(self.header, fg_color=self.color_bg, corner_radius=12, height=60, width=550)
        domain_container.pack(side="right", padx=50, pady=40); domain_container.pack_propagate(False)
        domain_container.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(domain_container, text="🌐", font=("Inter", 20)).grid(row=0, column=0, padx=(20, 5))
        self.global_domain = ctk.CTkEntry(domain_container, width=500, height=55, fg_color="transparent", border_width=0, font=("Consolas", 16, "bold"), text_color=self.color_main, placeholder_text="Input Domain URL.........")
        self.global_domain.grid(row=0, column=1, padx=10, sticky="ew")

        # SIDEBAR
        self.sidebar = ctk.CTkFrame(self, width=280, fg_color=self.color_card, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        ctk.CTkLabel(self.sidebar, text="BOTS STATUS", font=("Inter", 12, "bold"), text_color=self.color_main).pack(pady=(25, 10), padx=20, anchor="w")
        self.status_container = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.status_container.pack(fill="both", expand=True, padx=5, pady=5)

        # CONTENT AREA
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(side="right", expand=True, fill="both", padx=15, pady=15)
        self.tabview = ctk.CTkTabview(content, segmented_button_selected_color=self.color_main, fg_color=self.color_card)
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_dash = self.tabview.add("LIVE MONITOR")
        self.tab_link = self.tabview.add("CONNECT SHEETS")
        self.tab_config = self.tabview.add("RULES ENGINE")
        self.tab_run = self.tabview.add("BOT INSTANCES")

        self.setup_dashboard(); self.setup_link_tab(); self.setup_config_tab(); self.setup_running_tab()

    def setup_dashboard(self):
        self.log_box = ctk.CTkTextbox(self.tab_dash, fg_color="#050505", font=("Consolas", 14), border_width=1)
        self.log_box.pack(fill="both", expand=True, padx=15, pady=15)

    def setup_link_tab(self):
        container = ctk.CTkFrame(self.tab_link, fg_color="transparent"); container.pack(fill="both", expand=True, padx=30, pady=20)
        top_bar = ctk.CTkFrame(container, fg_color="transparent"); top_bar.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(top_bar, text="🔄 REFRESH DATABASE", width=120, fg_color=self.color_dark_purple, command=self.refresh_link_list).pack(side="right")
        
        input_card = ctk.CTkFrame(container, fg_color="#1C1C20", border_width=1, border_color="#27272A"); input_card.pack(fill="x", pady=10, padx=20)
        self.link_name = ctk.CTkEntry(input_card, width=250, placeholder_text="Sheet Category", height=40); self.link_name.grid(row=0, column=0, padx=10, pady=20)
        self.link_url = ctk.CTkEntry(input_card, width=500, placeholder_text="Full Google Sheet URL", height=40); self.link_url.grid(row=0, column=1, padx=10, pady=20)
        self.btn_save_link = ctk.CTkButton(input_card, text="SAVE TO DATABASE", fg_color=self.color_main, height=40, command=self.save_link_json, state="disabled"); self.btn_save_link.grid(row=0, column=2, padx=10, pady=20)
        self.link_name.bind("<KeyRelease>", lambda e: self.check_link_inputs()); self.link_url.bind("<KeyRelease>", lambda e: self.check_link_inputs())
        self.link_list_frame = ctk.CTkScrollableFrame(container, fg_color="transparent"); self.link_list_frame.pack(fill="both", expand=True, pady=20)

    def setup_config_tab(self):
        container = ctk.CTkFrame(self.tab_config, fg_color="transparent"); container.pack(fill="both", expand=True, padx=30, pady=20)
        top_bar = ctk.CTkFrame(container, fg_color="transparent"); top_bar.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(top_bar, text="🔄 REFRESH ENGINE", width=120, fg_color=self.color_dark_purple, command=self.refresh_config_list).pack(side="right")
        
        input_grid = ctk.CTkFrame(container, fg_color="#1C1C20", border_width=1, border_color="#27272A"); input_grid.pack(fill="x", padx=20, pady=10)
        self.cfg_entries = {}
        labels = ["Name Col", "Nominal Col", "Username Col", "Status Col", "Max", "Timeout (m)", "DupTime (m)"]
        for i, label in enumerate(labels):
            r, c = (0, i) if i < 4 else (2, i-4)
            ctk.CTkLabel(input_grid, text=label.upper(), font=("Inter", 10, "bold"), text_color="gray").grid(row=r, column=c, padx=15, pady=(15,0), sticky="w")
            en = ctk.CTkEntry(input_grid, width=160, height=35)
            en.grid(row=r+1, column=c, padx=15, pady=(0, 15), sticky="w")
            en.bind("<KeyRelease>", lambda e: self.check_cfg_inputs())
            self.cfg_entries[label] = en
            
        self.btn_save_cfg = ctk.CTkButton(input_grid, text="REGISTER NEW RULE", fg_color=self.color_main, width=200, height=40, command=self.save_cfg_json, state="disabled"); self.btn_save_cfg.grid(row=3, column=3, pady=20, padx=15, sticky="e")
        self.cfg_list_frame = ctk.CTkScrollableFrame(container, fg_color="transparent"); self.cfg_list_frame.pack(fill="both", expand=True, pady=10)

    def setup_running_tab(self):
        top = ctk.CTkFrame(self.tab_run, fg_color="transparent"); top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="+ DEPLOY NEW BOT", fg_color=self.color_main, height=45, font=("Inter", 13, "bold"), command=self.add_bot_row).pack(side="left")
        ctk.CTkButton(top, text="Refresh All Assets", fg_color=self.color_dark_purple, width=150, height=45, command=self.refresh_all_bot_dropdowns).pack(side="right")
        self.h_row = ctk.CTkFrame(self.tab_run, fg_color="#1F1F23", height=45); self.h_row.pack(fill="x", padx=20, pady=(10, 0))
        headers = ["IDENTIFIER", "SHEET", "ROW", "CONFIG", "DATA SOURCE", "JSON KEY", "COMMANDS"]
        for i, txt in enumerate(headers):
            self.h_row.grid_columnconfigure(i, weight=self.col_weights[i])
            ctk.CTkLabel(self.h_row, text=txt, font=("Inter", 10, "bold"), text_color="gray").grid(row=0, column=i, padx=5, sticky="nsew")
        self.run_container = ctk.CTkScrollableFrame(self.tab_run, fg_color="transparent"); self.run_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def add_bot_row(self, saved_info=None):
        rid = f"Bot_{int(time.time() * 1000)}"
        row_card = ctk.CTkFrame(self.run_container, fg_color="#16161A", border_color="#27272A", border_width=1); row_card.pack(fill="x", pady=5)
        for i, w in enumerate(self.col_weights): row_card.grid_columnconfigure(i, weight=w)
        vcmd = (self.register(self.validate_only_numbers), '%P')

        n_en = ctk.CTkEntry(row_card, height=35, fg_color="#0A0A0B"); n_en.grid(row=0, column=0, padx=8, pady=12, sticky="ew")
        s_en = ctk.CTkEntry(row_card, height=35, fg_color="#0A0A0B"); s_en.grid(row=0, column=1, padx=8, pady=12, sticky="ew")
        r_en = ctk.CTkEntry(row_card, height=35, fg_color="#0A0A0B", validate="key", validatecommand=vcmd); r_en.insert(0, "2"); r_en.grid(row=0, column=2, padx=8, pady=12, sticky="ew")
        cfg_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("cfg_")]), height=35, fg_color="#1F1F23"); cfg_dd.set("Select"); cfg_dd.grid(row=0, column=3, padx=8, pady=12, sticky="ew")
        lnk_dd = ctk.CTkOptionMenu(row_card, values=["Select"]+sorted([f for f in os.listdir() if f.startswith("link_")]), height=35, fg_color="#1F1F23"); lnk_dd.set("Select"); lnk_dd.grid(row=0, column=4, padx=8, pady=12, sticky="ew")
        
        j_f = ctk.CTkFrame(row_card, fg_color="transparent"); j_f.grid(row=0, column=5, padx=8, pady=12, sticky="ew")
        j_en = ctk.CTkEntry(j_f, height=35, fg_color="#0A0A0B"); j_en.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(j_f, text="📂", width=35, height=35, command=lambda r=rid: self.browse_json_path(r)).pack(side="left", padx=(2,0))
        
        btn_c = ctk.CTkFrame(row_card, fg_color="transparent"); btn_c.grid(row=0, column=6, padx=8, pady=12, sticky="ew")
        b_web = ctk.CTkButton(btn_c, text="WEB", width=65, state="disabled", fg_color="#52525B", command=lambda r=rid: self.bot_open_ui(r)); b_web.pack(side="left", padx=2)
        b_start = ctk.CTkButton(btn_c, text="START", width=60, state="disabled", fg_color="#166534", command=lambda r=rid: self.bot_start_ui(r)); b_start.pack(side="left", padx=2)
        b_stop = ctk.CTkButton(btn_c, text="STOP", width=60, state="disabled", fg_color="#991B1B", command=lambda r=rid: self.bot_stop_ui(r)); b_stop.pack(side="left", padx=2)
        ctk.CTkButton(btn_c, text="×", width=35, fg_color="transparent", hover_color="#EF4444", text_color="gray", command=lambda r=rid, f=row_card: self.bot_del(r, f)).pack(side="left", padx=2)
        
        status_lbl = ctk.CTkLabel(self.status_container, text=f"IDLE ➜ {rid}", font=("Consolas", 11), text_color="#52525B", anchor="w")
        status_lbl.pack(fill="x", padx=15, pady=8)
        
        self.bots[rid] = {'n_en': n_en, 's_en': s_en, 'r_en': r_en, 'cfg_dd': cfg_dd, 'lnk_dd': lnk_dd, 'j_en': j_en, 'b_web': b_web, 'b_start': b_start, 'b_stop': b_stop, 'status_lbl': status_lbl, 'driver': None, 'is_running': False, 'browser_ready': False}
        if saved_info:
            n_en.insert(0, saved_info.get('identifier', '')); s_en.insert(0, saved_info.get('sheet', '')); r_en.delete(0, 'end'); r_en.insert(0, saved_info.get('row', '2'))
            if saved_info.get('config') in cfg_dd.cget("values"): cfg_dd.set(saved_info['config'])
            if saved_info.get('link') in lnk_dd.cget("values"): lnk_dd.set(saved_info['link'])
            j_en.insert(0, saved_info.get('json_path', ''))

        n_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); s_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); j_en.bind("<KeyRelease>", lambda e, r=rid: self.lock_logic(r)); self.lock_logic(rid)

    # --- BOT CONTROL ---
    def lock_logic(self, rid):
        b = self.bots[rid]; name = b['n_en'].get().strip()
        is_dup = any(name == ob['n_en'].get().strip() for oid, ob in self.bots.items() if oid != rid)
        b['n_en'].configure(text_color=self.color_error if (is_dup and name) else "white")
        ready = all([name, not is_dup, b['s_en'].get().strip(), b['j_en'].get().strip(), b['cfg_dd'].get() != "Select", b['lnk_dd'].get() != "Select"])
        b['b_web'].configure(state="normal" if ready else "disabled", fg_color=self.color_main if ready else "#52525B")

    def bot_open_ui(self, rid):
        self.save_session(); b = self.bots[rid]; b['b_web'].configure(state="disabled")
        b['status_lbl'].configure(text="LAUNCHING...", text_color=self.color_accent)
        threading.Thread(target=self.open_browser_task, args=(rid,), daemon=True).start()

    def open_browser_task(self, rid):
        b = self.bots[rid]
        try:
            url = self.global_domain.get().strip()
            service = Service(ChromeDriverManager().install())
            b['driver'] = webdriver.Chrome(service=service)
            b['driver'].get(url if url.startswith("http") else "https://"+url)
            self.after(0, lambda: b['status_lbl'].configure(text=f"READY ➜ {b['n_en'].get()}", text_color=self.color_accent))
            self.after(0, lambda: b['b_start'].configure(state="normal"))
            while b['driver']: time.sleep(1); _ = b['driver'].window_handles
        except Exception: b['driver'] = None
        finally:
            self.after(0, lambda: [b['status_lbl'].configure(text=f"IDLE ➜ {rid}", text_color="#52525B"), b['b_web'].configure(state="normal"), b['b_start'].configure(state="disabled")])

    def bot_start_ui(self, rid):
        self.save_session(); b = self.bots[rid]; b['is_running'] = True
        b['b_start'].configure(state="disabled"); b['b_stop'].configure(state="normal")
        b['status_lbl'].configure(text=f"RUNNING ➜ {b['n_en'].get()}", text_color="#10B981")

    def bot_stop_ui(self, rid):
        b = self.bots[rid]; b['is_running'] = False
        b['b_stop'].configure(state="disabled"); b['b_start'].configure(state="normal")
        b['status_lbl'].configure(text=f"HALTED ➜ {b['n_en'].get()}", text_color="#F59E0B")

    def bot_del(self, rid, frame):
        if self.bots[rid]['driver']: 
            try: self.bots[rid]['driver'].quit()
            except: pass
        self.bots[rid]['status_lbl'].destroy(); del self.bots[rid]; frame.destroy(); self.save_session()

    # --- REFRESH LOGIC (FULL DETAIL) ---
    def refresh_config_list(self):
        for w in self.cfg_list_frame.winfo_children(): w.destroy()
        # Header Daftar
        h = ctk.CTkFrame(self.cfg_list_frame, fg_color="transparent")
        h.pack(fill="x", pady=(0,5))
        ctk.CTkLabel(h, text="DETAILED RULE CONFIGURATIONS", font=("Inter", 11, "bold"), text_color=self.color_main).pack(side="left", padx=10)
        
        for fn in sorted(os.listdir()):
            if fn.startswith("cfg_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.cfg_list_frame, fg_color="#1C1C20", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=5, padx=5)
                    
                    # Teks Detail Lengkap
                    detail_text = (
                        f"📁 FILE: {fn}\n"
                        f"------------------------------------------------------------------------------------------------------------------------\n"
                        f"📍 KOLOM NAMA: {d.get('Name Col')}  |  📍 KOLOM NOMINAL: {d.get('Nominal Col')}  |  📍 KOLOM USERNAME: {d.get('Username Col')}  |  📍 KOLOM STATUS: {d.get('Status Col')}\n"
                        f"⚙️ BATAS MAKSIMAL: {d.get('Max')} Baris  |  ⏳ TIMEOUT: {d.get('Timeout (m)')} Menit  |  🔄 CEK DUPLIKAT: {d.get('DupTime (m)')} Menit"
                    )
                    
                    ctk.CTkLabel(c, text=detail_text, font=("Consolas", 11), justify="left", text_color="#E2E8F0").pack(side="left", padx=20, pady=15)
                    ctk.CTkButton(c, text="HAPUS", width=100, height=40, fg_color="#450a0a", hover_color="#EF4444", command=lambda f=fn: [os.remove(f), self.refresh_config_list(), self.refresh_all_bot_dropdowns()]).pack(side="right", padx=20)
                except: pass

    def refresh_link_list(self):
        for w in self.link_list_frame.winfo_children(): w.destroy()
        h = ctk.CTkFrame(self.link_list_frame, fg_color="transparent")
        h.pack(fill="x", pady=(0,5))
        ctk.CTkLabel(h, text="REGISTERED DATA SOURCES (GOOGLE SHEETS)", font=("Inter", 11, "bold"), text_color=self.color_main).pack(side="left", padx=10)
        
        for fn in sorted(os.listdir()):
            if fn.startswith("link_") and fn.endswith(".json"):
                try:
                    with open(fn, "r") as f: d = json.load(f)
                    c = ctk.CTkFrame(self.link_list_frame, fg_color="#1C1C20", border_width=1, border_color=self.color_main)
                    c.pack(fill="x", pady=5, padx=5)
                    
                    # Teks Detail Lengkap
                    detail_text = (
                        f"🏷️ NAMA KATEGORI : {d.get('name')}\n"
                        f"🔗 URL GOOGLE SHEET: {d.get('url')}"
                    )
                    
                    ctk.CTkLabel(c, text=detail_text, font=("Consolas", 11), justify="left", text_color="#E2E8F0").pack(side="left", padx=20, pady=15)
                    ctk.CTkButton(c, text="HAPUS", width=100, height=40, fg_color="#450a0a", hover_color="#EF4444", command=lambda f=fn: [os.remove(f), self.refresh_link_list(), self.refresh_all_bot_dropdowns()]).pack(side="right", padx=20)
                except: pass

    # --- HELPERS ---
    def save_session(self):
        data = {rid: {"identifier": b['n_en'].get(), "sheet": b['s_en'].get(), "row": b['r_en'].get(), "config": b['cfg_dd'].get(), "link": b['lnk_dd'].get(), "json_path": b['j_en'].get()} for rid, b in self.bots.items()}
        with open("session_bots.json", "w") as f: json.dump(data, f, indent=4)

    def load_session(self):
        if os.path.exists("session_bots.json"):
            try:
                with open("session_bots.json", "r") as f:
                    for rid, info in json.load(f).items(): self.add_bot_row(saved_info=info)
            except: pass

    def refresh_all_bot_dropdowns(self):
        cf = ["Select"] + sorted([f for f in os.listdir() if f.startswith("cfg_")])
        ln = ["Select"] + sorted([f for f in os.listdir() if f.startswith("link_")])
        for b in self.bots.values(): b['cfg_dd'].configure(values=cf); b['lnk_dd'].configure(values=ln)

    def check_link_inputs(self): self.btn_save_link.configure(state="normal" if (self.link_name.get().strip() and self.link_url.get().strip()) else "disabled")
    
    def browse_json_path(self, r):
        fn = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if fn: self.bots[r]['j_en'].delete(0, 'end'); self.bots[r]['j_en'].insert(0, fn); self.lock_logic(r)

    def save_cfg_json(self):
        d = {k: v.get().strip() for k, v in self.cfg_entries.items()}
        n_col = d.get('Name Col', '')
        m_col = d.get('Nominal Col', '')
        u_col = d.get('Username Col', '')
        s_col = d.get('Status Col', '')
        raw_name = f"{n_col}_{m_col}_{u_col}_{s_col}"
        clean_name = re.sub(r'[^a-zA-Z0-9_]', '', raw_name)
        file_name = f"cfg_{clean_name}.json"
        with open(file_name, "w") as j:
            json.dump(d, j, indent=4)
        for e in self.cfg_entries.values():
            e.delete(0, 'end')
        self.add_log(f"Config baru berhasil didaftarkan: {file_name}", "SYSTEM", "green")
        self.refresh_config_list()
        self.refresh_all_bot_dropdowns()
        self.check_cfg_inputs()
        
    def save_link_json(self):
        d = {"name": self.link_name.get().strip(), "url": self.link_url.get().strip()}
        with open(f"link_{d['name']}.json", "w") as j: json.dump(d, j, indent=4)
        self.link_name.delete(0, 'end'); self.link_url.delete(0, 'end')
        self.refresh_link_list(); self.refresh_all_bot_dropdowns()

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Tutup aplikasi dan matikan semua bot?"):
            self.stop_all_drivers(); self.save_session(); self.destroy(); os._exit(0)

if __name__ == "__main__":
    app = AutomationBotApp(); app.mainloop()
