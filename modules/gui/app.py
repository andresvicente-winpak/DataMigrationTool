import customtkinter as ctk
import tkinter as tk
import sys
import os
import platform
import subprocess
from tkinter import messagebox

# Import Hubs
from modules.gui.tab_migration import MigrationHub
from modules.gui.tab_config import ConfigHub
from modules.gui.tab_rules import RulesHub
from modules.gui.tab_utils import UtilitiesHub
from modules.gui.tab_sync import SyncHub 
from modules.gui.utils import TextRedirector, bind_context_help # <--- IMPORTED HERE

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class M3MigrationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("M3 Data Migration Platform")
        self.geometry("1280x850")
        
        # Main Grid: 2 Columns (Sidebar | Content)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR (Fixed) ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text="M3 MIGRATION", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(30, 20))
        
        # Buttons
        self.btn_mig = ctk.CTkButton(self.sidebar, text="  Run Migration  ", height=40, anchor="w", fg_color="transparent", border_width=1, command=lambda: self.select_frame("MigrateHub"))
        self.btn_mig.pack(fill="x", padx=10, pady=5)
        bind_context_help(self.btn_mig, "Go to the Migration Hub.\n\nUse this to run Standard, Auto-Detect, Surgical, or Batch migration jobs.")
        
        self.btn_conf = ctk.CTkButton(self.sidebar, text="  Configuration  ", height=40, anchor="w", fg_color="transparent", border_width=1, command=lambda: self.select_frame("ConfigHub"))
        self.btn_conf.pack(fill="x", padx=10, pady=5)
        bind_context_help(self.btn_conf, "Go to Configuration.\n\nImport MCO specs, Reverse Engineer data, or edit system maps.")
        
        self.btn_rule = ctk.CTkButton(self.sidebar, text="  Rules & Admin  ", height=40, anchor="w", fg_color="transparent", border_width=1, command=lambda: self.select_frame("RulesHub"))
        self.btn_rule.pack(fill="x", padx=10, pady=5)
        bind_context_help(self.btn_rule, "Go to Rules & Admin.\n\nEdit transformation rules interactively, view history, or perform system maintenance.")
        
        self.btn_util = ctk.CTkButton(self.sidebar, text="  Utilities  ", height=40, anchor="w", fg_color="transparent", border_width=1, command=lambda: self.select_frame("UtilsHub"))
        self.btn_util.pack(fill="x", padx=10, pady=5)
        bind_context_help(self.btn_util, "Go to Utilities.\n\nAccess helper tools like Copy Sheet, Merge Files, or run external Python scripts.")

        # Sync Button
        self.btn_sync = ctk.CTkButton(self.sidebar, text="  Sync / Merge  ", height=40, anchor="w", fg_color="transparent", border_width=1, command=lambda: self.select_frame("SyncHub"))
        self.btn_sync.pack(fill="x", padx=10, pady=5)
        bind_context_help(self.btn_sync, "Go to Sync Hub.\n\nCompare and merge rule changes from other users/files.")

        # Help Button (Pushed to bottom)
        ctk.CTkFrame(self.sidebar, fg_color="transparent").pack(fill="both", expand=True) # Spacer
        self.btn_help = ctk.CTkButton(self.sidebar, text="  📘 User Guide  ", height=40, anchor="w", fg_color="transparent", text_color="#AAAAAA", hover_color="#333333", command=self.open_documentation)
        self.btn_help.pack(fill="x", padx=10, pady=20)
        bind_context_help(self.btn_help, "Opens the PDF User Manual.\n\n(Run pack_project.py to generate the PDF if missing).")

        # --- RIGHT SIDE (Splitter) ---
        self.splitter = tk.PanedWindow(self, orient="vertical", bg="#2B2B2B", bd=0, sashwidth=10, sashrelief="raised", sashpad=0)
        self.splitter.grid(row=0, column=1, sticky="nsew")

        # PANE 1: MAIN CONTENT AREA
        self.main_area = ctk.CTkFrame(self.splitter, corner_radius=0, fg_color="transparent")
        self.splitter.add(self.main_area, stretch="always", height=600)
        
        self.frames = {}
        self.frames["MigrateHub"] = MigrationHub(self.main_area)
        self.frames["ConfigHub"] = ConfigHub(self.main_area)
        self.frames["RulesHub"] = RulesHub(self.main_area)
        self.frames["UtilsHub"] = UtilitiesHub(self.main_area)
        self.frames["SyncHub"] = SyncHub(self.main_area)

        # PANE 2: CONSOLE AREA
        self.console_frame = ctk.CTkFrame(self.splitter, corner_radius=0, fg_color="#1A1A1A")
        self.splitter.add(self.console_frame, stretch="never", height=200)

        con_head = ctk.CTkFrame(self.console_frame, height=28, corner_radius=0, fg_color="#333333")
        con_head.pack(fill="x")
        ctk.CTkLabel(con_head, text="  System Log (Drag bar above to resize)", font=ctk.CTkFont(size=11), text_color="#AAAAAA", anchor="w").pack(side="left", fill="x")
        
        self.console = ctk.CTkTextbox(self.console_frame, font=("Consolas", 12))
        self.console.pack(fill="both", expand=True)
        self.console.configure(state="disabled")
        
        sys.stdout = TextRedirector(self.console)
        sys.stderr = TextRedirector(self.console)

        self.select_frame("MigrateHub")
        
        # --- ADDED LOGGING HERE ---
        cwd = os.getcwd()
        print(f"System Initialized.")
        print(f"Base Directory: {cwd}")
        if os.path.exists(os.path.join(cwd, "config")):
            print("Config folder found.")
        else:
            print("WARNING: Config folder NOT found in base directory.")

    def select_frame(self, name):
        # Reset Button Colors (Keep transparent + border)
        self.btn_mig.configure(fg_color="transparent")
        self.btn_conf.configure(fg_color="transparent")
        self.btn_rule.configure(fg_color="transparent")
        self.btn_util.configure(fg_color="transparent")
        self.btn_sync.configure(fg_color="transparent")

        # Highlight Selected (Solid Fill)
        if name == "MigrateHub": self.btn_mig.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        elif name == "ConfigHub": self.btn_conf.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        elif name == "RulesHub": self.btn_rule.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        elif name == "UtilsHub": self.btn_util.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        elif name == "SyncHub": self.btn_sync.configure(fg_color=["#3B8ED0", "#1F6AA5"])

        # Frame Switching
        for f in self.frames.values(): 
            f.pack_forget()
        self.frames[name].pack(fill="both", expand=True)

    def open_documentation(self):
        """Opens the User Guide PDF in the default system viewer."""
        doc_path = "M3_Migration_User_Guide.pdf"
        
        if not os.path.exists(doc_path):
            messagebox.showinfo("Docs Not Found", "User Guide PDF not found.\nPlease run 'pack_project.py' or 'generate_docs.py' to create it.")
            return

        try:
            if platform.system() == 'Darwin':       # macOS
                subprocess.call(('open', doc_path))
            elif platform.system() == 'Windows':    # Windows
                os.startfile(doc_path)
            else:                                   # linux variants
                subprocess.call(('xdg-open', doc_path))
        except Exception as e:
            messagebox.showerror("Error", f"Could not open documentation: {e}")