import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import glob
import threading
from modules.sync_manager import SyncManager
from modules.gui.utils import bind_context_help  # <--- IMPORTED

class SyncHub(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.syncer = SyncManager()
        self.diffs = []
        
        # --- HEADER ---
        ctk.CTkLabel(self, text="Rule Synchronization (Offline Merge)", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        ctk.CTkLabel(self, text="Use this to merge changes from another user's Excel file into your local system.", text_color="gray").pack(anchor="w", pady=(0, 10))

        # --- SELECTION ---
        sel_frame = ctk.CTkFrame(self)
        sel_frame.pack(fill="x", padx=5)
        
        # Local Selection
        ctk.CTkLabel(sel_frame, text="1. Select Local API:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.local_var = ctk.StringVar(value="")
        self.local_menu = ctk.CTkOptionMenu(sel_frame, variable=self.local_var)
        self.local_menu.grid(row=0, column=1, padx=10, sticky="ew")
        bind_context_help(self.local_menu, "Select the Rule Configuration file on your machine that you want to update.")
        
        # Incoming Selection
        ctk.CTkLabel(sel_frame, text="2. Select Incoming File:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.incoming_entry = ctk.CTkEntry(sel_frame, placeholder_text="Path to other user's file...")
        self.incoming_entry.grid(row=1, column=1, padx=10, sticky="ew")
        bind_context_help(self.incoming_entry, "Enter the path to the Excel Rule file containing new changes.")

        btn_browse = ctk.CTkButton(sel_frame, text="Browse", width=80, command=self._browse)
        btn_browse.grid(row=1, column=2, padx=10)
        bind_context_help(btn_browse, "Open file dialog to pick the incoming Excel file.")
        
        sel_frame.grid_columnconfigure(1, weight=1)
        
        # Action
        btn_compare = ctk.CTkButton(self, text="COMPARE FILES", fg_color="orange", text_color="black", height=40, command=self._run_compare)
        btn_compare.pack(fill="x", padx=5, pady=10)
        bind_context_help(btn_compare, "Analyze the two files and list differences (New rules, conflicts, etc.).")

        # --- DIFF VIEW ---
        self.scroll = ctk.CTkScrollableFrame(self, label_text="Detected Differences")
        self.scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # --- FOOTER ---
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.pack(fill="x", pady=10)
        
        btn_all = ctk.CTkButton(foot, text="Select All", width=100, fg_color="gray", command=self._select_all)
        btn_all.pack(side="left", padx=5)
        bind_context_help(btn_all, "Check all detected changes for merging.")

        btn_merge = ctk.CTkButton(foot, text="MERGE SELECTED", width=200, fg_color="green", command=self._run_merge)
        btn_merge.pack(side="right", padx=5)
        bind_context_help(btn_merge, "Apply the checked changes to your local file and merge the audit history.")

        self._refresh_locals()

    def _refresh_locals(self):
        files = glob.glob('config/rules/*.xlsx')
        names = [os.path.basename(f).replace('.xlsx','') for f in files]
        if names:
            self.local_menu.configure(values=names)
            self.local_var.set(names[0])

    def _browse(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f: 
            self.incoming_entry.delete(0, "end")
            self.incoming_entry.insert(0, f)

    def _run_compare(self):
        local = self.local_var.get()
        incoming = self.incoming_entry.get()
        
        if not local or not incoming: return
        
        # Clear list
        for w in self.scroll.winfo_children(): w.destroy()
        
        diffs, err = self.syncer.compare_files(local, incoming)
        if err:
            messagebox.showerror("Error", err)
            return
            
        self.diffs = diffs # Store for merge
        self.check_vars = []
        
        if not diffs:
            ctk.CTkLabel(self.scroll, text="No differences found. Files are identical.").pack(pady=20)
            return

        for i, d in enumerate(diffs):
            # Row Container
            row = ctk.CTkFrame(self.scroll, fg_color="#333333" if i%2==0 else "transparent")
            row.pack(fill="x", pady=2)
            
            # Checkbox
            var = ctk.BooleanVar(value=True)
            self.check_vars.append(var)
            ctk.CTkCheckBox(row, text="", variable=var, width=24).pack(side="left", padx=5)
            
            # Info
            info_frame = ctk.CTkFrame(row, fg_color="transparent")
            info_frame.pack(side="left", fill="x", expand=True)
            
            # Title: TARGET [SCOPE] (TYPE)
            title = f"{d['Target']} [{d['Scope']}] - {d['Type']}"
            color = "#44FF44" if d['Type'] == 'NEW' else "#FFFF44"
            ctk.CTkLabel(info_frame, text=title, text_color=color, font=("Arial", 12, "bold")).pack(anchor="w")
            
            # Detail
            detail = f"Local: {d['Local']}  |  Incoming: {d['Incoming']}"
            ctk.CTkLabel(info_frame, text=detail, text_color="gray", font=("Consolas", 10)).pack(anchor="w")

    def _select_all(self):
        for v in self.check_vars: v.set(True)

    def _run_merge(self):
        selected_indices = [i for i, v in enumerate(self.check_vars) if v.get()]
        if not selected_indices:
            messagebox.showwarning("Warning", "No changes selected.")
            return
            
        to_merge = [self.diffs[i] for i in selected_indices]
        
        if messagebox.askyesno("Confirm Merge", f"Are you sure you want to merge {len(to_merge)} changes?\nThis will update your local rule file."):
            success, msg = self.syncer.perform_merge(self.local_var.get(), self.incoming_entry.get(), to_merge)
            if success:
                messagebox.showinfo("Success", msg)
                self._run_compare() # Refresh view
            else:
                messagebox.showerror("Error", msg)