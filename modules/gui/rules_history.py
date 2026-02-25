import customtkinter as ctk
from modules.audit_manager import AuditManager
from modules.gui.utils import bind_context_help # <--- IMPORTED

class RulesHistory(ctk.CTkFrame):
    def __init__(self, parent, rule_manager):
        super().__init__(parent, fg_color="transparent")
        self.manager = rule_manager
        self.audit_manager = AuditManager(self.manager.rule_dir)
        
        # Toolbar
        ctrl = ctk.CTkFrame(self, fg_color="transparent")
        ctrl.pack(fill="x", pady=5)
        
        self.hist_var = ctk.StringVar(value="Select Rule File")
        self.hist_menu = ctk.CTkOptionMenu(ctrl, variable=self.hist_var, command=self._load_log)
        self.hist_menu.pack(side="left", padx=5)
        bind_context_help(self.hist_menu, "Select the Rule Configuration to view its change history.")
        
        self.hist_filter = ctk.CTkEntry(ctrl, placeholder_text="Filter User/Field...")
        self.hist_filter.pack(side="left", fill="x", expand=True, padx=5)
        self.hist_filter.bind("<Return>", lambda e: self._load_log()) 
        bind_context_help(self.hist_filter, "Type a User Name or Field Name and press Enter to filter the logs.")
        
        btn_refresh = ctk.CTkButton(ctrl, text="Refresh", width=80, command=self._load_log)
        btn_refresh.pack(side="right")
        bind_context_help(btn_refresh, "Reload the history log.")
        
        # Text Area
        self.hist_box = ctk.CTkTextbox(self, font=("Consolas", 11), wrap="none")
        self.hist_box.pack(fill="both", expand=True, pady=5)
        
        self._refresh()

    def _refresh(self):
        names = self.manager.get_available_rules()
        if names:
            self.hist_menu.configure(values=names)
            self.hist_var.set(names[0])
            self.after(100, self._load_log)
        else:
            self.hist_menu.configure(values=["No Rules"])
            self.hist_var.set("No Rules")

    def _load_log(self, _=None):
        selection = self.hist_var.get()
        if selection == "No Rules" or selection == "Select Rule File": return

        self.hist_box.configure(state="normal")
        self.hist_box.delete("0.0", "end")

        # Use AuditManager logic
        df = self.audit_manager.get_history_dataframe(
            selection, 
            filter_text=self.hist_filter.get().strip()
        )

        if not df.empty:
            # Format nicely
            text_table = df.to_string(index=False, justify='left', col_space=15)
            self.hist_box.insert("0.0", text_table)
        else:
            self.hist_box.insert("0.0", "No history found (or file is locked/missing).")

        self.hist_box.configure(state="disabled")