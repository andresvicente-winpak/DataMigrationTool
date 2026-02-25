import customtkinter as ctk
import glob
import os
from modules.audit_manager import AuditManager
from modules.gui.utils import bind_context_help # <--- IMPORTED

class RulesTools(ctk.CTkFrame):
    def __init__(self, parent, rule_manager):
        super().__init__(parent, fg_color="transparent")
        ctk.CTkLabel(self, text="System Maintenance", font=("Arial", 18, "bold")).pack(pady=20)
        
        btn_commit = ctk.CTkButton(self, text="Commit All Excel Edits", command=self._commit)
        btn_commit.pack(pady=10)
        bind_context_help(btn_commit, "Scans all Excel rule files for manual edits made outside this tool\nand updates the Audit History log.")
        
        btn_reset = ctk.CTkButton(self, text="HARD RESET SYSTEM", fg_color="red", command=lambda: AuditManager().hard_reset())
        btn_reset.pack(pady=20)
        bind_context_help(btn_reset, "DANGER: Deletes all Audit History and Snapshots.\nUse only if you want to clear the entire change log.")

    def _commit(self):
        for f in glob.glob('config/rules/*.xlsx'): 
            AuditManager().commit_changes(os.path.basename(f).replace('.xlsx',''))
        print("All files audited.")