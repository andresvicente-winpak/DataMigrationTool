import customtkinter as ctk
from modules.rule_manager import RuleManager
from modules.gui.rules_editor import RulesEditor
from modules.gui.rules_history import RulesHistory
from modules.gui.rules_tools import RulesTools
from modules.gui.utils import bind_context_help # <--- IMPORT

class RulesHub(ctk.CTkTabview):
    def __init__(self, parent):
        super().__init__(parent)
        self.manager = RuleManager()
        
        self.add("Editor")
        self.add("History")
        self.add("Tools")
        
        # --- BIND HELP TO TABS ---
        try:
            btns = self._segmented_button._buttons_dict
            if "Editor" in btns: bind_context_help(btns["Editor"], "Rule Editor: Interactive form for modifying transformation rules.")
            if "History" in btns: bind_context_help(btns["History"], "Audit History: View logs of who changed what and when.")
            if "Tools" in btns: bind_context_help(btns["Tools"], "System Tools: Commit edits, reset history, and maintenance.")
        except: pass
        
        RulesEditor(self.tab("Editor"), self.manager).pack(fill="both", expand=True)
        RulesHistory(self.tab("History"), self.manager).pack(fill="both", expand=True)
        RulesTools(self.tab("Tools"), self.manager).pack(fill="both", expand=True)