import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import glob
import os
from colorama import Fore, Style
from modules.audit_manager import AuditManager
from modules.ui import select_file, MultiSelectPopup
from modules.gui.utils import bind_context_help

# --- COMPONENT 1: THE LIST PANEL (LEFT) ---
class RuleListPanel(ctk.CTkFrame):
    def __init__(self, parent, on_select_callback):
        super().__init__(parent, corner_radius=0, fg_color="transparent")
        self.on_select = on_select_callback
        self.df = None
        self.active_type_filters = [] # Default empty = All
        
        # 1. Filters Row 1 (Search + Scope)
        filter_frame = ctk.CTkFrame(self, fg_color="transparent")
        filter_frame.pack(fill="x", pady=(0, 5))
        
        self.search_var = ctk.StringVar()
        self.search_var.trace("w", self._filter_list)
        self.search_entry = ctk.CTkEntry(filter_frame, placeholder_text="Search Field...", textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True)
        bind_context_help(self.search_entry, "Type to filter the list by Target Field or Description.")
        
        self.scope_var = ctk.StringVar(value="ALL")
        self.scope_menu = ctk.CTkOptionMenu(filter_frame, variable=self.scope_var, width=90, command=lambda x: self._filter_list())
        self.scope_menu.pack(side="right", padx=(5,0))
        bind_context_help(self.scope_menu, "Filter the list by Rule Scope (Global vs Specific Division).")

        # 2. Filters Row 2 (Type Filter)
        type_frame = ctk.CTkFrame(self, fg_color="transparent")
        type_frame.pack(fill="x", pady=(0, 5))
        
        self.btn_type_filter = ctk.CTkButton(type_frame, text="Filter Types (All)", fg_color="#444", height=24, command=self._open_type_filter)
        self.btn_type_filter.pack(fill="x")
        bind_context_help(self.btn_type_filter, "Click to filter by Rule Type (e.g., Show only CONST or TODO items).")

        # 3. Header
        header_frame = ctk.CTkFrame(self, height=30, corner_radius=4, fg_color="#202020")
        header_frame.pack(fill="x")
        ctk.CTkLabel(header_frame, text=f"{'TARGET':<10} {'DESCRIPTION':<22} {'TYPE':<8} {'SOURCE/VAL'}", 
                     font=("Consolas", 11, "bold"), anchor="w").pack(side="left", padx=10)

        # 4. Scrollable List
        self.scroll = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.scroll.pack(fill="both", expand=True, pady=(2, 0))

    def load_data(self, df, scopes):
        self.df = df
        current_scope = self.scope_var.get()
        # Keep current scope if valid, else reset
        valid_scopes = ["ALL"] + scopes
        if current_scope not in valid_scopes:
            self.scope_var.set("ALL")
        
        self.scope_menu.configure(values=valid_scopes)
        self._filter_list()

    def _open_type_filter(self):
        options = ["DIRECT", "CONST", "MAP", "PYTHON", "IGNORE", "TODO", "FILTER"]
        
        def on_select(selected):
            self.active_type_filters = selected
            if not selected:
                self.btn_type_filter.configure(text="Filter Types (All)", fg_color="#444")
            else:
                self.btn_type_filter.configure(text=f"Filter: {','.join(selected)}", fg_color="orange", text_color="black")
            self._filter_list()

        current = self.active_type_filters if self.active_type_filters else []
        MultiSelectPopup(self, options, current, on_select)

    def _filter_list(self, *args):
        # Clear existing buttons
        for w in self.scroll.winfo_children(): w.destroy()
        if self.df is None: return

        search = self.search_var.get().upper().strip()
        scope_filter = self.scope_var.get()
        
        count = 0
        for idx, row in self.df.iterrows():
            # Safe string conversion and casing
            tgt = str(row.get('TARGET_FIELD', '')).strip().upper()
            desc = str(row.get('BUSINESS_DESC', '')).strip().upper()
            rtype = str(row.get('RULE_TYPE', '')).strip()
            
            # 1. Search Filter (Case Insensitive Match)
            # Logic: If search exists, it MUST be in TARGET or DESCRIPTION
            if search:
                if (search not in tgt) and (search not in desc):
                    continue
            
            # 2. Scope Filter
            row_scope = str(row.get('SCOPE', 'GLOBAL'))
            if scope_filter != "ALL" and row_scope != scope_filter: 
                continue
            
            # 3. Type Filter
            if self.active_type_filters and rtype not in self.active_type_filters: 
                continue
            
            count += 1
            if count > 300: break

            # Formatting for display
            display_desc = str(row.get('BUSINESS_DESC', ''))
            desc_short = (display_desc[:23] + '..') if len(display_desc) > 25 else display_desc
            
            if rtype in ['CONST', 'PYTHON', 'FILTER']:
                val_txt = str(row.get('RULE_VALUE', ''))
                val = f"V:{val_txt[:10]}"
            else:
                val = f"S:{str(row.get('SOURCE_FIELD', ''))[:10]}"
            
            txt = f"{tgt:<10} {desc_short:<22} {rtype:<8} {val}"
            
            col = "white"
            if rtype == 'TODO': col = "#FF5555"
            elif rtype == 'IGNORE': col = "gray"
            elif rtype == 'MAP': col = "#44AAFF"
            elif rtype == 'CONST': col = "#FFAA00"
            elif rtype == 'PYTHON': col = "#AAFF00"
            elif rtype == 'FILTER': col = "#FF00FF"
            
            btn = ctk.CTkButton(self.scroll, text=txt, font=("Consolas", 12), fg_color="transparent", 
                                anchor="w", text_color=col, height=32, hover_color="#3A3A3A", 
                                command=lambda i=idx: self.on_select(i))
            btn.pack(fill="x", pady=1)

# --- COMPONENT 2: THE FORM PANEL (RIGHT) ---
class RuleFormPanel(ctk.CTkFrame):
    def __init__(self, parent, on_save_callback, on_override_callback):
        super().__init__(parent, corner_radius=0)
        self.on_save = on_save_callback
        self.on_override = on_override_callback
        self.curr_idx = None
        self.scopes_list = ["GLOBAL"]
        
        ctk.CTkLabel(self, text="EDIT RULE", font=("Arial", 16, "bold")).pack(anchor="w", padx=20, pady=(20, 10))
        
        # Fields
        self.lbl_tgt = ctk.CTkLabel(self, text="-", font=("Arial", 24, "bold"), text_color="#3B8ED0")
        self.lbl_tgt.pack(anchor="w", padx=20)
        self.lbl_meta = ctk.CTkLabel(self, text="", text_color="orange", font=("Arial", 11))
        self.lbl_meta.pack(anchor="w", padx=20)
        self.lbl_desc = ctk.CTkLabel(self, text="", text_color="gray", wraplength=400, justify="left")
        self.lbl_desc.pack(anchor="w", padx=20, pady=(0, 15))

        # Scope
        self._add_lbl("Context (Scope):")
        scope_box = ctk.CTkFrame(self, fg_color="transparent"); scope_box.pack(fill="x", padx=20)
        self.ent_scope = ctk.CTkEntry(scope_box)
        self.ent_scope.pack(side="left", fill="x", expand=True)
        
        btn_sel = ctk.CTkButton(scope_box, text="Select...", width=60, command=self._open_scope)
        btn_sel.pack(side="right", padx=(5,0))
        btn_over = ctk.CTkButton(scope_box, text="+ Override", width=80, fg_color="#555", command=self._click_override)
        btn_over.pack(side="right", padx=(5,0))

        # Type
        self._add_lbl("Rule Type:")
        self.opt_type = ctk.CTkOptionMenu(self, values=["DIRECT", "CONST", "MAP", "PYTHON", "FILTER", "IGNORE", "TODO"], command=self._update_state)
        self.opt_type.pack(fill="x", padx=20)

        # Source
        self.lbl_src_title = self._add_lbl("Source Field:")
        self.src_var = ctk.StringVar()
        self.src_var.trace("w", self._auto_rename_target) 
        self.ent_src = ctk.CTkEntry(self, textvariable=self.src_var)
        self.ent_src.pack(fill="x", padx=20)

        # Value
        self.lbl_val_title = self._add_lbl("Value / Logic:")
        val_box = ctk.CTkFrame(self, fg_color="transparent"); val_box.pack(fill="both", expand=True, padx=20)
        self.txt_val = ctk.CTkTextbox(val_box, height=120, font=("Consolas", 12))
        self.txt_val.pack(side="left", fill="both", expand=True)

        self.btn_map = ctk.CTkButton(val_box, text="...", width=30, command=self._run_map_wizard)
        
        btn_save = ctk.CTkButton(self, text="SAVE CHANGE", fg_color="green", height=40, command=self._save_click)
        btn_save.pack(fill="x", padx=20, pady=20)

    def _add_lbl(self, txt):
        l = ctk.CTkLabel(self, text=txt)
        l.pack(anchor="w", padx=20, pady=(10,0))
        return l

    def _auto_rename_target(self, *args):
        """
        SAFE UPDATE: Only updates Target Name if it's explicitly a placeholder 
        or already a filter-named rule. PREVENTS OVERWRITING 'ITNO'.
        """
        if self.opt_type.get() == "FILTER":
            current_target = self.lbl_tgt.cget("text")
            src = self.src_var.get().strip().upper()
            
            # SAFEGUARD: Only rename if it's a temp row or already a filter row
            is_placeholder = (current_target == "_ROW_") or (current_target == "NEW_FILTER")
            is_already_filter = current_target.startswith("FILTER_")
            
            if is_placeholder or is_already_filter:
                if src:
                    self.lbl_tgt.configure(text=f"FILTER_{src}")
                else:
                    self.lbl_tgt.configure(text="FILTER_???")

    def set_data(self, idx, row, scopes):
        self.curr_idx = idx
        self.scopes_list = scopes
        
        tgt = str(row['TARGET_FIELD'])
        self.lbl_tgt.configure(text=tgt)
        self.lbl_desc.configure(text=str(row.get('BUSINESS_DESC', '')))
        
        # Meta info safely
        m_type = row.get('M3_TYPE', '')
        m_len = row.get('M3_LENGTH', '')
        m_dec = row.get('M3_DECIMALS', '')
        self.lbl_meta.configure(text=f"Type: {m_type} | Len: {m_len} | Dec: {m_dec}")
        
        self.ent_scope.delete(0,"end"); self.ent_scope.insert(0, str(row.get('SCOPE', 'GLOBAL')))
        self.opt_type.set(row.get('RULE_TYPE', 'TODO'))
        
        # Populate widgets
        self.ent_src.configure(state="normal")
        self.ent_src.delete(0, "end")
        self.ent_src.insert(0, str(row.get('SOURCE_FIELD', '')))
        
        self.txt_val.configure(state="normal")
        self.txt_val.delete("0.0", "end")
        self.txt_val.insert("0.0", str(row.get('RULE_VALUE', '')))
        
        self._update_state(row.get('RULE_TYPE', 'TODO'), insert_template=False)

    def _update_state(self, choice, insert_template=False):
        # Reset to enabled
        self.ent_src.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        self.txt_val.configure(state="normal", fg_color=["#F9F9FA", "#343638"])
        self.btn_map.pack_forget()

        if choice == 'CONST':
            self.lbl_src_title.configure(text="Source (Disabled):")
            self.ent_src.delete(0,"end") 
            self.ent_src.configure(state="disabled", fg_color=["#EBEBEB", "#2B2B2B"])
            self.lbl_val_title.configure(text="Constant Value:")
            
        elif choice == 'DIRECT':
            self.lbl_val_title.configure(text="Value (Disabled):")
            self.txt_val.delete("0.0","end")
            self.txt_val.configure(state="disabled", fg_color=["#EBEBEB", "#2B2B2B"])
            
        elif choice == 'PYTHON':
            self.lbl_src_title.configure(text="Source (Var 'source'):")
            self.lbl_val_title.configure(text="Python Code:")
            if insert_template and not self.txt_val.get("0.0", "end").strip():
                self.txt_val.insert("0.0", "# if source == '1': return 'A'\n# return source")
        
        elif choice == 'FILTER':
            self.lbl_src_title.configure(text="Field to Check (Safe Rename):")
            self.lbl_val_title.configure(text="Filter Logic (Returns True/False):")
            # Only trigger safe rename
            self._auto_rename_target()
            if insert_template and not self.txt_val.get("0.0", "end").strip():
                self.txt_val.insert("0.0", "val = str(row.get('MMITNO') or '').strip()\nif val.startswith('ZZ'): return False\nreturn True")
                
        elif choice == 'MAP':
            self.lbl_src_title.configure(text="Source (Key):")
            self.lbl_val_title.configure(text="Map Config:")
            self.btn_map.pack(side="right", padx=5, fill="y")
            
        elif choice in ['IGNORE', 'TODO']:
            self.ent_src.delete(0,"end"); self.ent_src.configure(state="disabled", fg_color=["#EBEBEB", "#2B2B2B"])
            self.txt_val.delete("0.0","end"); self.txt_val.configure(state="disabled", fg_color=["#EBEBEB", "#2B2B2B"])

    def _open_scope(self):
        def on_select(sel): 
            self.ent_scope.delete(0,"end")
            self.ent_scope.insert(0, ",".join(sel))
        options = self.scopes_list if self.scopes_list else ["GLOBAL"]
        MultiSelectPopup(self, options, self.ent_scope.get(), on_select)

    def _click_override(self):
        current = self.ent_scope.get().strip()
        options = [s for s in self.scopes_list if s != current]
        def on_select(sel):
            if not sel: return
            self.on_override(self.curr_idx, sel[0])
        MultiSelectPopup(self, options, "", on_select)

    def _run_map_wizard(self):
        f = select_file("Select Map", [("Data", "*.xlsx *.csv")])
        if f:
            try: rel = os.path.relpath(f); 
            except: rel = f
            self.txt_val.delete("0.0","end"); self.txt_val.insert("0.0", f"{rel}|KEY|VAL")

    def _save_click(self):
        if self.curr_idx is None: return
        final_target = self.lbl_tgt.cget("text")
        
        data = {
            'TARGET_FIELD': final_target,
            'RULE_TYPE': self.opt_type.get(),
            'SOURCE_FIELD': self.ent_src.get(),
            'RULE_VALUE': self.txt_val.get("0.0", "end").strip(),
            'SCOPE': self.ent_scope.get()
        }
        self.on_save(self.curr_idx, data)

# --- COMPONENT 3: MAIN CONTROLLER ---
class RulesEditor(ctk.CTkFrame):
    def __init__(self, parent, rule_manager):
        super().__init__(parent, fg_color="transparent")
        self.manager = rule_manager
        self.curr_df = None
        self.curr_api = None
        self.scopes = ["GLOBAL"]

        # Top Bar
        top = ctk.CTkFrame(self, fg_color="transparent", height=40)
        top.pack(fill="x", pady=5)
        
        self.ed_var = ctk.StringVar()
        self.ed_menu = ctk.CTkOptionMenu(top, variable=self.ed_var, command=self._load_api, width=250)
        self.ed_menu.pack(side="left")
        
        btn_refresh = ctk.CTkButton(top, text="Refresh", width=80, command=self._refresh)
        btn_refresh.pack(side="left", padx=10)
        
        btn_add_filter = ctk.CTkButton(top, text="+ Filter", width=80, fg_color="#AA44AA", command=self._add_filter_rule)
        btn_add_filter.pack(side="right", padx=10)
        bind_context_help(btn_add_filter, "Add a new Row Filter. It will be named 'FILTER_[Source]'.")

        # Splitter
        paned = tk.PanedWindow(self, orient="horizontal", bg="#2B2B2B", bd=0, sashwidth=10, sashrelief="raised")
        paned.pack(fill="both", expand=True)

        self.list_panel = RuleListPanel(paned, self._on_row_select)
        self.form_panel = RuleFormPanel(paned, self._on_save, self._on_override)
        
        paned.add(self.list_panel, width=450)
        paned.add(self.form_panel, stretch="always")
        
        self._refresh()

    def _refresh(self):
        new_scopes = ["GLOBAL"]
        path = "config/business_units.csv"
        
        if os.path.exists(path):
            try:
                df = pd.read_csv(path, dtype=str).fillna("")
                df.columns = [c.strip().upper() for c in df.columns]
                
                if 'UNIT' in df.columns:
                    raw_units = df['UNIT'].tolist()
                    clean_units = [u.strip() for u in raw_units if u.strip() and u.strip().upper() != 'GLOBAL']
                    
                    seen = set(new_scopes)
                    for u in clean_units:
                        if u not in seen:
                            new_scopes.append(u)
                            seen.add(u)
            except Exception as e:
                print(f"{Fore.RED}ERROR loading business units: {e}{Style.RESET_ALL}")
            
        self.scopes = new_scopes 
        
        names = self.manager.get_available_rules()
        if names:
            self.ed_menu.configure(values=names)
            if not self.ed_var.get() or self.ed_var.get() == "No Rules":
                self.ed_var.set(names[0])
            self.after(100, lambda: self._load_api(self.ed_var.get()))
        else:
            self.ed_menu.configure(values=["No Rules"])
            self.ed_var.set("No Rules")

    def _load_api(self, api_name):
        if api_name == "No Rules": return
        self.curr_api = api_name
        self.curr_df, _ = self.manager.load_rules(api_name)
        if self.curr_df is not None:
            self.list_panel.load_data(self.curr_df, self.scopes)

    def _on_row_select(self, idx):
        if self.curr_df is not None:
            self.form_panel.set_data(idx, self.curr_df.iloc[idx], self.scopes)

    def _on_override(self, idx, new_scope):
        if idx is None: return
        success, msg, new_idx = self.manager.create_scope_override(self.curr_api, idx, new_scope)
        if success:
            self._load_api(self.curr_api)
            self.list_panel.scroll.update() 
            self._on_row_select(new_idx)
            messagebox.showinfo("Success", f"Created override for {new_scope}.")
        else:
            messagebox.showerror("Error", msg)

    def _on_save(self, idx, data):
        success, msg = self.manager.save_rule_update(self.curr_api, idx, data)
        if success:
            self._load_api(self.curr_api)
            self._on_row_select(idx)
        else:
            messagebox.showerror("Error", msg)

    def _add_filter_rule(self):
        if not self.curr_api: return
        new_row = {
            'TARGET_FIELD': 'NEW_FILTER',
            'RULE_TYPE': 'FILTER',
            'SOURCE_FIELD': '',
            'RULE_VALUE': "val = str(row.get('COL') or '').strip()\nif val == 'X': return True\nreturn False",
            'SCOPE': 'GLOBAL',
            'BUSINESS_DESC': 'Custom Row Filter'
        }
        self.curr_df = pd.concat([self.curr_df, pd.DataFrame([new_row])], ignore_index=True)
        new_idx = len(self.curr_df) - 1
        success, msg = self.manager.save_rule_update(self.curr_api, new_idx, new_row)
        
        if success:
            self._load_api(self.curr_api)
            self._on_row_select(new_idx)
            messagebox.showinfo("Success", "New Filter created. Type the Source Field to rename it.")
        else:
            messagebox.showerror("Error", msg)