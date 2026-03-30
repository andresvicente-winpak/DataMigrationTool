import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import glob
import os
import pandas as pd
import configparser
from modules.mco_importer import MCOImporter
from modules.audit_manager import AuditManager
from modules.extractor import DataExtractor
from modules.validator_analyzer import ValidatorAnalyzer
from modules.rule_manager import RuleManager
from modules.mco_checker import MCOChecker
from modules.ui import select_file 
from modules.gui.utils import bind_context_help

class ConfigHub(ctk.CTkTabview):
    def __init__(self, parent):
        super().__init__(parent)
        self.add("Import MCO")
        self.add("Reverse Engineer")
        self.add("Map Editor")
        
        try:
            btns = self._segmented_button._buttons_dict
            if "Import MCO" in btns: bind_context_help(btns["Import MCO"], "Import Specification: Convert an Excel MCO file into a system Rule Config.")
            if "Reverse Engineer" in btns: bind_context_help(btns["Reverse Engineer"], "Incremental Analysis: Use AI to guess rules by comparing Legacy vs M3 data.")
            if "Map Editor" in btns: bind_context_help(btns["Map Editor"], "Map Editor: Modify the internal CSV configuration files directly.")
        except: pass

        self._build_imp(self.tab("Import MCO"))
        self._build_an(self.tab("Reverse Engineer"))
        self._build_editor(self.tab("Map Editor"))

        self.selected_merge_sheets = [] 

    def _browse(self, entry, initial_folder=None):
        base = os.getcwd()
        start_dir = base
        if initial_folder:
            potential_path = os.path.abspath(os.path.join(base, initial_folder))
            if os.path.exists(potential_path): start_dir = potential_path
        f = filedialog.askopenfilename(initialdir=start_dir, title="Select File", filetypes=[("All Files", "*.*")])
        if f: 
            entry.delete(0, "end"); entry.insert(0, f)
            entry.event_generate("<FocusOut>")

    def _read_csv_flexible(self, path):
        encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
        last_err = None
        for enc in encodings:
            try:
                return pd.read_csv(path, encoding=enc).fillna("")
            except Exception as e:
                last_err = e
        raise last_err

    # --- TAB 1: IMPORT ---
    def _build_imp(self, frame):
        ctk.CTkLabel(frame, text="Import Specification", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        
        self.imp_mco = ctk.CTkEntry(frame, placeholder_text="MCO File Path")
        self.imp_mco.pack(fill="x", pady=5)
        bind_context_help(self.imp_mco, "Path to the Excel MCO (Mapping) file.")
        
        btn_browse = ctk.CTkButton(frame, text="Browse", command=lambda: self._browse(self.imp_mco, "config/mco_specs"))
        btn_browse.pack(fill="x", pady=5)
        bind_context_help(btn_browse, "Select the MCO file from the 'config/mco_specs' folder.")
        
        ctk.CTkLabel(frame, text="Select Sheet:").pack(anchor="w")
        self.imp_sheet_var = ctk.StringVar(value="Load File First")
        self.imp_sheet_menu = ctk.CTkOptionMenu(frame, variable=self.imp_sheet_var)
        self.imp_sheet_menu.pack(fill="x", pady=5)
        bind_context_help(self.imp_sheet_menu, "Select the specific Excel sheet containing the mapping logic.")
        self.imp_mco.bind("<FocusOut>", self._load_sheets_trigger)
        
        self.imp_api = ctk.CTkEntry(frame, placeholder_text="Target API Name (e.g. MMS200MI)")
        self.imp_api.pack(fill="x", pady=10)
        bind_context_help(self.imp_api, "Name of the Rule Config to create (e.g. MMS001MI).")
        
        self.imp_overwrite = ctk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(frame, text="Force Overwrite (Discard manual rule edits)", variable=self.imp_overwrite)
        chk.pack(anchor="w", pady=5)
        bind_context_help(chk, "If checked, any manual changes made in the Rules Editor will be lost.\nIf unchecked, it merges new MCO fields while keeping your edits.")

        btn_imp = ctk.CTkButton(frame, text="IMPORT RULES", fg_color="green", height=40, command=self._run_imp)
        btn_imp.pack(fill="x", pady=10)
        bind_context_help(btn_imp, "Parse the MCO file and generate/update the Rule Configuration Excel.")
        
        btn_val = ctk.CTkButton(frame, text="VALIDATE MCO HEALTH", fg_color="#AA0000", command=self._run_check)
        btn_val.pack(fill="x", pady=5)
        bind_context_help(btn_val, "Scan the MCO file for duplicate rules or missing required logic.")

    def _load_sheets_trigger(self, event=None):
        path = self.imp_mco.get()
        if os.path.exists(path):
            sheets = MCOImporter().get_sheet_names(path)
            self.imp_sheet_menu.configure(values=sheets)
            if sheets: self.imp_sheet_var.set(sheets[0])

    def _run_check(self):
        path = self.imp_mco.get()
        if not path or not os.path.exists(path): return
        def _t(): MCOChecker().check_file(path)
        threading.Thread(target=_t).start()

    def _run_imp(self):
        def _thread():
            path = self.imp_mco.get(); sheet = self.imp_sheet_var.get(); api = self.imp_api.get()
            if not api: api = sheet.split(' ')[0] + "MI"
            MCOImporter().run_import_headless(path, sheet, api, overwrite_all=self.imp_overwrite.get())
            AuditManager().commit_changes(api)
        threading.Thread(target=_thread).start()

    # --- TAB 2: ANALYZE ---
    def _build_an(self, frame):
        ctk.CTkLabel(frame, text="Incremental Analysis", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        self.an_leg = ctk.CTkEntry(frame, placeholder_text="Legacy File"); self.an_leg.pack(fill="x", pady=5)
        
        btn_brow_leg = ctk.CTkButton(frame, text="Browse Legacy", command=lambda: self._browse(self.an_leg, "raw_data"))
        btn_brow_leg.pack(anchor="e")
        bind_context_help(btn_brow_leg, "Select the source Movex/Legacy file.")

        self.an_leg_sheet = ctk.CTkOptionMenu(frame, values=["Load File First"]); self.an_leg_sheet.pack(fill="x", pady=5)
        
        self.an_gold = ctk.CTkEntry(frame, placeholder_text="Gold M3 File"); self.an_gold.pack(fill="x", pady=5)
        
        btn_brow_gold = ctk.CTkButton(frame, text="Browse Gold", command=lambda: self._browse(self.an_gold, "config/sdt_templates"))
        btn_brow_gold.pack(anchor="e")
        bind_context_help(btn_brow_gold, "Select the target M3 SDT (Gold Standard) file.")

        self.an_gold_sheet = ctk.CTkOptionMenu(frame, values=["Load File First"]); self.an_gold_sheet.pack(fill="x", pady=5)

        self.an_target_var = ctk.StringVar(value="Target Rule Config")
        self.an_target_menu = ctk.CTkOptionMenu(frame, variable=self.an_target_var)
        self.an_target_menu.pack(fill="x", pady=10)
        bind_context_help(self.an_target_menu, "Select which Rule Config will receive the suggested rules.")
        self._refresh_an_configs()
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent"); btn_frame.pack(fill="x", pady=20)
        
        btn_an = ctk.CTkButton(btn_frame, text="ANALYZE & MERGE", fg_color="orange", text_color="black", height=40, command=self._run_an)
        btn_an.pack(side="left", fill="x", expand=True, padx=5)
        bind_context_help(btn_an, "Compare files, detect patterns (CONST/DIRECT/LOGIC), and update the rules.")

        btn_dr = ctk.CTkButton(btn_frame, text="MERGE DRAFT", fg_color="#444", height=40, command=self._run_merge_draft)
        btn_dr.pack(side="right", fill="x", expand=True, padx=5)
        bind_context_help(btn_dr, "Import a previously generated Draft Excel into the Rule Config.")

        self.an_leg.bind("<FocusOut>", lambda e: self._load_sheets(self.an_leg.get(), self.an_leg_sheet))
        self.an_gold.bind("<FocusOut>", lambda e: self._load_sheets(self.an_gold.get(), self.an_gold_sheet))

    def _load_sheets(self, path, menu_widget):
        if os.path.exists(path):
            sheets = MCOImporter().get_sheet_names(path)
            menu_widget.configure(values=sheets)
            if sheets: menu_widget.set(sheets[0])

    def _refresh_an_configs(self):
        files = glob.glob('config/rules/*.xlsx')
        names = [os.path.basename(f).replace('.xlsx','') for f in files] or ["No Rules Found"]
        if names: self.an_target_menu.configure(values=names); self.an_target_var.set(names[0])

    def _run_an(self):
        pass # Placeholder for logic

    def _run_merge_draft(self):
        draft = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not draft: return
        target = self.an_target_var.get()
        def _t(): RuleManager().merge_draft_file(draft, target)
        threading.Thread(target=_t).start()

    # --- TAB 3: MAP EDITOR ---
    def _build_editor(self, frame):
        top = ctk.CTkFrame(frame, fg_color="transparent"); top.pack(fill="x", pady=5)
        
        self.map_files = {
            "Objects API Map (Merged)": "objects_api.csv",
            "Surgical Def (Objects)": "surgical_def.csv",
            "Business Units (Scopes)": "business_units.csv"
        }
        
        self.map_var = ctk.StringVar(value=list(self.map_files.keys())[0])
        self.map_menu = ctk.CTkOptionMenu(top, variable=self.map_var, values=list(self.map_files.keys()), command=self._load_map)
        self.map_menu.pack(side="left", fill="x", expand=True)
        bind_context_help(self.map_menu, "Select which system configuration file to edit.")
        
        # --- DB BUTTON ---
        btn_db = ctk.CTkButton(top, text="DB Connection", width=100, fg_color="#555", command=self._open_db_config)
        btn_db.pack(side="left", padx=10)
        bind_context_help(btn_db, "Configure SQL Server connection settings for 'SQL:...' data sources.")

        btn_save = ctk.CTkButton(top, text="Save", width=80, fg_color="green", command=self._save_map)
        btn_save.pack(side="right", padx=5)
        bind_context_help(btn_save, "Save changes to the CSV file.")

        btn_reload = ctk.CTkButton(top, text="Reload", width=80, command=lambda: self._load_map(None))
        btn_reload.pack(side="right")
        bind_context_help(btn_reload, "Discard unsaved changes and reload from disk.")

        self.grid_frame = ctk.CTkScrollableFrame(frame, label_text="Config Data")
        self.grid_frame.pack(fill="both", expand=True, pady=10)
        
        btn_add = ctk.CTkButton(frame, text="+ Add Row", command=self._add_row)
        btn_add.pack(fill="x", pady=10)
        bind_context_help(btn_add, "Append a new empty row to the configuration.")

        self.cells = []
        self.headers = []
        self._load_map(None)

    def _open_db_config(self):
        """ Opens a popup to edit config/db_config.ini """
        popup = ctk.CTkToplevel(self)
        popup.title("Database Configuration")
        popup.geometry("400x400")
        popup.grab_set() # Modal
        
        config_path = os.path.join(os.getcwd(), 'config', 'db_config.ini')
        config = configparser.ConfigParser()
        
        fields = ['Driver', 'Server', 'Database', 'Trusted_Connection']
        tooltips = {
            'Driver': "ODBC Driver name (e.g. 'ODBC Driver 17 for SQL Server').",
            'Server': "Network name or IP of the SQL Server instance.",
            'Database': "Name of the source database.",
            'Trusted_Connection': "'yes' for Windows Auth, 'no' for SQL Auth."
        }
        
        current_values = {
            'Driver': 'ODBC Driver 17 for SQL Server', 
            'Server': 'LOCALHOST', 
            'Database': 'M3_DATA', 
            'Trusted_Connection': 'yes'
        }

        if os.path.exists(config_path):
            try:
                config.read(config_path)
                if 'DEFAULT' in config:
                    loaded_conf = config['DEFAULT']
                    for field in fields:
                        val = loaded_conf.get(field, loaded_conf.get(field.lower()))
                        if val is not None:
                            current_values[field] = val
            except Exception as e:
                print(f"Error loading db_config: {e}")
        
        ctk.CTkLabel(popup, text="SQL Server Connection", font=("Arial", 16, "bold")).pack(pady=10)
        
        entries = {}
        for key in fields:
            ctk.CTkLabel(popup, text=key + ":").pack(anchor="w", padx=20)
            ent = ctk.CTkEntry(popup)
            ent.insert(0, str(current_values.get(key, '')))
            ent.pack(fill="x", padx=20, pady=(0, 5))
            bind_context_help(ent, tooltips[key])
            entries[key] = ent
        
        def save():
            config['DEFAULT'] = {k: v.get() for k, v in entries.items()}
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w') as f: config.write(f)
            messagebox.showinfo("Saved", f"Database configuration saved to:\n{config_path}")
            popup.destroy()

        def test_popup_conn():
            try:
                config['DEFAULT'] = {k: v.get() for k, v in entries.items()}
                os.makedirs(os.path.dirname(config_path), exist_ok=True)
                with open(config_path, 'w') as f: config.write(f)
                ex = DataExtractor()
                eng = ex._get_sql_connection()
                with eng.connect() as conn: pass 
                messagebox.showinfo("Success", "Connection Successful!")
            except Exception as e:
                messagebox.showerror("Failed", f"Connection failed:\n{e}")
            
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=20)
        btn_test = ctk.CTkButton(btn_frame, text="Test Connection", fg_color="#0055AA", width=120, command=test_popup_conn)
        btn_test.pack(side="left")
        btn_save = ctk.CTkButton(btn_frame, text="Save & Close", fg_color="green", width=120, command=save)
        btn_save.pack(side="right")

    def _load_map(self, _):
        for widget in self.grid_frame.winfo_children(): widget.destroy()
        self.cells = []
        self.headers = []

        selection_key = self.map_var.get()
        filename = self.map_files[selection_key]
        path = os.path.join('config', filename)
        
        is_sql_map = "Objects API Map" in selection_key

        if not os.path.exists(path):
            if "business_units" in filename: 
                pd.DataFrame({'UNIT': ['GLOBAL'], 'DESCRIPTION': ['Global Rules']}).to_csv(path, index=False)
            elif "surgical_def" in filename:
                pd.DataFrame(columns=['OBJECT_TYPE', 'MCO_SHEET', 'KEY_COLUMN']).to_csv(path, index=False)
            else: 
                pd.DataFrame().to_csv(path, index=False)
        
        if "surgical_def" in filename:
            try:
                df_check = self._read_csv_flexible(path)
                if 'KEY_COLUMN' not in df_check.columns:
                    df_check['KEY_COLUMN'] = ''
                    df_check.to_csv(path, index=False)
            except: pass

        try:
            df = self._read_csv_flexible(path)
            self.headers = list(df.columns)
            
            for i, h in enumerate(self.headers):
                ctk.CTkLabel(self.grid_frame, text=h, font=("Arial", 12, "bold")).grid(row=0, column=i, padx=5, pady=5, sticky="w")
            
            # Additional column headers for Actions
            action_col_start = len(self.headers)
            if is_sql_map and ("SOURCE_FILE" in [h.upper().strip() for h in self.headers] or "SOURCE" in [h.upper().strip() for h in self.headers]):
                ctk.CTkLabel(self.grid_frame, text="ACTIONS", font=("Arial", 12, "bold"), text_color="cyan").grid(row=0, column=action_col_start, padx=5, sticky="w")

            for r_idx, row in df.iterrows():
                row_widgets = []
                for c_idx, col in enumerate(self.headers):
                    ent = ctk.CTkEntry(self.grid_frame)
                    ent.insert(0, str(row[col]))
                    ent.grid(row=r_idx+1, column=c_idx, padx=2, pady=2, sticky="ew")
                    row_widgets.append(ent)
                
                btn_col = action_col_start
                if is_sql_map and ("SOURCE_FILE" in [h.upper().strip() for h in self.headers] or "SOURCE" in [h.upper().strip() for h in self.headers]):
                    # NEW: Edit Button for Long SQL
                    btn_edit = ctk.CTkButton(self.grid_frame, text="Edit", width=40, fg_color="#555", command=lambda r=r_idx: self._edit_source_sql(r))
                    btn_edit.grid(row=r_idx+1, column=btn_col, padx=2)
                    bind_context_help(btn_edit, "Open larger editor for SQL queries.")
                    btn_col += 1
                    
                    btn_test = ctk.CTkButton(self.grid_frame, text="Test", width=40, fg_color="#0055AA", command=lambda r=r_idx: self._test_sql_row(r))
                    btn_test.grid(row=r_idx+1, column=btn_col, padx=2)
                    bind_context_help(btn_test, "Test connection and run query (Only if line starts with SQL:)")
                    btn_col += 1

                btn_del = ctk.CTkButton(self.grid_frame, text="X", width=30, fg_color="#550000", command=lambda r=r_idx: self._delete_row(r))
                btn_del.grid(row=r_idx+1, column=btn_col, padx=2)
                
                self.cells.append(row_widgets)
                
        except Exception as e: print(f"Error loading map: {e}")

    def _add_row(self):
        r_idx = len(self.cells) + 1
        row_widgets = []
        
        selection_key = self.map_var.get()
        is_sql_map = "Objects API Map" in selection_key
        has_source_col = "SOURCE_FILE" in [h.upper().strip() for h in self.headers] or "SOURCE" in [h.upper().strip() for h in self.headers]
        
        for c_idx, _ in enumerate(self.headers):
            ent = ctk.CTkEntry(self.grid_frame)
            ent.grid(row=r_idx, column=c_idx, padx=2, pady=2, sticky="ew")
            row_widgets.append(ent)
            
        btn_col = len(self.headers)
        if is_sql_map and has_source_col:
            curr_idx = len(self.cells) 
            # NEW: Edit Button
            btn_edit = ctk.CTkButton(self.grid_frame, text="Edit", width=40, fg_color="#555", command=lambda r=curr_idx: self._edit_source_sql(r))
            btn_edit.grid(row=r_idx, column=btn_col, padx=2)
            btn_col += 1
            
            btn_test = ctk.CTkButton(self.grid_frame, text="Test", width=40, fg_color="#0055AA", command=lambda r=curr_idx: self._test_sql_row(r))
            btn_test.grid(row=r_idx, column=btn_col, padx=2)
            btn_col += 1

        # Delete Button (for new row, we can't easily delete until reload, but UI symmetry)
        ctk.CTkLabel(self.grid_frame, text="(Save)", font=("Arial", 10)).grid(row=r_idx, column=btn_col, padx=2)

        self.cells.append(row_widgets)

    def _edit_source_sql(self, row_idx):
        """ Opens a large multi-line text editor for the SOURCE_FILE column """
        if row_idx >= len(self.cells): return
        
        # Find SOURCE_FILE column index
        try:
            col_map = {h.upper().strip(): i for i, h in enumerate(self.headers)}
            if 'SOURCE_FILE' in col_map: col_idx = col_map['SOURCE_FILE']
            elif 'SOURCE' in col_map: col_idx = col_map['SOURCE']
            else: return # No source file column
        except: return

        entry_widget = self.cells[row_idx][col_idx]
        current_val = entry_widget.get()

        popup = ctk.CTkToplevel(self)
        popup.title("SQL Query Editor")
        popup.geometry("600x400")
        popup.grab_set()

        ctk.CTkLabel(popup, text="Source Query Editor", font=("Arial", 14, "bold")).pack(pady=10)
        
        txt = ctk.CTkTextbox(popup, wrap="word")
        txt.pack(fill="both", expand=True, padx=20, pady=5)
        txt.insert("0.0", current_val)
        
        def save_edit():
            # Get text and flatten to single line for CSV safety
            raw = txt.get("0.0", "end").strip()
            # Replace newlines with spaces so it stays on one CSV line
            flat = raw.replace("\n", " ").replace("\r", " ")
            while "  " in flat: flat = flat.replace("  ", " ") # Clean multi-spaces
            
            entry_widget.delete(0, "end")
            entry_widget.insert(0, flat)
            popup.destroy()
        
        ctk.CTkButton(popup, text="Update Row", fg_color="green", command=save_edit).pack(pady=10)

    def _test_sql_row(self, row_idx):
        if row_idx >= len(self.cells): return
        widgets = self.cells[row_idx]
        if not widgets: return
        
        try:
            col_map = {h.upper().strip(): i for i, h in enumerate(self.headers)}
            if 'SOURCE_FILE' in col_map: col_idx = col_map['SOURCE_FILE']
            elif 'SOURCE' in col_map: col_idx = col_map['SOURCE']
            else: raise ValueError("SOURCE/SOURCE_FILE column not found")
        except Exception: col_idx = 0
            
        val = widgets[col_idx].get().strip()
        
        if not val.upper().startswith("SQL:"):
            messagebox.showinfo("Ignored", f"This row is not a SQL connection.\nString must start with 'SQL:'")
            return
            
        try:
            df = DataExtractor().load_data(val)
            cols = list(df.columns)
            preview = df.head(3).to_string()
            messagebox.showinfo("Connection Successful", f"Query returned {len(df)} rows.\n\nColumns: {cols}\n\nPreview:\n{preview}")
        except Exception as e:
            messagebox.showerror("Connection Failed", f"Error:\n{e}")

    def _save_map(self):
        filename = self.map_files[self.map_var.get()]
        path = os.path.join('config', filename)

        issues = []  
        for r_idx, row_widgets in enumerate(self.cells, start=1):
            if not row_widgets or not row_widgets[0].winfo_exists(): continue

            for c_idx, w in enumerate(row_widgets):
                val = w.get()
                if isinstance(val, str) and val != val.rstrip():
                    col_name = self.headers[c_idx] if c_idx < len(self.headers) else f"COL_{c_idx}"
                    issues.append((r_idx, col_name, val))
                    try: w.configure(border_color="red")
                    except: pass
                else:
                    try: w.configure(border_color=None)
                    except: pass

        if issues:
            preview_lines = []
            for (r, col, v) in issues[:10]:
                preview_lines.append(f"Row {r}, Column '{col}': '{v}'")
            
            messagebox.showwarning(
                "Trailing spaces detected",
                "Some cells end with trailing spaces. Please remove them:\n\n" + "\n".join(preview_lines)
            )
            return

        data = []
        for row_widgets in self.cells:
            if row_widgets and row_widgets[0].winfo_exists():
                data.append({self.headers[i]: w.get() for i, w in enumerate(row_widgets)})

        pd.DataFrame(data).to_csv(path, index=False)
        print(f"Saved {filename}")
        self._load_map(None)
    
    def _delete_row(self, row_idx):
        filename = self.map_files[self.map_var.get()]
        path = os.path.join('config', filename)
        try:
            df = self._read_csv_flexible(path)
            if row_idx < len(df):
                df = df.drop(row_idx)
                df.to_csv(path, index=False)
                self._load_map(None)
        except Exception as e:
            print(f"Delete failed: {e}")
