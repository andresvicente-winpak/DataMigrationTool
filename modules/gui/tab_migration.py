import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import glob
import threading
import pandas as pd
import datetime
from colorama import Fore, Style

# Logic Imports
from modules.migration_runner import MigrationRunner
from modules.auto_detector import AutoDetector
from modules.surgical_extractor import SurgicalExtractor
from modules.batch_processor import BatchProcessor
from modules.gui.utils import bind_context_help

class MigrationHub(ctk.CTkTabview):
    def __init__(self, parent):
        super().__init__(parent)
        self.add("Standard")
        self.add("Auto-Detect")
        self.add("Load by ID")
        self.add("Batch")
        
        try:
            btns = self._segmented_button._buttons_dict
            if "Standard" in btns: bind_context_help(btns["Standard"], "Classic migration: Select a Rule Config and a defined Source Map entry.")
            if "Auto-Detect" in btns: bind_context_help(btns["Auto-Detect"], "Magic Drop: Identify the API from the file header automatically.")
            if "Load by ID" in btns: bind_context_help(btns["Load by ID"], "Surgical Extraction: Load specific records based on a list of IDs.")
            if "Batch" in btns: bind_context_help(btns["Batch"], "Process multiple migration jobs sequentially via an Excel manifest.")
        except: pass
        
        self._build_standard(self.tab("Standard"))
        self._build_auto_detect(self.tab("Auto-Detect"))
        self._build_load_by_id(self.tab("Load by ID"))
        self._build_batch(self.tab("Batch"))

    def _browse(self, entry, initial_folder=None):
        start_dir = os.getcwd()
        if initial_folder:
            if os.path.isabs(initial_folder): potential_path = initial_folder
            else: potential_path = os.path.join(start_dir, initial_folder)
            if os.path.exists(potential_path): start_dir = potential_path

        f = filedialog.askopenfilename(initialdir=start_dir, filetypes=[("All Files", "*.*"), ("Excel", "*.xlsx")])
        if f: 
            entry.delete(0, "end"); entry.insert(0, f)

    def _refresh_configs(self, menu, var):
        files = glob.glob('config/rules/*.xlsx')
        names = [os.path.basename(f).replace('.xlsx','') for f in files] or ["No Rules Found"]
        menu.configure(values=names); var.set(names[0])

    def _refresh_sources(self, menu, var):
        """Loads MCO_SHEET list from source_map.csv"""
        try:
            if os.path.exists("config/source_map.csv"):
                df = self._read_csv_flexible("config/source_map.csv")
                if 'MCO_SHEET' in df.columns:
                    sheets = sorted(df['MCO_SHEET'].dropna().unique().tolist())
                    if sheets:
                        menu.configure(values=sheets)
                        var.set(sheets[0])
                        return
            
            menu.configure(values=["No Sources Found"])
            var.set("No Sources Found")
        except Exception as e:
            print(f"Error loading sources: {e}")
            menu.configure(values=["Error"])
            var.set("Error")

    def _read_csv_flexible(self, path):
        """
        Read CSVs that may come from Excel/Windows exports with non-UTF8 bytes.
        """
        encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
        last_err = None
        for enc in encodings:
            try:
                return pd.read_csv(path, encoding=enc).fillna("")
            except Exception as e:
                last_err = e
        raise last_err

    def _load_object_program_df(self):
        try:
            # Preferred source: merged objects_api.csv (OBJECTS + source/migration metadata)
            if os.path.exists("config/objects_api.csv"):
                df = self._read_csv_flexible("config/objects_api.csv")
            elif os.path.exists("config/source_map.csv"):
                # Backward compatibility fallback
                df = self._read_csv_flexible("config/source_map.csv")
            else:
                return pd.DataFrame()

            df.columns = [c.strip().upper() for c in df.columns]
            if "MCO_SHEET" not in df.columns:
                return pd.DataFrame()

            # If merged file has no OBJECTS, try enriching from source_map/legacy files
            if "OBJECTS" not in df.columns:
                if os.path.exists("config/objects_api.csv"):
                    obj_df = self._read_csv_flexible("config/objects_api.csv")
                    obj_df.columns = [c.strip().upper() for c in obj_df.columns]
                elif os.path.exists("config/source_map.csv"):
                    obj_df = self._read_csv_flexible("config/source_map.csv")
                    obj_df.columns = [c.strip().upper() for c in obj_df.columns]
                else:
                    obj_df = pd.DataFrame()

                if not obj_df.empty and {"OBJECTS", "MCO_SHEET"}.issubset(set(obj_df.columns)):
                    df = df.merge(
                        obj_df[["OBJECTS", "MCO_SHEET"]].drop_duplicates(),
                        on="MCO_SHEET",
                        how="left"
                    )

            required_cols = {"OBJECTS", "MCO_SHEET"}
            if not required_cols.issubset(set(df.columns)):
                return pd.DataFrame()

            df["OBJECT_NORM"] = df["OBJECTS"].astype(str).str.strip()
            df["MCO_NORM"] = df["MCO_SHEET"].astype(str).str.strip()
            if "API_NAME" in df.columns:
                df["API_NORM"] = df["API_NAME"].astype(str).str.strip()
            else:
                df["API_NORM"] = ""

            # Accept either SOURCE_FILE or SOURCE as the source/sql column
            if "SOURCE_FILE" in df.columns:
                df["SOURCE_NORM"] = df["SOURCE_FILE"].astype(str).str.strip()
            elif "SOURCE" in df.columns:
                df["SOURCE_NORM"] = df["SOURCE"].astype(str).str.strip()
            else:
                df["SOURCE_NORM"] = ""

            # If source is missing in merged file, fallback to source_map by MCO_SHEET
            if (df["SOURCE_NORM"] == "").any() and os.path.exists("config/source_map.csv"):
                src_df = self._read_csv_flexible("config/source_map.csv")
                src_df.columns = [c.strip().upper() for c in src_df.columns]
                if {"MCO_SHEET", "SOURCE_FILE"}.issubset(set(src_df.columns)):
                    src_df["MCO_KEY"] = src_df["MCO_SHEET"].astype(str).str.strip().str.upper()
                    src_df["SRC_KEY"] = src_df["SOURCE_FILE"].astype(str).str.strip()
                    source_map = src_df.drop_duplicates(subset=["MCO_KEY"]).set_index("MCO_KEY")["SRC_KEY"].to_dict()
                    df["MCO_KEY"] = df["MCO_SHEET"].astype(str).str.strip().str.upper()
                    df["SOURCE_NORM"] = df.apply(
                        lambda r: r["SOURCE_NORM"] if str(r["SOURCE_NORM"]).strip() else source_map.get(r["MCO_KEY"], ""),
                        axis=1
                    )
                    df = df.drop(columns=["MCO_KEY"], errors="ignore")

            df = df[(df["OBJECT_NORM"] != "") & (df["MCO_NORM"] != "")]
            return df
        except Exception as e:
            print(f"Error loading objects/programs: {e}")
            return pd.DataFrame()

    def _refresh_objects(self, menu, var):
        df = self._load_object_program_df()
        if df.empty:
            menu.configure(values=["No Objects Found"])
            var.set("No Objects Found")
            self._programs_by_object = {}
            return

        objects = sorted(df["OBJECT_NORM"].unique().tolist())
        self._programs_by_object = {
            obj: sorted(df[df["OBJECT_NORM"] == obj]["MCO_NORM"].unique().tolist())
            for obj in objects
        }
        menu.configure(values=objects)
        var.set(objects[0])

    def _on_object_change(self, selected_object):
        programs = self._programs_by_object.get(selected_object, [])
        if not programs:
            self.std_prog_menu.configure(values=["No Programs Found"])
            self.std_prog_var.set("No Programs Found")
            return
        self.std_prog_menu.configure(values=programs)
        self.std_prog_var.set(programs[0])

    def _load_scopes(self, menu):
        try:
            if os.path.exists("config/business_units.csv"):
                df = self._read_csv_flexible("config/business_units.csv")
                units = ["GLOBAL"] + df['UNIT'].dropna().unique().tolist()
                menu.configure(values=units)
        except: pass

    # --- TAB 1: STANDARD MIGRATION ---
    def _build_standard(self, frame):
        ctk.CTkLabel(frame, text="Standard Migration (Full Load)", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        
        # 1. Object Selection
        ctk.CTkLabel(frame, text="1. Select Object:").pack(anchor="w")
        self.std_obj_var = ctk.StringVar(value="Select Object")
        self.std_obj_menu = ctk.CTkOptionMenu(frame, variable=self.std_obj_var)
        self.std_obj_menu.pack(fill="x", pady=5)
        self._refresh_objects(self.std_obj_menu, self.std_obj_var)
        self.std_obj_menu.configure(command=self._on_object_change)
        bind_context_help(self.std_obj_menu, "Select the business object (from objects_api.csv).")

        # 2. Program Selection
        ctk.CTkLabel(frame, text="2. Select Program (MCO Sheet):").pack(anchor="w", pady=(10,0))
        
        prog_box = ctk.CTkFrame(frame, fg_color="transparent")
        prog_box.pack(fill="x")
        
        self.std_prog_var = ctk.StringVar(value="Select Program")
        self.std_prog_menu = ctk.CTkOptionMenu(prog_box, variable=self.std_prog_var)
        self.std_prog_menu.pack(side="left", fill="x", expand=True, padx=(0, 5), pady=5)
        self._on_object_change(self.std_obj_var.get())
        
        btn_ref_prog = ctk.CTkButton(prog_box, text="↻", width=30, command=lambda: (self._refresh_objects(self.std_obj_menu, self.std_obj_var), self._on_object_change(self.std_obj_var.get())))
        btn_ref_prog.pack(side="right", pady=5)
        
        bind_context_help(self.std_prog_menu, "Select the MCO Sheet/Program. Source and API are resolved automatically.")
        
        # 3. Scope
        ctk.CTkLabel(frame, text="3. Scope (Optional):").pack(anchor="w", pady=(10,0))
        self.std_scope_var = ctk.StringVar(value="GLOBAL")
        self.std_scope_menu = ctk.CTkOptionMenu(frame, variable=self.std_scope_var, values=["GLOBAL"])
        self.std_scope_menu.pack(fill="x", pady=5)
        self._load_scopes(self.std_scope_menu)
        bind_context_help(self.std_scope_menu, "Select the Business Unit/Division for rule overrides.")
        
        btn_run = ctk.CTkButton(frame, text="RUN MIGRATION", fg_color="green", height=40, command=self._run_std)
        btn_run.pack(fill="x", pady=20)
        bind_context_help(btn_run, "Starts the migration process.")

    def _run_std(self):
        selected_object = self.std_obj_var.get()
        sheet_name = self.std_prog_var.get()
        scope = self.std_scope_var.get()
        
        if (
            not selected_object or selected_object in ["Select Object", "No Objects Found"]
            or not sheet_name or sheet_name in ["Select Program", "No Programs Found"]
        ):
            messagebox.showwarning("Missing Input", "Please select a valid Object and Program.")
            return
        
        # Resolve program/API/source using merged objects_api map (or legacy fallbacks)
        conf = None
        source_path = None
        try:
            obj_df = self._load_object_program_df()
            if obj_df.empty:
                messagebox.showerror("Error", "No valid mapping data found in objects_api.csv / source_map.csv")
                return
            obj_df["OBJECT_LOOKUP"] = obj_df["OBJECT_NORM"].astype(str).str.upper()
            obj_df["MCO_LOOKUP"] = obj_df["MCO_NORM"].astype(str).str.upper()
            selected_object_up = selected_object.strip().upper()
            sheet_name_up = sheet_name.strip().upper()

            obj_match = obj_df[
                (obj_df["OBJECT_LOOKUP"] == selected_object_up) &
                (obj_df["MCO_LOOKUP"] == sheet_name_up)
            ]
            if obj_match.empty:
                messagebox.showerror("Error", f"Could not find Object/Program combination: {selected_object} / {sheet_name}")
                return

            selected_row = obj_match.iloc[0]
            conf = selected_row.get("API_NORM", "").strip()
            source_path = selected_row.get("SOURCE_NORM", "").strip()

            # API fallback (for legacy files where API_NAME is missing in merged objects map)
            if not conf and os.path.exists("config/objects_api.csv"):
                legacy_obj = self._read_csv_flexible("config/objects_api.csv")
                legacy_obj.columns = [c.strip().upper() for c in legacy_obj.columns]
                if {"OBJECTS", "MCO_SHEET", "API_NAME"}.issubset(set(legacy_obj.columns)):
                    match_legacy = legacy_obj[
                        (legacy_obj["OBJECTS"].astype(str).str.strip().str.upper() == selected_object_up) &
                        (legacy_obj["MCO_SHEET"].astype(str).str.strip().str.upper() == sheet_name_up)
                    ]
                    if not match_legacy.empty:
                        conf = str(match_legacy.iloc[0]["API_NAME"]).strip()

            if not conf:
                messagebox.showerror("Error", f"Could not resolve API_NAME for '{sheet_name}'.")
                return
            if not source_path:
                messagebox.showerror("Error", f"Could not resolve SOURCE_FILE/SQL for '{sheet_name}'.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resolve Object/Program mapping: {e}")
            return

        def _t():
            runner = MigrationRunner()
            runner.execute_migration(conf, source_path, division=scope, mco_context=sheet_name)
        threading.Thread(target=_t).start()
    
    # ... (Rest of the class: auto-detect, surgical, batch remain unchanged)
    # Re-paste the rest of your class below if needed, or just update the _run_std method
    # For completeness, here is the Auto Detect block to ensure file integrity
    
    def _build_auto_detect(self, frame):
        ctk.CTkLabel(frame, text="Auto-Detect Migration", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        ctk.CTkLabel(frame, text="Automatically identifies the API based on file headers.", text_color="gray").pack(anchor="w", pady=(0,10))
        ctk.CTkLabel(frame, text="1. Select Master MCO (to learn signatures):").pack(anchor="w")
        box1 = ctk.CTkFrame(frame, fg_color="transparent"); box1.pack(fill="x")
        self.auto_mco_ent = ctk.CTkEntry(box1, placeholder_text="Path to MCO.xlsx")
        self.auto_mco_ent.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(box1, text="Browse", width=80, command=lambda: self._browse(self.auto_mco_ent, "config/mco_specs")).pack(side="right")
        ctk.CTkLabel(frame, text="2. Select Legacy File (Unknown Content):").pack(anchor="w", pady=(10,0))
        box2 = ctk.CTkFrame(frame, fg_color="transparent"); box2.pack(fill="x")
        self.auto_leg_ent = ctk.CTkEntry(box2, placeholder_text="Path to Unknown.xlsx")
        self.auto_leg_ent.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(box2, text="Browse", width=80, command=lambda: self._browse(self.auto_leg_ent)).pack(side="right")
        ctk.CTkLabel(frame, text="3. Scope (Optional):").pack(anchor="w", pady=(10,0))
        self.auto_scope_var = ctk.StringVar(value="GLOBAL")
        self.auto_scope_menu = ctk.CTkOptionMenu(frame, variable=self.auto_scope_var, values=["GLOBAL"])
        self.auto_scope_menu.pack(fill="x", pady=5)
        self._load_scopes(self.auto_scope_menu)
        btn_auto = ctk.CTkButton(frame, text="DETECT & RUN", fg_color="orange", text_color="black", height=40, command=self._run_auto)
        btn_auto.pack(fill="x", pady=20)

    def _run_auto(self):
        mco = self.auto_mco_ent.get(); leg = self.auto_leg_ent.get(); scope = self.auto_scope_var.get()
        if not mco or not leg: return
        def _t():
            print(f"\n{Fore.CYAN}--- AUTO-DETECT STARTED ---{Style.RESET_ALL}")
            detector = AutoDetector(mco)
            prefix, sheet, api = detector.identify_file(leg)
            runner = MigrationRunner()
            if (not api or api == "Unknown") and sheet:
                print(f"{Fore.YELLOW}   API not found in header. Searching Map for '{sheet}'...{Style.RESET_ALL}")
                found_api, _, _ = runner.resolve_from_map_public(sheet, 'MCO_SHEET')
                if found_api:
                    api = found_api
                    print(f"{Fore.GREEN}   -> Found API in Map: {api}{Style.RESET_ALL}")
            if api and api != "Unknown":
                print(f"{Fore.GREEN}Detected: {prefix} -> {sheet} -> {api}{Style.RESET_ALL}")
                runner.execute_migration(api, leg, division=scope)
            else:
                print(f"{Fore.RED}MCO match not found-cannot proceed.{Style.RESET_ALL}")
        threading.Thread(target=_t).start()

    def _build_load_by_id(self, frame):
        ctk.CTkLabel(frame, text="Surgical Extraction (Load by ID)", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        ctk.CTkLabel(frame, text="1. Select Business Object:").pack(anchor="w")
        self.surg_obj_var = ctk.StringVar(value="Select Object")
        self.surg_obj_menu = ctk.CTkOptionMenu(frame, variable=self.surg_obj_var, values=SurgicalExtractor().get_available_objects())
        self.surg_obj_menu.pack(fill="x", pady=5)
        ctk.CTkLabel(frame, text="2. Scope Override (Optional):").pack(anchor="w", pady=(10,0))
        self.surg_scope = ctk.CTkEntry(frame, placeholder_text="e.g. DIV_US")
        self.surg_scope.pack(fill="x", pady=5)
        ctk.CTkLabel(frame, text="3. Enter IDs (Comma Separated):").pack(anchor="w", pady=(10,0))
        self.surg_ids = ctk.CTkTextbox(frame, height=100)
        self.surg_ids.pack(fill="x", pady=5)
        btn_surg = ctk.CTkButton(frame, text="RUN DELTA LOAD", fg_color="purple", height=40, command=self._run_surg)
        btn_surg.pack(fill="x", pady=20)

    def _run_surg(self):
        obj = self.surg_obj_var.get()
        raw_ids = self.surg_ids.get("0.0", "end").strip()
        if not raw_ids or obj == "Select Object": return
        id_list = [x.strip() for x in raw_ids.split(',')]
        def _t():
            extractor = SurgicalExtractor()
            runner = MigrationRunner()
            tasks = extractor.perform_extraction(obj, id_list)
            if not tasks:
                print("No extraction tasks generated.")
                return
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            master_name = f"SURGICAL_LOAD_{obj}_{timestamp}.xlsx"
            print(f"{Fore.CYAN}   -> Target Master File: {master_name}{Style.RESET_ALL}")
            first_valid_run = True
            for t in tasks:
                map_file = 'config/objects_api.csv' if os.path.exists('config/objects_api.csv') else 'config/migration_map.csv'
                map_df = self._read_csv_flexible(map_file)
                map_df.columns = [c.upper().strip() for c in map_df.columns]
                match = map_df[map_df['MCO_SHEET'].str.upper() == str(t['mco_sheet']).upper()]
                if match.empty:
                    print(f"{Fore.RED}    No mapping entry for {t['mco_sheet']} in {map_file}. Skipping.{Style.RESET_ALL}")
                    continue
                trans = str(match.iloc[0]['TRANSACTION_SHEET'])
                targets = [x.strip() for x in trans.split(',') if x.strip()]
                if not targets:
                    print(f"{Fore.YELLOW}    Warning: No TRANSACTION_SHEET defined for {t['mco_sheet']}. Skipping.{Style.RESET_ALL}")
                    continue
                should_append = not first_valid_run
                runner.execute_migration(
                    t['program_name'], 
                    t['legacy_path'], 
                    division=self.surg_scope.get().upper() or "GLOBAL", 
                    target_sheets=targets, 
                    silent=True, 
                    output_name_override=master_name,
                    append_if_exists=should_append
                )
                first_valid_run = False
        threading.Thread(target=_t).start()

    def _build_batch(self, frame):
        ctk.CTkLabel(frame, text="Batch Processor", font=("Arial", 18, "bold")).pack(anchor="w", pady=10)
        box = ctk.CTkFrame(frame, fg_color="transparent"); box.pack(fill="x")
        self.batch_ent = ctk.CTkEntry(box, placeholder_text="Batch Definition File"); self.batch_ent.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(box, text="Browse", width=80, command=lambda: self._browse(self.batch_ent)).pack(side="right")
        btn_batch = ctk.CTkButton(frame, text="EXECUTE BATCH", fg_color="#880000", height=40, command=self._run_batch)
        btn_batch.pack(fill="x", pady=20)

    def _run_batch(self):
        path = self.batch_ent.get()
        if not path: return
        def _t():
            BatchProcessor().run_batch(path)
        threading.Thread(target=_t).start()
