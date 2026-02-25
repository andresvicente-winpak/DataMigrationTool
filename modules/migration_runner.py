import pandas as pd
import os
import glob
import datetime
from colorama import Fore, Style
from modules.config_loader import ConfigLoader
from modules.extractor import DataExtractor
from modules.sdt_writer import SDTWriter
from modules.transform_engine import FilterEngine
import modules.ui as ui

class MigrationRunner:
    def __init__(self, map_path_override=None):
        self.output_dir = 'output'
        ui.ensure_folder(self.output_dir)
        self.map_path = map_path_override if map_path_override else 'config/migration_map.csv'

    def _resolve_from_map(self, lookup_val, lookup_col):
        current_map = self.map_path
        
        if not os.path.exists(current_map): 
            abs_path = os.path.join(os.getcwd(), current_map)
            if os.path.exists(abs_path): current_map = abs_path
            else:
                print(f"{Fore.RED}Map file missing: {current_map}{Style.RESET_ALL}")
                return None, None, None
        
        try:
            df = pd.read_csv(current_map).fillna("")
            df.columns = [c.upper().strip() for c in df.columns]
            
            if not lookup_val: return None, None, None
            lookup_val_norm = str(lookup_val).strip().upper()
            lookup_col_norm = lookup_col.strip().upper()
            
            if lookup_col_norm not in df.columns: return None, None, None

            df['LOOKUP_NORM'] = df[lookup_col_norm].astype(str).str.strip().str.upper()
            match = df[df['LOOKUP_NORM'] == lookup_val_norm]
            
            if match.empty and lookup_col_norm == 'MCO_SHEET':
                match = df[df['LOOKUP_NORM'].str.contains(lookup_val_norm, regex=False, na=False)]
            
            if match.empty: return None, None, None
            
            row = match.iloc[0]
            api = row.get('API_NAME', 'Unknown')
            sdt = row.get('SDT_TEMPLATE', '')
            
            raw_sheets = str(row.get('TRANSACTION_SHEET', ''))
            sheets = [s.strip() for s in raw_sheets.split(',') if s.strip()]
            
            return api, sdt, sheets
            
        except Exception as e:
            print(f"{Fore.RED}Error reading migration map: {e}{Style.RESET_ALL}")
            return None, None, None

    def _get_unique_filename(self, directory, filename):
        base, ext = os.path.splitext(filename)
        counter = 1
        new_filename = filename
        while os.path.exists(os.path.join(directory, new_filename)):
            new_filename = f"{base}_v{counter}{ext}"
            counter += 1
        return new_filename

    def execute_migration(self, program_name, legacy_path, division="GLOBAL", target_sheets=None, silent=False, output_name_override=None, append_if_exists=False, mco_context=None):
        try:
            # --- CRITICAL FIX: MAP RESOLUTION LOGIC ---
            if mco_context:
                # If we know the MCO Sheet (from UI), use it as the MASTER KEY.
                # This handles cases where one API is used by multiple sheets differently.
                map_api, map_sdt, map_sheets = self._resolve_from_map(mco_context, 'MCO_SHEET')
                
                # Check consistency (optional but good for debugging)
                if map_api and map_api != program_name:
                    if not silent:
                        print(f"{Fore.YELLOW}   [Info] API in Map ({map_api}) differs from selection ({program_name}). Using Map definition.{Style.RESET_ALL}")
                    program_name = map_api # Trust the map
            else:
                # Legacy behavior: Look up by API Name (ambiguous if duplicates exist)
                _, map_sdt, map_sheets = self._resolve_from_map(program_name, 'API_NAME')
            
            # --- SDT RESOLUTION ---
            sdt_path = None
            if map_sdt:
                if os.path.exists(os.path.join('config/sdt_templates', map_sdt)):
                    sdt_path = os.path.join('config/sdt_templates', map_sdt)
                elif os.path.exists(map_sdt): sdt_path = map_sdt
            
            if not sdt_path and program_name and program_name != "Unknown":
                candidates = glob.glob(f"config/sdt_templates/{program_name}*.xlsx")
                if candidates: sdt_path = candidates[0]

            if not sdt_path:
                print(f"{Fore.YELLOW}   Warning: Could not auto-resolve SDT Template for '{program_name}'.{Style.RESET_ALL}")
                return

            if not target_sheets:
                target_sheets = map_sheets if map_sheets else [ui.get_sheet_selection(sdt_path)]
                if not target_sheets[0]: return

            if not silent: print(f"\n{Fore.CYAN}--- STARTING MIGRATION ({program_name}) ---{Style.RESET_ALL}")
            
            config_loader = ConfigLoader(program_name)
            rules, lookups = config_loader.load_config(division_code=division)
            
            extractor = DataExtractor()
            df_legacy = extractor.load_data(legacy_path, format_type='MOVEX', sheet_name=0)

            filter_engine = FilterEngine(rules)
            df_legacy = filter_engine.apply_filters(df_legacy)

            if df_legacy.empty:
                print(f"{Fore.RED}   [ABORT] All rows filtered out.{Style.RESET_ALL}")
                return

            writer = SDTWriter(self.output_dir)
            
            if output_name_override:
                out_name = output_name_override
            else:
                date_str = datetime.datetime.now().strftime("%Y%m%d")
                base_name = f"LOAD_{program_name}_{date_str}.xlsx"
                
                if append_if_exists:
                    out_name = base_name
                else:
                    out_name = self._get_unique_filename(self.output_dir, base_name)
            
            writer.generate_from_template(sdt_path, df_legacy, rules, target_sheets, out_name, append_if_exists=append_if_exists)
            
        except Exception as e:
            print(f"{Fore.RED}FATAL ERROR: {e}{Style.RESET_ALL}")
            import traceback
            traceback.print_exc()

    def resolve_from_map_public(self, val, col):
        return self._resolve_from_map(val, col)