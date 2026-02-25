import pandas as pd
import os
import warnings
from colorama import Fore, Style
from modules.extractor import DataExtractor

class SurgicalExtractor:
    def __init__(self):
        self.config_dir = 'config'
        self.staging_dir = 'surgical_staging'
        if not os.path.exists(self.staging_dir):
            os.makedirs(self.staging_dir)
        self.extractor = DataExtractor()
        
    def _load_csv(self, filename):
        path = os.path.join(self.config_dir, filename)
        if not os.path.exists(path): return pd.DataFrame()
        try:
            df = pd.read_csv(path).fillna("")
            df.columns = [c.upper().strip() for c in df.columns]
            return df
        except: return pd.DataFrame()

    def get_available_objects(self):
        df = self._load_csv('surgical_def.csv')
        if df.empty: return []
        if 'OBJECT_TYPE' not in df.columns: return []
        return sorted(df['OBJECT_TYPE'].unique().tolist())

    def perform_extraction(self, object_type, id_list):
        print(f"\n{Fore.CYAN}--- SURGICAL EXTRACTION: {object_type} ---{Style.RESET_ALL}")
        
        warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

        # 1. Load Configurations
        df_def = self._load_csv('surgical_def.csv')
        df_source = self._load_csv('source_map.csv')
        df_mig = self._load_csv('migration_map.csv')
        
        if df_def.empty or df_source.empty:
            print(f"{Fore.RED}Missing definition files (surgical_def or source_map).{Style.RESET_ALL}")
            return []

        # 2. Filter definition by Object Type
        subset = df_def[df_def['OBJECT_TYPE'] == object_type]
        if subset.empty:
            print(f"No definitions found for {object_type}")
            return []

        tasks = []

        # 3. Create Lookup Maps (Normalized Keys)
        # Source Map: MCO_SHEET -> SOURCE_FILE
        source_lookup = {}
        for _, row in df_source.iterrows():
            key = str(row.get('MCO_SHEET', '')).strip().upper()
            if key: source_lookup[key] = row.get('SOURCE_FILE', '')

        # Migration Map: MCO_SHEET -> API_NAME
        api_lookup = {}
        for _, row in df_mig.iterrows():
            key = str(row.get('MCO_SHEET', '')).strip().upper()
            if key: api_lookup[key] = row.get('API_NAME', '')

        # 4. Iterate through definition lines
        for _, row in subset.iterrows():
            sheet = row.get('MCO_SHEET', '')
            sheet_key = str(sheet).strip().upper()
            
            if 'KEY_COLUMN' in row:
                key_col = row['KEY_COLUMN']
            else:
                print(f"{Fore.RED}      [ERROR] 'KEY_COLUMN' missing in surgical_def.csv.{Style.RESET_ALL}")
                continue
            
            print(f"   Processing scope: {sheet}...")
            
            # Resolve Source File
            source_path = source_lookup.get(sheet_key, '')
            if not source_path:
                print(f"{Fore.YELLOW}      No Source Map entry for {sheet}. Skipping.{Style.RESET_ALL}")
                continue

            # Resolve API Name (The fix for 'Unknown')
            api_name = api_lookup.get(sheet_key, 'Unknown')
            
            # Bypass File Existence Check for SQL
            is_sql = str(source_path).strip().upper().startswith("SQL:")
            
            if not is_sql and not os.path.exists(source_path):
                print(f"{Fore.RED}      Source file not found: {source_path}{Style.RESET_ALL}")
                continue

            try:
                # Load Full Data (File or SQL)
                df_full = self.extractor.load_data(source_path)
                
                # Check Key Column
                cols_norm = {c.upper().strip(): c for c in df_full.columns}
                target_key = str(key_col).upper().strip()
                
                if target_key not in cols_norm:
                    print(f"{Fore.RED}      Key column '{key_col}' not found in source.{Style.RESET_ALL}")
                    print(f"      Available: {list(cols_norm.keys())}")
                    continue
                
                real_key_col = cols_norm[target_key]
                
                # Filter Data
                id_list_str = [str(x).strip().upper() for x in id_list]
                mask = df_full[real_key_col].astype(str).str.strip().str.upper().isin(id_list_str)
                df_filtered = df_full[mask].copy()
                
                if df_filtered.empty:
                    print(f"{Fore.YELLOW}      No matches found for provided IDs.{Style.RESET_ALL}")
                    continue

                # Clean Up Columns (Aliases)
                new_cols = {}
                for col in df_filtered.columns:
                    col_u = col.upper()
                    if col_u.startswith('MM'):
                        alias = 'M9' + col_u[2:]
                        if alias not in cols_norm: new_cols[alias] = df_filtered[col]
                    if col_u.startswith('M9'):
                        alias = 'MM' + col_u[2:]
                        if alias not in cols_norm: new_cols[alias] = df_filtered[col]
                    if col_u[:2] in ['UA', 'UB', 'UC']:
                        alias = 'OK' + col_u[2:]
                        if alias not in cols_norm: new_cols[alias] = df_filtered[col]
                    if col_u.startswith('II') or col_u.startswith('IB'):
                        alias = 'ID' + col_u[2:]
                        if alias not in cols_norm: new_cols[alias] = df_filtered[col]

                if new_cols:
                    df_filtered = pd.concat([df_filtered, pd.DataFrame(new_cols, index=df_filtered.index)], axis=1)

                # Save Staged File
                safe_sheet = "".join([c if c.isalnum() else "_" for c in sheet])
                # Using resolved API name here
                staged_name = f"STAGED_{api_name}_{safe_sheet}.xlsx"
                staged_path = os.path.join(self.staging_dir, staged_name)
                
                df_filtered.to_excel(staged_path, index=False)
                
                print(f"{Fore.GREEN}      -> Staged {len(df_filtered)} rows to {staged_path}{Style.RESET_ALL}")
                
                tasks.append({
                    'program_name': api_name,
                    'legacy_path': staged_path,
                    'mco_sheet': sheet 
                })

            except Exception as e:
                print(f"{Fore.RED}      Error processing {sheet}: {e}{Style.RESET_ALL}")
                import traceback
                traceback.print_exc()
                
        return tasks