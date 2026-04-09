import pandas as pd
import os
import glob
from colorama import Fore, Style
import time
from modules.audit_manager import AuditManager

class MCOImporter:
    def __init__(self, sdt_folder='config/sdt_templates'):
        self.sdt_folder = sdt_folder

    # =========================================================================
    # GUI HELPERS
    # =========================================================================
    def get_sheet_names(self, mco_path):
        try:
            return pd.ExcelFile(mco_path, engine='openpyxl').sheet_names
        except Exception:
            try:
                return pd.ExcelFile(mco_path).sheet_names
            except Exception:
                return []

    def run_import_headless(self, mco_path, selected_sheet, api_name, output_dir='config/rules', overwrite_all=False):
        try:
            print(f"Loading MCO Sheet: {selected_sheet}...")
            df_mco = self._find_header_row(mco_path, selected_sheet)
            self._generate_master_rule_file(df_mco, api_name, output_dir, overwrite_all)
            return True
        except Exception as e:
            print(f"{Fore.RED}Import Error: {e}{Style.RESET_ALL}")
            return False

    # =========================================================================
    # CLI HELPERS
    # =========================================================================
    def _smart_pick(self, options, title_prompt):
        filtered_indices = list(range(len(options)))
        filter_text = ""
        while True:
            print(f"\n{Fore.CYAN}--- {title_prompt} ---{Style.RESET_ALL}")
            if filter_text: print(f"{Fore.YELLOW}[Filter: '{filter_text}'] (Type 'all' to clear){Style.RESET_ALL}")
            limit = 20; count = 0; display_map = {}
            for i in filtered_indices:
                count += 1
                if count > limit:
                    print(f"   ... ({len(filtered_indices) - limit} more matches)")
                    break
                print(f"   {count}. {options[i]}")
                display_map[count] = i
            
            print(f"\n{Fore.GREEN}Type a Number to select, or Text to filter.{Style.RESET_ALL}")
            user_input = input(f"{Fore.CYAN}>> Selection: {Style.RESET_ALL}").strip()
            if not user_input: continue
            
            if user_input.isdigit():
                choice = int(user_input)
                if choice in display_map: return options[display_map[choice]]
            elif user_input.lower() in ['all', 'clear']: 
                filter_text = ""; filtered_indices = list(range(len(options)))
            else:
                filter_text = user_input
                filtered_indices = [i for i, opt in enumerate(options) if filter_text.lower() in str(opt).lower()]
                if not filtered_indices: 
                    print(f"{Fore.RED}No matches found.{Style.RESET_ALL}")
                    time.sleep(0.5); filter_text = ""; filtered_indices = list(range(len(options)))

    def interactive_import(self, mco_path, output_dir='config/rules'):
        try:
            xls = pd.ExcelFile(mco_path)
            mco_sheets = xls.sheet_names
            selected_mco_sheet = self._smart_pick(mco_sheets, "Select MCO Sheet")
            df_mco = self._find_header_row(mco_path, selected_mco_sheet)
            
            default_name = selected_mco_sheet.split(' ')[0] + "MI"
            api_name = input(f"{Fore.CYAN}   >> Name this Rule Set (Default: {default_name}): {Style.RESET_ALL}").strip().upper()
            if not api_name: api_name = default_name.upper()
            
            self._generate_master_rule_file(df_mco, api_name, output_dir, False)

        except Exception as e:
            print(f"{Fore.RED}Error during import: {e}{Style.RESET_ALL}")
            import traceback; traceback.print_exc()

    # =========================================================================
    # CORE LOGIC
    # =========================================================================
    def _find_header_row(self, file_path, sheet_name):
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=15)
        header_idx = -1
        keywords = ['FIELD NAME', 'M3 FIELD', 'TECHNICAL NAME']
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(x).upper() for x in row.values])
            if any(k in row_str for k in keywords): header_idx = idx; break
        if header_idx == -1: header_idx = 2 
        print(f"      -> Detected headers on Row {header_idx + 1}")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_idx)
        df.columns = [str(c).strip().replace('\n', ' ').replace('_', ' ').upper() for c in df.columns]
        return df

    def _generate_master_rule_file(self, df_mco, api_name, output_dir, overwrite_all):
        if not os.path.exists(output_dir): os.makedirs(output_dir)
        target_path = f"{output_dir}/{api_name}.xlsx"
        
        existing_rules = pd.DataFrame()
        
        if os.path.exists(target_path):
            auditor = AuditManager(os.path.dirname(output_dir) if 'rules' in output_dir else output_dir) 
            auditor.create_snapshot(api_name, "AUTO_PRE_MCO_UPDATE")
            
            if not overwrite_all:
                print(f"{Fore.YELLOW}   Merging into existing rule file...{Style.RESET_ALL}")
                try:
                    existing_rules = pd.read_excel(target_path, sheet_name='Rules')
                    # Ensure columns exist
                    for c in ['BUSINESS_DESC', 'M3_TYPE', 'M3_LENGTH', 'M3_DECIMALS']:
                        if c not in existing_rules.columns: existing_rules[c] = ""
                except: pass
            else:
                print(f"{Fore.RED}   FORCE OVERWRITE: Resetting rules.{Style.RESET_ALL}")

        print(f"   -> Parsing MCO content...")
        cols = df_mco.columns
        
        col_target = next((c for c in cols if any(a in c for a in ['FIELD NAME', 'M3 FIELD', 'TECHNICAL NAME'])), None)
        col_req    = next((c for c in cols if 'CUSTOMER REQUIRED' in c or 'REQUIRED' in c), None)
        col_source = next((c for c in cols if 'CONVERSION SOURCE' in c or 'SOURCE' in c or 'LEGACY' in c), None)
        col_logic  = next((c for c in cols if 'TRANSFORMATION RULE' in c or 'LOGIC' in c or 'RULE' in c), None)
        col_desc   = next((c for c in cols if 'DESCRIPTION' in c), None)
        
        col_type   = next((c for c in cols if 'DATA TYPE' in c or 'TYPE' in c), None)
        col_len    = next((c for c in cols if 'LENGTH' in c), None)
        col_dec    = next((c for c in cols if 'DECIMAL' in c), None)

        if not col_target: print(f"{Fore.RED}      CRITICAL: Could not find Target Column.{Style.RESET_ALL}"); return

        new_rules = []
        for _, row in df_mco.iterrows():
            tgt = str(row.get(col_target, '')).strip().upper()
            if not tgt or tgt == 'NAN': continue
            if len(tgt) == 6: tgt = tgt[2:] 

            raw_src = str(row.get(col_source, '')).strip().replace('nan', '').upper()
            if len(raw_src) == 6 and raw_src[:2].isalpha(): raw_src = raw_src[2:]

            raw_req = str(row.get(col_req, '0')).strip().replace('nan', '0')
            raw_logic = str(row.get(col_logic, '')).strip().replace('nan', '')
            raw_desc = str(row.get(col_desc, '')).strip().replace('nan', '') if col_desc else ""
            
            m3_type = str(row.get(col_type, '')).strip().replace('nan', '') if col_type else ""
            m3_len  = str(row.get(col_len, '')).strip().replace('nan', '') if col_len else ""
            m3_dec  = str(row.get(col_dec, '')).strip().replace('nan', '') if col_dec else ""
            
            r_type, r_val, r_src, desc = 'IGNORE', '', '', 'Imported'
            
            if raw_src:
                r_type = 'DIRECT'; r_src = raw_src; desc = f"Mapped from {raw_src}"
            elif raw_req.startswith('1') or raw_req.startswith('Y'):
                if 'CONST' in raw_logic.upper() or 'FIXED' in raw_logic.upper():
                        r_type = 'CONST'; desc = f"Required Constant: {raw_logic}"
                else:
                    r_type = 'TODO'; desc = f"Required! Logic: {raw_logic}"
            else:
                r_type = 'IGNORE'; desc = "MCO listed but not required"
            
            new_rules.append({
                'TARGET_API': api_name, 
                'TARGET_FIELD': tgt, 
                'SOURCE_FIELD': r_src, 
                'RULE_TYPE': r_type, 
                'RULE_VALUE': r_val, 
                'SCOPE': 'GLOBAL', 
                'DESCRIPTION': desc,
                'BUSINESS_DESC': raw_desc,
                'M3_TYPE': m3_type,
                'M3_LENGTH': m3_len,
                'M3_DECIMALS': m3_dec
            })

        df_new = pd.DataFrame(new_rules)

        # --- MERGE LOGIC ---
        if not existing_rules.empty and not overwrite_all:
            final_rows = []
            new_rules_dict = {row['TARGET_FIELD']: row for _, row in df_new.iterrows()}
            
            # Keep existing rules including FILTERs that are not in MCO
            # (Because MCO usually only lists fields, not logic/filters)
            
            for idx, row in existing_rules.iterrows():
                tgt = row['TARGET_FIELD']
                # If existing is a FILTER, keep it regardless of MCO
                if row['RULE_TYPE'] == 'FILTER':
                    final_rows.append(row)
                    continue

                if tgt in new_rules_dict:
                    mco_data = new_rules_dict[tgt]
                    # Update metadata
                    row['BUSINESS_DESC'] = mco_data['BUSINESS_DESC']
                    row['M3_TYPE'] = mco_data['M3_TYPE']
                    row['M3_LENGTH'] = mco_data['M3_LENGTH']
                    row['M3_DECIMALS'] = mco_data['M3_DECIMALS']
                    
                    # Only update logic if weak
                    curr_type = str(row['RULE_TYPE']).upper()
                    if curr_type in ['TODO', 'IGNORE', '', 'NAN']:
                        row['RULE_TYPE'] = mco_data['RULE_TYPE']
                        row['SOURCE_FIELD'] = mco_data['SOURCE_FIELD']
                        row['RULE_VALUE'] = mco_data['RULE_VALUE']
                        row['DESCRIPTION'] = mco_data['DESCRIPTION']
                    
                    del new_rules_dict[tgt]
                final_rows.append(row)
            
            for tgt, data in new_rules_dict.items():
                final_rows.append(pd.Series(data))
            
            final_rules = pd.DataFrame(final_rows)
        else:
            final_rules = df_new.drop_duplicates(subset=['TARGET_FIELD'], keep='last')

        final_rules = final_rules.copy()
        
        def sorter(x):
            if x == 'FILTER': return 0 # Top Priority
            if x == 'TODO': return 1
            if x in ['DIRECT', 'CONST', 'MAP', 'PYTHON']: return 2
            return 3
            
        final_rules['Sort'] = final_rules['RULE_TYPE'].apply(sorter)
        final_rules = final_rules.sort_values(['Sort', 'TARGET_FIELD']).drop(columns=['Sort'])

        with pd.ExcelWriter(target_path, engine='xlsxwriter') as writer:
            final_rules.to_excel(writer, sheet_name='Rules', index=False)
            pd.DataFrame(columns=['TIMESTAMP']).to_excel(writer, sheet_name='_Audit_Log', index=False)
            
        print(f"      -> Config Saved: {target_path} (Total Fields: {len(final_rules)})")