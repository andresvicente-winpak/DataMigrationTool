import pandas as pd
import os
import glob
import math
from colorama import Fore, Style
from modules.audit_manager import AuditManager

class RuleManager:
    def __init__(self, rule_dir='config/rules'):
        self.rule_dir = rule_dir
        if not os.path.exists(self.rule_dir):
            os.makedirs(self.rule_dir)
            
    # =========================================================================
    # GUI / API METHODS
    # =========================================================================
    
    def get_available_rules(self):
        files = glob.glob(os.path.join(self.rule_dir, "*.xlsx"))
        return [os.path.basename(f).replace('.xlsx','') for f in files]

    def load_rules(self, api_name):
        """Returns the DataFrame and full path for a specific API."""
        path = os.path.join(self.rule_dir, f"{api_name}.xlsx")
        if not os.path.exists(path):
            return None, None
        
        try:
            # FIX: keep_default_na=False ensures strings like "NA" are read as "NA", not NaN
            df = pd.read_excel(path, sheet_name='Rules', keep_default_na=False)
            
            # Convert to object to prevent FutureWarning when filling NaNs
            df = df.astype(object)
            df.fillna("", inplace=True)
            
            # Ensure metadata columns exist
            for c in ['BUSINESS_DESC', 'M3_TYPE', 'M3_LENGTH', 'M3_DECIMALS', 'SCOPE']:
                if c not in df.columns: df[c] = ""
                
            return df, path
        except Exception as e:
            print(f"Error loading rules: {e}")
            return None, None

    def validate_const(self, val, m3_type, m3_len):
        """Returns error message string if invalid, else None."""
        val = str(val)
        
        # Length Check
        if m3_len and str(m3_len).split('.')[0].isdigit():
            max_len = int(str(m3_len).split('.')[0])
            if len(val) > max_len:
                return f"Value '{val}' exceeds max length {max_len}."

        # Numeric Check
        if m3_type and ('DECIMAL' in str(m3_type).upper() or 'NUMERIC' in str(m3_type).upper()):
            try: float(val)
            except: return f"Value '{val}' must be numeric for type {m3_type}."
            
        return None

    def save_rule_update(self, api_name, row_index, updates):
        """
        Updates a specific row in the Excel file and commits audit.
        updates: dict of {col: value}
        """
        df, path = self.load_rules(api_name)
        if df is None: return False, "Could not load file."
        
        try:
            # Apply updates
            for col, val in updates.items():
                df.at[row_index, col] = val
            
            # Save
            with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
                df.to_excel(w, sheet_name='Rules', index=False)
            
            # Audit
            AuditManager(self.rule_dir).commit_changes(api_name)
            return True, "Saved successfully."
            
        except Exception as e:
            return False, str(e)

    def create_scope_override(self, api_name, original_idx, new_scope):
        """
        Creates a copy of the rule at original_idx but with new_scope.
        Returns (Success, Message, NewIndex)
        """
        df, path = self.load_rules(api_name)
        if df is None: return False, "Load failed", None
        
        target_field = df.iloc[original_idx]['TARGET_FIELD']
        
        # 1. Check if exists
        existing = df[(df['TARGET_FIELD'] == target_field) & (df['SCOPE'] == new_scope)]
        if not existing.empty:
            return False, f"Rule for {target_field} in scope {new_scope} already exists.", existing.index[0]
            
        # 2. Create Copy
        new_row = df.iloc[original_idx].copy()
        new_row['SCOPE'] = new_scope
        new_row['DESCRIPTION'] = f"Override for {new_scope}"
        
        # 3. Append
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
        # 4. Save
        try:
            with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
                df.to_excel(w, sheet_name='Rules', index=False)
            
            AuditManager(self.rule_dir).commit_changes(api_name)
            
            # Return the index of the new row (last one)
            return True, "Override created.", len(df) - 1
            
        except Exception as e:
            return False, str(e), None

    def merge_draft_file(self, draft_path, program_name, scope='GLOBAL', overwrite=False):
        if not os.path.exists(draft_path):
            print(f"{Fore.RED}Draft file not found: {draft_path}{Style.RESET_ALL}")
            return
        
        try:
            df_draft = pd.read_excel(draft_path)
            self.merge_draft_to_production(program_name, df_draft, scope, overwrite)
        except Exception as e:
            print(f"{Fore.RED}Error reading draft Excel: {e}{Style.RESET_ALL}")

    def merge_draft_to_production(self, program_name, df_draft, scope='GLOBAL', overwrite=False):
        path = f"{self.rule_dir}/{program_name}.xlsx"
        if not os.path.exists(path): 
            print(f"{Fore.RED}Target rule file not found: {path}{Style.RESET_ALL}")
            return

        print(f"Merging {len(df_draft)} patterns into {program_name}...")
        try:
            # FIX: keep_default_na=False here as well
            df_existing = pd.read_excel(path, sheet_name='Rules', keep_default_na=False)
            df_existing = df_existing.astype(object)
            df_existing.fillna("", inplace=True)
            
            target_map = {tgt: idx for idx, tgt in enumerate(df_existing['TARGET_FIELD'])}
            updates = 0
            
            for _, row in df_draft.iterrows():
                # Handle column name variations (TARGET vs FIELD_NAME)
                tgt = row.get('FIELD_NAME') or row.get('TARGET')
                src_val = row.get('SOURCE_FIELD') or row.get('SOURCE')
                
                if tgt in target_map:
                    idx = target_map[tgt]
                    curr_type = str(df_existing.at[idx, 'RULE_TYPE']).upper()
                    
                    if overwrite or curr_type in ['TODO', 'IGNORE', '', 'NAN']:
                        df_existing.at[idx, 'RULE_TYPE'] = row['TYPE']
                        df_existing.at[idx, 'SOURCE_FIELD'] = src_val
                        
                        if row['TYPE'] == 'CONST': 
                            df_existing.at[idx, 'RULE_VALUE'] = row['LOGIC']
                        elif row['TYPE'] == 'MAP': 
                            df_existing.at[idx, 'RULE_VALUE'] = f"MAP_{src_val}_TO_{tgt}"
                            
                        df_existing.at[idx, 'SCOPE'] = scope
                        df_existing.at[idx, 'DESCRIPTION'] = f"Auto-Detected ({row.get('CONFIDENCE','?')})"
                        updates += 1
            
            with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
                df_existing.to_excel(w, sheet_name='Rules', index=False)
            
            print(f"{Fore.GREEN}Merged {updates} rules.{Style.RESET_ALL}")
            AuditManager(self.rule_dir).commit_changes(program_name)

        except Exception as e:
            print(f"{Fore.RED}Merge failed: {e}{Style.RESET_ALL}")

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

    def interactive_manual_entry(self):
        files = glob.glob(os.path.join(self.rule_dir, "*.xlsx"))
        if not files:
            print(f"{Fore.RED}No rule files found. Please import MCO first.{Style.RESET_ALL}"); return

        filenames = [os.path.basename(f) for f in files]
        selected_name = self._smart_pick(filenames, "Select Rule File to Edit")
        selected_path = os.path.join(self.rule_dir, selected_name)
        program_name = selected_name.replace('.xlsx', '')

        try: 
            # FIX: keep_default_na=False here too for CLI
            df = pd.read_excel(selected_path, sheet_name='Rules', keep_default_na=False)
            df.fillna("", inplace=True)
            if 'BUSINESS_DESC' not in df.columns: df['BUSINESS_DESC'] = ""
        except: return

        for col in ['RULE_VALUE', 'SOURCE_FIELD', 'RULE_TYPE', 'DESCRIPTION', 'BUSINESS_DESC']:
            if col in df.columns: df[col] = df[col].astype(object)
        
        changes_made = False; current_page = 0; PAGE_SIZE = 20; active_filter = ""
        while True:
            df_view = df.copy()
            
            if active_filter: 
                mask = df_view['TARGET_FIELD'].str.contains(active_filter.upper(), na=False) | \
                       df_view['BUSINESS_DESC'].str.contains(active_filter, case=False, na=False)
                df_view = df_view[mask]

            total_rows = len(df_view); total_pages = math.ceil(total_rows / PAGE_SIZE)
            if total_pages == 0: total_pages = 1
            if current_page >= total_pages: current_page = 0
            start_idx = current_page * PAGE_SIZE; end_idx = start_idx + PAGE_SIZE
            df_page = df_view.iloc[start_idx:end_idx]

            print(f"\n{Fore.YELLOW}--- Editing {program_name} (Page {current_page+1}/{total_pages}) ---{Style.RESET_ALL}")
            if active_filter: print(f"{Fore.CYAN}[Filter Active: '{active_filter}']{Style.RESET_ALL}")
            
            print(f"{'#':<4} {'TARGET':<10} | {'DESCRIPTION':<30} | {'TYPE':<8} | {'SOURCE/VAL'}")
            print("-" * 80)
            
            display_map = {}; row_counter = 1
            for idx, row in df_page.iterrows():
                desc_disp = str(row['BUSINESS_DESC'])[:28]
                
                if row['RULE_TYPE'] in ['CONST', 'PYTHON']:
                    val_disp = f"Val: {str(row['RULE_VALUE'])[:15]}"
                else:
                    val_disp = f"Src: {str(row['SOURCE_FIELD'])[:15]}"

                color = Style.RESET_ALL
                if row['RULE_TYPE'] == 'TODO': color = Fore.RED
                elif row['RULE_TYPE'] == 'IGNORE': color = Fore.LIGHTBLACK_EX
                elif row['RULE_TYPE'] == 'MAP': color = Fore.BLUE
                
                print(f"{color}{row_counter:<4} {row['TARGET_FIELD']:<10} | {desc_disp:<30} | {row['RULE_TYPE']:<8} | {val_disp}{Style.RESET_ALL}")
                display_map[row_counter] = idx; row_counter += 1
            print("-" * 80)
            
            print("Commands: [N]ext, [P]revious, [F]ilter, [S]ave & Exit, [Q]uit")
            cmd = input(f"{Fore.CYAN}>> Enter # to Edit or Command: {Style.RESET_ALL}").strip().upper()
            
            if cmd == 'N': 
                if current_page < total_pages - 1: current_page += 1
            elif cmd == 'P': 
                if current_page > 0: current_page -= 1
            elif cmd == 'F': active_filter = input("Enter Filter Text (Enter to clear): ").strip(); current_page = 0
            elif cmd == 'S':
                if changes_made:
                    with pd.ExcelWriter(selected_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
                        df.to_excel(w, sheet_name='Rules', index=False)
                    AuditManager(self.rule_dir).commit_changes(program_name)
                    print(f"{Fore.GREEN}   -> Saved and Audited.{Style.RESET_ALL}")
                return
            elif cmd == 'Q': return
            elif cmd.isdigit() and int(cmd) in display_map:
                self._edit_row(df, display_map[int(cmd)])
                changes_made = True

    def _edit_row(self, df, idx):
        target = df.at[idx, 'TARGET_FIELD']
        desc = df.at[idx, 'BUSINESS_DESC']
        curr_type = str(df.at[idx, 'RULE_TYPE']).strip()
        curr_src = str(df.at[idx, 'SOURCE_FIELD']).strip()
        curr_val = str(df.at[idx, 'RULE_VALUE']).strip()

        print(f"\n{Fore.GREEN}EDITING: {target} ({desc}){Style.RESET_ALL}")
        print(f"Current: Type={curr_type}, Src={curr_src}, Val={curr_val}")
        print(f"{Fore.CYAN}Types: CONST, DIRECT, MAP, IGNORE, TODO, PYTHON{Style.RESET_ALL}")
        
        new_type = input(f"New Type ({curr_type}): ").strip().upper()
        if not new_type: new_type = curr_type
        if new_type == 'CONSTANT': new_type = 'CONST'
        
        new_src = curr_src; new_val = curr_val

        if new_type in ['DIRECT', 'PYTHON', 'MAP']:
            new_src = input(f"Source Field ({curr_src}): ").strip().upper()
            if not new_src: new_src = curr_src
        else:
            new_src = ""

        if new_type == 'CONST':
            new_val = input(f"Constant Value ({curr_val}): ").strip()
            if not new_val: new_val = curr_val
        elif new_type == 'PYTHON':
             print(f"{Fore.YELLOW}(Tip: Use GUI for multi-line code editing){Style.RESET_ALL}")
             new_val = input(f"Python Code ({curr_val}): ").strip()
             if not new_val: new_val = curr_val
        elif new_type == 'MAP':
            # Just a placeholder for CLI, GUI uses wizard
            new_val = input(f"Map Config ({curr_val}): ").strip()
            if not new_val: new_val = curr_val
        elif new_type in ['IGNORE', 'TODO']:
            new_val = ""

        df.at[idx, 'RULE_TYPE'] = new_type
        df.at[idx, 'SOURCE_FIELD'] = new_src
        df.at[idx, 'RULE_VALUE'] = new_val
        df.at[idx, 'DESCRIPTION'] = 'Manual Edit'
        print("Updated in memory.")