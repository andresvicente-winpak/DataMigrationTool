import pandas as pd
import os
from colorama import Fore, Style
import time
from modules.audit_manager import AuditManager

class MCOImporter:
    AUTO_DESCRIPTION_PREFIXES = (
        'Mapped from ',
        'Lookup Table:',
        'Lookup Table (',
        'Constant:',
        'Constant (',
        'Required!',
        'Required Constant:',
        'MCO listed',
    )

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
    @staticmethod
    def _clean_cell(value, default=''):
        """Return a trimmed string without treating text containing 'nan' as empty."""
        if value is None or pd.isna(value):
            return default
        return str(value).strip()

    @staticmethod
    def _find_column(columns, aliases, excluded_terms=()):
        """Find the first normalized header matching an alias and no exclusions."""
        return next(
            (
                column
                for column in columns
                if any(alias in column for alias in aliases)
                and not any(term in column for term in excluded_terms)
            ),
            None,
        )

    @staticmethod
    def _classify_rule(raw_src, raw_req, raw_logic, raw_usage, raw_usage_comments):
        """Translate one MCO row into rule type, value, source, and description."""
        usage_upper = raw_usage.upper()

        # Field Usage is the customer's rule decision.  A blank decision must
        # remain blank rather than being inferred from M3-required or source
        # metadata, while an explicit Ignore must likewise take precedence.
        if not usage_upper:
            return '', '', '', ''

        if usage_upper in ['IGNORE', 'IGNORED']:
            return 'IGNORE', '', '', "Explicitly ignored in MCO Field Usage"

        if 'LOOKUP' in usage_upper or usage_upper in ['MAP', 'MAPPING']:
            description = (
                f"Lookup Table: {raw_usage_comments}"
                if raw_usage_comments
                else "Lookup Table (mapping configuration required)"
            )
            return 'MAP', raw_usage_comments, raw_src, description

        if 'CONSTANT' in usage_upper or usage_upper in ['CONST', 'FIXED']:
            rule_value = raw_usage_comments or raw_logic
            description = (
                f"Constant: {rule_value}"
                if rule_value
                else "Constant (value required)"
            )
            return 'CONST', rule_value, '', description

        if raw_src:
            return 'DIRECT', '', raw_src, f"Mapped from {raw_src}"

        if raw_req.startswith('1') or raw_req.startswith('Y'):
            if 'CONST' in raw_logic.upper() or 'FIXED' in raw_logic.upper():
                return 'CONST', raw_logic, '', f"Required Constant: {raw_logic}"
            return 'TODO', '', '', f"Required! Logic: {raw_logic}"

        return 'IGNORE', '', '', "MCO listed but not required"

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
        
        col_target = self._find_column(cols, ['FIELD NAME', 'M3 FIELD', 'TECHNICAL NAME'])
        # Prefer the customer's decision over the separate M3-required flag.
        # Both columns are present in the standard MCO layout, and the M3 flag
        # normally appears first.
        col_req = self._find_column(cols, ['CUSTOMER REQUIRED'])
        if col_req is None:
            col_req = self._find_column(cols, ['REQUIRED'])
        col_source = self._find_column(cols, ['CONVERSION SOURCE', 'SOURCE', 'LEGACY'])
        col_logic = self._find_column(cols, ['TRANSFORMATION RULE', 'LOGIC', 'RULE'])
        col_desc = self._find_column(cols, ['DESCRIPTION'])
        col_usage = self._find_column(cols, ['FIELD USAGE'], ['COMMENT'])
        col_usage_comments = self._find_column(cols, ['FIELD USAGE COMMENTS'])
        
        col_type = self._find_column(cols, ['DATA TYPE', 'TYPE'])
        col_len = self._find_column(cols, ['LENGTH'])
        col_dec = self._find_column(cols, ['DECIMAL'])

        if not col_target: print(f"{Fore.RED}      CRITICAL: Could not find Target Column.{Style.RESET_ALL}"); return

        new_rules = []
        for _, row in df_mco.iterrows():
            tgt = self._clean_cell(row.get(col_target)).upper()
            if not tgt: continue
            if len(tgt) == 6: tgt = tgt[2:] 

            raw_src = self._clean_cell(row.get(col_source)).upper()

            # The conversion source is the column name in the legacy extract.
            # M3 database columns commonly include a two-character table prefix
            # (for example, OKCUNO).  The old importer removed that prefix from
            # every six-character source, producing CUNO and rules that could
            # not find the actual input column.  Only target fields are
            # normalized to the four-character API field name; source names
            # must be preserved exactly as specified by the MCO.

            raw_req = self._clean_cell(row.get(col_req), default='0')
            raw_logic = self._clean_cell(row.get(col_logic))
            raw_desc = self._clean_cell(row.get(col_desc))
            raw_usage = self._clean_cell(row.get(col_usage))
            raw_usage_comments = self._clean_cell(row.get(col_usage_comments))
            
            m3_type = self._clean_cell(row.get(col_type))
            m3_len = self._clean_cell(row.get(col_len))
            m3_dec = self._clean_cell(row.get(col_dec))
            
            # Field Usage is authoritative and is evaluated before source and
            # required-field fallbacks by the classifier.
            r_type, r_val, r_src, desc = self._classify_rule(
                raw_src,
                raw_req,
                raw_logic,
                raw_usage,
                raw_usage_comments,
            )
            
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
                'M3_DECIMALS': m3_dec,
                # Internal merge hint; removed before the workbook is written.
                # An explicit MCO Field Usage is authoritative even if an older
                # generated workbook currently contains a strong rule type.
                '_MCO_EXPLICIT_USAGE': bool(raw_usage)
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
                    
                    # Only update weak rules or rules previously generated by
                    # this importer. This lets a corrected MCO classification
                    # (such as DIRECT -> MAP) replace an earlier automatic
                    # result without discarding a manually authored rule.
                    curr_type = str(row['RULE_TYPE']).upper()
                    curr_desc = str(row.get('DESCRIPTION', ''))
                    auto_generated = curr_desc.startswith(self.AUTO_DESCRIPTION_PREFIXES)
                    explicit_mco_usage = bool(mco_data.get('_MCO_EXPLICIT_USAGE', False))
                    if explicit_mco_usage or curr_type in ['TODO', 'IGNORE', '', 'NAN'] or auto_generated:
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
        final_rules.drop(columns=['_MCO_EXPLICIT_USAGE'], errors='ignore', inplace=True)
        
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
