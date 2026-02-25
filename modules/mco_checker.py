import pandas as pd
import os
from colorama import Fore, Style

class MCOChecker:
    def __init__(self):
        pass

    def check_file(self, mco_path):
        if not os.path.exists(mco_path):
            print(f"{Fore.RED}File not found: {mco_path}{Style.RESET_ALL}")
            return

        print(f"\n{Fore.CYAN}--- ANALYZING MCO HEALTH: {os.path.basename(mco_path)} ---{Style.RESET_ALL}")
        
        try:
            xls = pd.ExcelFile(mco_path)
        except Exception as e:
            print(f"{Fore.RED}Critical Error: Cannot open Excel file. {e}{Style.RESET_ALL}")
            return

        total_issues = 0
        
        for sheet in xls.sheet_names:
            issues = self._analyze_sheet(mco_path, sheet)
            if issues:
                total_issues += len(issues)
                print(f"\n{Fore.YELLOW}SHEET: {sheet}{Style.RESET_ALL}")
                for issue in issues:
                    color = Fore.RED if "CRITICAL" in issue else Fore.YELLOW
                    print(f"   {color}{issue}{Style.RESET_ALL}")
        
        if total_issues == 0:
            print(f"\n{Fore.GREEN}RESULT: HEALTHY. No obvious structural issues found.{Style.RESET_ALL}")
        else:
            print(f"\n{Fore.RED}RESULT: {total_issues} ISSUES FOUND. Please review above.{Style.RESET_ALL}")

    def _analyze_sheet(self, path, sheet_name):
        issues = []
        
        try:
            # Read raw to find headers
            df_raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=15)
            
            header_idx = -1
            keywords = ['FIELD NAME', 'M3 FIELD', 'TECHNICAL NAME']
            for idx, row in df_raw.iterrows():
                row_str = " ".join([str(x).upper() for x in row.values])
                if any(k in row_str for k in keywords):
                    header_idx = idx
                    break
            
            if header_idx == -1:
                return [] # Skip non-spec sheets silently

            # Read Data
            df = pd.read_excel(path, sheet_name=sheet_name, header=header_idx)
            df.columns = [str(c).strip().replace('\n', ' ').replace('_', ' ').upper() for c in df.columns]

            # Columns
            target_aliases = ['FIELD NAME', 'M3 FIELD', 'TECHNICAL NAME']
            col_target = next((c for c in df.columns if any(a in c for a in target_aliases)), None)
            col_req = next((c for c in df.columns if 'CUSTOMER REQUIRED' in c or 'REQUIRED' in c), None)
            col_source = next((c for c in df.columns if 'CONVERSION SOURCE' in c or 'SOURCE' in c), None)
            col_logic = next((c for c in df.columns if 'TRANSFORMATION RULE' in c or 'LOGIC' in c), None)

            if not col_target: return ["CRITICAL: Missing 'M3 Field' column."]

            seen_targets = set()
            
            for idx, row in df.iterrows():
                excel_row = idx + header_idx + 2
                tgt = str(row.get(col_target, '')).strip().upper()
                
                if not tgt or tgt == 'NAN': continue
                if len(tgt) == 6: tgt = tgt[2:] # Normalize for dup check
                
                if tgt in seen_targets:
                    issues.append(f"Row {excel_row}: Duplicate Rule for Target '{tgt}'")
                seen_targets.add(tgt)

                if col_req:
                    req = str(row.get(col_req, '0')).strip().upper()
                    if req.startswith('1') or req.startswith('Y'):
                        src = str(row.get(col_source, '')).strip().replace('nan', '') if col_source else ""
                        logic = str(row.get(col_logic, '')).strip().replace('nan', '') if col_logic else ""
                        
                        if not src and not logic:
                             issues.append(f"Row {excel_row}: Field '{tgt}' is REQUIRED but has no Source/Logic.")

        except Exception as e:
            return [f"CRITICAL: Cannot read sheet. Error: {e}"]

        return issues