import pandas as pd
import re
import os
from colorama import Fore, Style

class AutoDetector:
    def __init__(self, mco_path):
        self.mco_path = mco_path
        self.prefix_map = {}  # {'MM': {'sheet': 'Item Master', 'api': 'MMS200MI'}}
        self.is_learned = False

    def learn_signatures(self):
        """
        Scans the MCO file to map 2-char Prefixes -> MCO Sheets -> API Names.
        """
        print(f"{Fore.CYAN}   Analyzing MCO for Field Signatures...{Style.RESET_ALL}")
        
        try:
            xls = pd.ExcelFile(self.mco_path)
            learned_list = []
            
            for sheet in xls.sheet_names:
                # 1. READ HEADER (Try Row 3 first, then scan)
                try:
                    df = pd.read_excel(xls, sheet_name=sheet, header=2, nrows=5)
                except Exception:
                    continue

                # 2. VALIDATE IF MCO SHEET
                cols_upper = [str(c).upper() for c in df.columns]
                
                # Check for Target and Source columns (handling typos)
                has_target = any('FIELD NAME' in c or 'M3 FIELD' in c for c in cols_upper)
                has_source = any('SOURCE' in c or 'CONVERISON' in c or 'CONVERSION' in c for c in cols_upper)
                
                if not (has_target and has_source):
                    continue

                # 3. EXTRACT PREFIX from Data Converison Source / Source column
                src_col_idx = next(
                    (i for i, c in enumerate(cols_upper) if 'SOURCE' in c or 'CONVERISON' in c or 'CONVERSION' in c),
                    None
                )
                if src_col_idx is None:
                    continue
                
                # Read full column
                df_full = pd.read_excel(xls, sheet_name=sheet, header=2, usecols=[src_col_idx])
                sample_values = df_full.iloc[:, 0].dropna().astype(str).tolist()
                
                prefix = None
                for val in sample_values:
                    clean_val = val.strip().upper()
                    # Accept the first 2 *alphanumeric* characters as the prefix
                    # This allows 'MMDIM1', 'OBV1', 'M9FACI', etc.
                    if len(clean_val) >= 2 and clean_val[:2].isalnum():
                        prefix = clean_val[:2]
                        break
                
                if not prefix:
                    # Optional debug: we had a source column but no usable prefix
                    print(
                        f"{Fore.YELLOW}      [AutoDetector] No usable 2-char prefix found in "
                        f"source column for sheet '{sheet}'. Sample values: "
                        f"{sample_values[:5]}{Style.RESET_ALL}"
                    )
                    continue

                # 4. GUESS API NAME
                api_name = "Unknown"
                
                # Try to find API Name in columns
                api_col_idx = next((i for i, c in enumerate(cols_upper) if 'API' in c), None)
                if api_col_idx is not None:
                    df_api = pd.read_excel(xls, sheet_name=sheet, header=2, usecols=[api_col_idx], nrows=10)
                    for val in df_api.iloc[:, 0].dropna().astype(str):
                        match = re.search(r'([A-Z]{3}\d{3}MI)', val, re.IGNORECASE)
                        if match:
                            api_name = match.group(1).upper()
                            break
                
                # Store Logic
                self.prefix_map[prefix] = {
                    'sheet': sheet,
                    'api': api_name
                }
                learned_list.append(f"{prefix} -> {sheet}")

            self.is_learned = True
            
            # VERBOSE OUTPUT
            print(f"{Fore.GREEN}   -> Learned signatures for {len(self.prefix_map)} Business Objects:{Style.RESET_ALL}")
            # Print in 3 columns to save space
            for i in range(0, len(learned_list), 3):
                print("      " + "   |   ".join(learned_list[i:i+3]))

        except Exception as e:
            print(f"{Fore.RED}Error learning MCO signatures: {e}{Style.RESET_ALL}")

    def identify_file(self, file_path):
        """
        Reads a Movex Excel file and returns the matching MCO info.
        """
        if not self.is_learned:
            self.learn_signatures()

        try:
            # Read first few rows to find headers
            df = pd.read_excel(file_path, nrows=10, header=None)
            
            header_row = None
            header_idx = 0
            
            # Heuristic Hunt
            for idx, row in df.iterrows():
                row_str = " ".join([str(x).upper() for x in row.values])
                # Look for common Movex fields
                if "ITNO" in row_str or "CUNO" in row_str or "SUNO" in row_str or "CONO" in row_str:
                    header_row = row
                    header_idx = idx
                    break
            
            if header_row is None:
                print(f"{Fore.YELLOW}   Warning: Could not detect standard Movex headers (ITNO/CUNO). Using Row 1.{Style.RESET_ALL}")
                header_row = df.iloc[0]

            print(f"   -> Scanning File Headers (Row {header_idx+1})...")

            # Analyze Headers
            for col in header_row.values:
                val = str(col).strip().upper()
                if len(val) >= 2:
                    p = val[:2]  # Extract 'MM' from 'MMITNO', 'M9' from 'M9FACI', etc.
                    
                    if p in self.prefix_map:
                        info = self.prefix_map[p]
                        print(f"      Matched Prefix '{p}' -> {info['sheet']}")
                        return p, info['sheet'], info['api']
            
            # Debug: Show what we checked against
            print(f"{Fore.RED}   No matching prefix found in headers.{Style.RESET_ALL}")
            print(f"   Checked against: {list(self.prefix_map.keys())}")
            return None, None, None

        except Exception as e:
            print(f"{Fore.RED}Error analyzing file: {e}{Style.RESET_ALL}")
            return None, None, None
