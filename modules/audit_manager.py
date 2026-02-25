import pandas as pd
import os
import getpass
import datetime
import shutil
import glob
from colorama import Fore, Style

class AuditManager:
    def __init__(self, rule_dir='config/rules'):
        self.rule_dir = rule_dir
        self.history_dir = f"{rule_dir}/.history"
        self.snapshot_dir = f"{rule_dir}/.snapshots"
        
        for d in [self.history_dir, self.snapshot_dir]:
            if not os.path.exists(d): os.makedirs(d)

    def create_snapshot(self, program_name, note="Manual"):
        src = f"{self.rule_dir}/{program_name}.xlsx"
        if not os.path.exists(src): return
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{program_name}_{timestamp}_{note}.xlsx"
        dst = f"{self.snapshot_dir}/{filename}"
        shutil.copy2(src, dst)

    def list_snapshots(self, program_name):
        pattern = f"{self.snapshot_dir}/{program_name}_*.xlsx"
        files = glob.glob(pattern)
        files.sort(key=os.path.getmtime, reverse=True)
        return [os.path.basename(f) for f in files]

    def restore_snapshot(self, program_name, snapshot_filename):
        src = f"{self.snapshot_dir}/{snapshot_filename}"
        dst = f"{self.rule_dir}/{program_name}.xlsx"
        if not os.path.exists(src): return
        self.create_snapshot(program_name, "PRE_RESTORE_BACKUP")
        shutil.copy2(src, dst)
        print(f"{Fore.GREEN}   -> Restored {program_name} from {snapshot_filename}{Style.RESET_ALL}")

    def get_history_dataframe(self, program_name, filter_text=None):
        """
        Returns the history as a Pandas DataFrame for the GUI.
        """
        path = f"{self.rule_dir}/{program_name}.xlsx"
        if not os.path.exists(path):
            return pd.DataFrame()

        try:
            df = pd.read_excel(path, sheet_name='_Audit_Log')
            
            # Standardize Columns
            expected_cols = ['TIMESTAMP', 'USER', 'ACTION', 'TARGET_FIELD', 'DETAILS']
            for col in expected_cols:
                if col not in df.columns:
                    df[col] = ""

            # Filter
            if filter_text:
                mask = df['TARGET_FIELD'].astype(str).str.contains(filter_text.upper(), na=False) | \
                       df['USER'].astype(str).str.contains(filter_text.upper(), na=False) | \
                       df['ACTION'].astype(str).str.contains(filter_text.upper(), na=False)
                df = df[mask]

            # --- FIX: Standardize Timestamps ---
            if 'TIMESTAMP' in df.columns:
                # Force conversion to datetime, turning errors (like bad strings) into NaT
                df['TIMESTAMP'] = pd.to_datetime(df['TIMESTAMP'], errors='coerce')
                # Sort descending (Latest first)
                df = df.sort_values('TIMESTAMP', ascending=False)
                # Convert back to string for display so NaT doesn't look ugly
                df['TIMESTAMP'] = df['TIMESTAMP'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna("Unknown")

            return df[expected_cols]
        except Exception as e:
            print(f"Error getting history: {e}")
            return pd.DataFrame()

    def view_history(self, program_name, filter_field=None):
        # Console-only viewer (legacy support)
        df = self.get_history_dataframe(program_name, filter_field)
        if df.empty:
            print(f"{Fore.YELLOW}No history found.{Style.RESET_ALL}"); return

        print(f"\n{Fore.CYAN}--- RULE HISTORY: {program_name} ---{Style.RESET_ALL}")
        print(f"{'TIMESTAMP':<20} | {'USER':<10} | {'ACTION':<6} | {'TARGET':<10} | {'DETAILS'}")
        print("-" * 120)
        for _, row in df.iterrows():
            ts = str(row.get('TIMESTAMP', ''))[:19]
            user = str(row.get('USER', ''))[:10]
            action = str(row.get('ACTION', ''))[:6]
            target = str(row.get('TARGET_FIELD', ''))[:10]
            details = str(row.get('DETAILS', ''))
            if len(details) > 60: details = details[:57] + "..."
            print(f"{ts:<20} | {user:<10} | {action:<6} | {target:<10} | {details}")
        print("-" * 120)

    def commit_changes(self, program_name):
        self.create_snapshot(program_name, "Auto_Commit")
        current_path = f"{self.rule_dir}/{program_name}.xlsx"
        history_path = f"{self.history_dir}/{program_name}.xlsx"
        if not os.path.exists(current_path): return

        try: df_curr = pd.read_excel(current_path, sheet_name='Rules').fillna('').astype(str)
        except: return
        
        df_curr = df_curr.drop_duplicates(subset=['TARGET_FIELD'], keep='last')

        changes = []
        if os.path.exists(history_path):
            df_prev = pd.read_excel(history_path, sheet_name='Rules').fillna('').astype(str)
            df_prev = df_prev.drop_duplicates(subset=['TARGET_FIELD'], keep='last')
            
            merged = pd.merge(df_curr, df_prev, on='TARGET_FIELD', how='outer', suffixes=('_NEW', '_OLD'), indicator=True)
            user = getpass.getuser()
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            for idx, row in merged.iterrows():
                target = row['TARGET_FIELD']
                if row['_merge'] == 'left_only':
                    changes.append([ts, user, 'ADD', target, 'New Rule Created'])
                elif row['_merge'] == 'right_only':
                     changes.append([ts, user, 'DELETE', target, 'Rule Deleted'])
                elif row['_merge'] == 'both':
                    diffs = []
                    if row['RULE_TYPE_NEW'] != row['RULE_TYPE_OLD']: diffs.append(f"Type: {row['RULE_TYPE_OLD']}->{row['RULE_TYPE_NEW']}")
                    if row['SOURCE_FIELD_NEW'] != row['SOURCE_FIELD_OLD']: diffs.append(f"Src: {row['SOURCE_FIELD_OLD']}->{row['SOURCE_FIELD_NEW']}")
                    if row['RULE_VALUE_NEW'] != row['RULE_VALUE_OLD']: diffs.append(f"Val: '{row['RULE_VALUE_OLD']}'->'{row['RULE_VALUE_NEW']}'")
                    
                    if diffs:
                        changes.append([ts, user, 'EDIT', target, ", ".join(diffs)])
        else:
            changes.append([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), getpass.getuser(), 'INIT', 'ALL', 'Initial Commit'])

        if changes:
            with pd.ExcelWriter(current_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                try: df_old = pd.read_excel(current_path, sheet_name='_Audit_Log')
                except: df_old = pd.DataFrame(columns=['TIMESTAMP', 'USER', 'ACTION', 'TARGET_FIELD', 'DETAILS'])
                
                df_new = pd.DataFrame(changes, columns=['TIMESTAMP', 'USER', 'ACTION', 'TARGET_FIELD', 'DETAILS'])
                
                if not df_new.empty:
                    # --- FIX: Avoid Concatenation Warning ---
                    if df_old.empty:
                        df_final = df_new
                    else:
                        df_final = pd.concat([df_old, df_new], ignore_index=True)
                    
                    df_final.to_excel(writer, sheet_name='_Audit_Log', index=False)
        
        shutil.copy2(current_path, history_path)
    
    def hard_reset(self):
        print(f"{Fore.YELLOW}   Cleaning history...{Style.RESET_ALL}")
        for folder in [self.history_dir, self.snapshot_dir]:
            files = glob.glob(os.path.join(folder, "*"))
            for f in files: os.remove(f)
        
        rule_files = glob.glob(os.path.join(self.rule_dir, "*.xlsx"))
        empty_log = pd.DataFrame(columns=['TIMESTAMP', 'USER', 'ACTION', 'TARGET_FIELD', 'DETAILS'])
        
        for f in rule_files:
            try:
                with pd.ExcelWriter(f, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
                    empty_log.to_excel(w, sheet_name='_Audit_Log', index=False)
            except: pass