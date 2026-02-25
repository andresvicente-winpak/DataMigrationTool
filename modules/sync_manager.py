import pandas as pd
import os
import shutil
import datetime
from colorama import Fore, Style

class SyncManager:
    def __init__(self, rule_dir='config/rules'):
        self.rule_dir = rule_dir

    def compare_files(self, local_api_name, incoming_file_path):
        """
        Compares the Local Rule File against an Incoming Excel file.
        Returns:
            - diffs: List of dicts describing changes per field.
            - error: String if something goes wrong.
        """
        local_path = os.path.join(self.rule_dir, f"{local_api_name}.xlsx")
        
        if not os.path.exists(local_path):
            return [], f"Local rule file {local_api_name} not found."
        if not os.path.exists(incoming_file_path):
            return [], "Incoming file not found."

        try:
            # Load Rules
            df_local = pd.read_excel(local_path, sheet_name='Rules').fillna("").astype(str)
            df_incoming = pd.read_excel(incoming_file_path, sheet_name='Rules').fillna("").astype(str)
            
            # Index by Target Field for fast lookup
            # Note: We assume Target Field is unique per scope, but for simple merge we use Target Field as key
            # Better key: TARGET_FIELD + SCOPE
            
            # Helper to create unique key
            def make_key(row): return f"{row['TARGET_FIELD']}|{row.get('SCOPE', 'GLOBAL')}"
            
            df_local['KEY'] = df_local.apply(make_key, axis=1)
            df_incoming['KEY'] = df_incoming.apply(make_key, axis=1)
            
            local_map = df_local.set_index('KEY').to_dict('index')
            incoming_map = df_incoming.set_index('KEY').to_dict('index')
            
            diffs = []
            
            # 1. Check for Modified or New Rules
            for key, inc_row in incoming_map.items():
                target = inc_row['TARGET_FIELD']
                scope = inc_row.get('SCOPE', 'GLOBAL')
                
                if key not in local_map:
                    # NEW RULE
                    diffs.append({
                        'Type': 'NEW',
                        'Target': target,
                        'Scope': scope,
                        'Local': '(Not Existent)',
                        'Incoming': f"{inc_row['RULE_TYPE']} -> {inc_row['SOURCE_FIELD'] or inc_row['RULE_VALUE']}",
                        'Raw_Incoming': inc_row
                    })
                else:
                    # EXISTNG RULE - CHECK DIFF
                    loc_row = local_map[key]
                    
                    # Compare critical fields
                    has_change = False
                    change_desc = []
                    
                    for field in ['RULE_TYPE', 'SOURCE_FIELD', 'RULE_VALUE']:
                        if inc_row[field] != loc_row[field]:
                            has_change = True
                            change_desc.append(f"{field}: {loc_row[field]} -> {inc_row[field]}")
                            
                    if has_change:
                        diffs.append({
                            'Type': 'CHANGE',
                            'Target': target,
                            'Scope': scope,
                            'Local': f"{loc_row['RULE_TYPE']} -> {loc_row['SOURCE_FIELD'] or loc_row['RULE_VALUE']}",
                            'Incoming': f"{inc_row['RULE_TYPE']} -> {inc_row['SOURCE_FIELD'] or inc_row['RULE_VALUE']}",
                            'Raw_Incoming': inc_row,
                            'Details': ", ".join(change_desc)
                        })

            return diffs, None

        except Exception as e:
            return [], str(e)

    def perform_merge(self, local_api_name, incoming_file_path, selected_diffs):
        """
        Applies selected changes from Incoming to Local.
        Also merges the Audit Logs.
        """
        local_path = os.path.join(self.rule_dir, f"{local_api_name}.xlsx")
        backup_path = os.path.join(self.rule_dir, ".snapshots", f"{local_api_name}_PRE_MERGE_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        # 1. Snapshot
        if not os.path.exists(os.path.dirname(backup_path)): os.makedirs(os.path.dirname(backup_path))
        shutil.copy2(local_path, backup_path)
        
        try:
            # --- A. MERGE RULES ---
            df_local = pd.read_excel(local_path, sheet_name='Rules')
            
            # Helper to match rows
            def get_row_idx(df, target, scope):
                # Find index where Target and Scope match
                mask = (df['TARGET_FIELD'] == target) & (df['SCOPE'] == scope)
                idx = df.index[mask].tolist()
                return idx[0] if idx else None

            updates_count = 0
            new_count = 0

            for item in selected_diffs:
                row_data = item['Raw_Incoming']
                target = item['Target']
                scope = item['Scope']
                
                idx = get_row_idx(df_local, target, scope)
                
                if idx is not None:
                    # Update Existing
                    for col in df_local.columns:
                        if col in row_data:
                            df_local.at[idx, col] = row_data[col]
                    updates_count += 1
                else:
                    # Append New (Must align columns)
                    # Convert dict to df then concat
                    new_row_df = pd.DataFrame([row_data])
                    # Filter to only relevant columns
                    new_row_df = new_row_df[[c for c in new_row_df.columns if c in df_local.columns]]
                    df_local = pd.concat([df_local, new_row_df], ignore_index=True)
                    new_count += 1

            # --- B. MERGE HISTORY ---
            # Load both logs
            try: df_hist_local = pd.read_excel(local_path, sheet_name='_Audit_Log')
            except: df_hist_local = pd.DataFrame()
            
            try: df_hist_inc = pd.read_excel(incoming_file_path, sheet_name='_Audit_Log')
            except: df_hist_inc = pd.DataFrame()
            
            # Add a marker to incoming logs so we know they were imported
            if not df_hist_inc.empty:
                df_hist_inc['DETAILS'] = df_hist_inc['DETAILS'].astype(str) + " [IMPORTED]"
            
            # Concatenate
            df_hist_final = pd.concat([df_hist_local, df_hist_inc])
            
            # Deduplicate (in case we imported this file before)
            # We assume TIMESTAMP + TARGET + ACTION + USER is a unique composite key
            df_hist_final.drop_duplicates(subset=['TIMESTAMP', 'TARGET_FIELD', 'ACTION', 'USER'], inplace=True)
            
            # Sort by Timestamp (Newest First)
            if 'TIMESTAMP' in df_hist_final.columns:
                df_hist_final['TIMESTAMP'] = pd.to_datetime(df_hist_final['TIMESTAMP'], errors='coerce')
                df_hist_final = df_hist_final.sort_values('TIMESTAMP', ascending=False)
                # Format back to string
                df_hist_final['TIMESTAMP'] = df_hist_final['TIMESTAMP'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna("Unknown")

            # --- C. SAVE ---
            with pd.ExcelWriter(local_path, engine='xlsxwriter') as writer:
                df_local.to_excel(writer, sheet_name='Rules', index=False)
                df_hist_final.to_excel(writer, sheet_name='_Audit_Log', index=False)
                
            return True, f"Merged {updates_count} updates and {new_count} new rules."

        except Exception as e:
            # Restore backup on failure
            shutil.copy2(backup_path, local_path)
            return False, f"Critical Error during merge: {e}"