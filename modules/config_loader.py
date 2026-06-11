import pandas as pd
import os

class ConfigLoader:
    def __init__(self, program_name, rule_dir='config/rules'):
        self.program_name = program_name
        self.rule_dir = rule_dir
        self.file_path = f"{self.rule_dir}/{program_name}.xlsx"
        self.rules_raw = None
        self.lookups = {}

    def load_config(self, division_code='GLOBAL'):
        if not os.path.exists(self.file_path):
             return pd.DataFrame(), {}
        
        print(f"Loading configuration for {self.program_name} (Scope: {division_code})...")
        
        try:
            # FIX: keep_default_na=False prevents 'NA' from becoming NaN (Null)
            df = pd.read_excel(self.file_path, sheet_name='Rules', keep_default_na=False)
            df.columns = [c.upper().strip() for c in df.columns]
            
            if 'SCOPE' not in df.columns:
                df['SCOPE'] = 'GLOBAL'
            
            # Helper to check scope (Handles "DIV_US, DIV_CA")
            def check_scope(cell_val):
                val_str = str(cell_val).upper()
                if 'GLOBAL' in val_str: return True
                parts = [s.strip() for s in val_str.split(',')]
                return division_code in parts
            
            mask = df['SCOPE'].apply(check_scope)
            df_filtered = df[mask].copy()

            # Scoring: Specific Match > Global. Pick the most-specific row for
            # each target, then restore the selected rows to the target's original
            # rule-file position. Transform rules can depend on earlier rules, so
            # scope selection must not group overrides ahead of unrelated globals
            # or move an overridden target to the physical override row.
            df_filtered['_RULE_ORDER'] = range(len(df_filtered))
            df_filtered['SCOPE_SCORE'] = df_filtered['SCOPE'].apply(lambda x: 2 if division_code in str(x).upper() else 1)
            target_order = df_filtered.groupby('TARGET_FIELD')['_RULE_ORDER'].min()
            scoped_rules = df_filtered.sort_values(
                ['TARGET_FIELD', 'SCOPE_SCORE', '_RULE_ORDER'],
                ascending=[True, False, True],
                kind='stable'
            )

            final_rules = scoped_rules.drop_duplicates(subset=['TARGET_FIELD'], keep='first').copy()
            final_rules['_EFFECTIVE_ORDER'] = final_rules['TARGET_FIELD'].map(target_order)
            final_rules = final_rules.sort_values('_EFFECTIVE_ORDER', kind='stable')
            self.rules_raw = final_rules.drop(columns=['SCOPE_SCORE', '_RULE_ORDER', '_EFFECTIVE_ORDER'])
            
            if self.rules_raw.empty:
                self.rules_raw = pd.DataFrame(columns=['TARGET_FIELD', 'SOURCE_FIELD', 'RULE_TYPE', 'RULE_VALUE', 'SCOPE'])

            # Load Lookups
            # FIX: keep_default_na=False here too for safety
            xls = pd.ExcelFile(self.file_path)
            for sheet in xls.sheet_names:
                if sheet not in ['Rules', '_Audit_Log']:
                    df_lookup = pd.read_excel(xls, sheet_name=sheet, dtype=str, keep_default_na=False)
                    if len(df_lookup.columns) >= 2:
                        key_col = df_lookup.columns[0]
                        val_col = df_lookup.columns[1]
                        self.lookups[sheet] = dict(zip(df_lookup[key_col], df_lookup[val_col]))

            return self.rules_raw, self.lookups

        except Exception as e:
            print(f"Config Load Error: {e}")
            return pd.DataFrame(), {}

    def get_existing_targets(self):
        if os.path.exists(self.file_path):
            try:
                # FIX: keep_default_na=False
                df = pd.read_excel(self.file_path, sheet_name='Rules', keep_default_na=False)
                return df['TARGET_FIELD'].unique().tolist()
            except: return []
        return []