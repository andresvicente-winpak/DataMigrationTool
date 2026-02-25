import pandas as pd
import os
import datetime

class RulePromoter:
    def __init__(self, output_dir='config/rules'):
        self.output_dir = output_dir
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def generate_production_rules(self, program_name, df_draft, df_legacy, df_gold, join_key_legacy, join_key_m3):
        print(f"Promoting Draft to Production for {program_name}...")
        
        production_rules = []
        lookup_sheets = {}
        
        # Handle list keys for merge
        key_leg = join_key_legacy if isinstance(join_key_legacy, list) else [join_key_legacy]
        key_m3 = join_key_m3 if isinstance(join_key_m3, list) else [join_key_m3]
        
        merged = pd.merge(df_legacy, df_gold, left_on=key_leg, right_on=key_m3, how='inner', suffixes=('_SRC', '_TGT'))
        
        for idx, row in df_draft.iterrows():
            target_col = row['TARGET']
            source_col = row['SOURCE']
            rule_type  = row['TYPE']
            logic_val  = row['LOGIC']
            
            rule_entry = {
                'TARGET_API': program_name,
                'TARGET_FIELD': target_col,
                'SOURCE_FIELD': '',
                'RULE_TYPE': rule_type,
                'RULE_VALUE': '',
                'SCOPE': 'GLOBAL',
                'DESCRIPTION': f"Auto-generated. Confidence: {row.get('CONFIDENCE', 'N/A')}"
            }

            if rule_type == 'CONST':
                rule_entry['RULE_VALUE'] = logic_val
            elif rule_type == 'DIRECT':
                rule_entry['SOURCE_FIELD'] = source_col
            elif rule_type == 'MAP':
                rule_entry['SOURCE_FIELD'] = source_col
                lookup_name = f"MAP_{source_col}_TO_{target_col}"[:30]
                rule_entry['RULE_VALUE'] = lookup_name
                
                try:
                    src_lookup_col = source_col if source_col in merged.columns else f"{source_col}_SRC"
                    tgt_lookup_col = target_col if target_col in merged.columns else f"{target_col}_TGT"
                    
                    lookup_df = merged[[src_lookup_col, tgt_lookup_col]].dropna().drop_duplicates(subset=[src_lookup_col])
                    lookup_df.columns = ['SOURCE_KEY', 'TARGET_VALUE']
                    lookup_sheets[lookup_name] = lookup_df.sort_values('SOURCE_KEY')
                except KeyError:
                    print(f"   Warning: Could not auto-generate lookup for {lookup_name}.")
            else:
                rule_entry['RULE_TYPE'] = 'TODO'
                rule_entry['DESCRIPTION'] = "Logic unknown. Requires manual intervention."

            production_rules.append(rule_entry)

        df_prod = pd.DataFrame(production_rules)
        file_path = f"{self.output_dir}/{program_name}.xlsx"
        
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df_prod.to_excel(writer, sheet_name='Rules', index=False)
            for sheet_name, df_table in lookup_sheets.items():
                df_table.to_excel(writer, sheet_name=sheet_name, index=False)
            pd.DataFrame(columns=['TIMESTAMP', 'USER', 'ACTION', 'TARGET_FIELD', 'OLD_VALUE', 'NEW_VALUE']).to_excel(writer, sheet_name='_Audit_Log', index=False)
            
        print(f"   -> Production Rule File Created: {file_path}")
        return file_path