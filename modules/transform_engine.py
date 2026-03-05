import pandas as pd
import numpy as np
import os
from colorama import Fore, Style

try:
    import modules.hooks as hooks
except ImportError:
    hooks = None

class TransformEngine:
    def __init__(self, rules_df, lookups):
        self.rules = rules_df
        self.lookups = lookups
        self._map_cache = {}

    def _normalize_const_value(self, value):
        """Normalize user-entered CONST markers that mean blank/empty string."""
        if value is None:
            return ""

        normalized = str(value).strip()
        if normalized in ['""', "''"]:
            return ""
        return normalized

    def _load_map_file(self, config_str):
        if config_str in self._map_cache:
            return self._map_cache[config_str]

        try:
            parts = config_str.split('|')
            if len(parts) != 3:
                return None
            
            path, key_col, val_col = parts[0].strip(), parts[1].strip(), parts[2].strip()
            
            if not os.path.exists(path):
                if os.path.exists(os.path.join('config', path)): path = os.path.join('config', path)
                elif os.path.exists(os.path.join('raw_data', path)): path = os.path.join('raw_data', path)
                else:
                    return None

            if path.endswith('.csv'): 
                df = pd.read_csv(path, dtype=str, keep_default_na=False)
            else: 
                df = pd.read_excel(path, dtype=str, keep_default_na=False)

            df.columns = [str(c).strip() for c in df.columns]
            
            if key_col not in df.columns or val_col not in df.columns:
                return None

            df[key_col] = df[key_col].astype(str).str.strip().str.upper()
            df = df.drop_duplicates(subset=[key_col])
            
            lookup_dict = pd.Series(df[val_col].values, index=df[key_col]).to_dict()
            self._map_cache[config_str] = lookup_dict
            return lookup_dict

        except Exception as e:
            print(f"{Fore.RED}   [MAP ERROR] Loading '{config_str}': {e}{Style.RESET_ALL}")
            return None

    def _execute_python_rule(self, code_snippet, source_val, row_data):
        def lookup_helper(filename, key_col, val_col, lookup_value):
            config_str = f"{filename}|{key_col}|{val_col}"
            mapping = self._load_map_file(config_str)
            if mapping:
                norm_key = str(lookup_value).strip().upper()
                return mapping.get(norm_key, None)
            return None

        try:
            local_vars = {'source': source_val, 'row': row_data, 'lookup': lookup_helper}
            func_def = f"def user_transform(source, row, lookup):\n"
            indented_code = "\n".join([f"    {line}" for line in code_snippet.split('\n')])
            exec(func_def + indented_code, {}, local_vars)
            return local_vars['user_transform'](source_val, row_data, lookup_helper)
        except Exception:
            return str(source_val)

    def process(self, df_source):
        df_target = pd.DataFrame(index=df_source.index)
        src_map = {c.upper(): c for c in df_source.columns}

        # Handle Missing Columns in Rules DataFrame gracefully
        required_cols = ['TARGET_FIELD', 'RULE_TYPE', 'RULE_VALUE', 'SOURCE_FIELD']
        for c in required_cols:
            if c not in self.rules.columns:
                self.rules[c] = ""

        if 'RULE_TYPE' in self.rules.columns:
            transform_rules = self.rules[self.rules['RULE_TYPE'] != 'FILTER']
        else:
            transform_rules = self.rules

        for index, rule in transform_rules.iterrows():
            target_col = rule['TARGET_FIELD']
            rule_type = rule['RULE_TYPE']
            # FIX: Use .get() or access checking
            r_src = str(rule.get('SOURCE_FIELD', '')).strip().upper() if pd.notna(rule.get('SOURCE_FIELD')) else ""
            r_val_raw = rule.get('RULE_VALUE', '')
            
            if pd.isna(r_val_raw): r_val = "" 
            else: r_val = str(r_val_raw).strip()

            if r_val.endswith('.0'):
                try: r_val = str(int(float(r_val)))
                except: pass

            try:
                if rule_type == 'DIRECT':
                    if r_src in src_map:
                        df_target[target_col] = df_source[src_map[r_src]]
                    else:
                        match = next((c for c in src_map if c.endswith(r_src) and len(c)==6), None)
                        if match: df_target[target_col] = df_source[src_map[match]]

                elif rule_type == 'CONST':
                    df_target[target_col] = self._normalize_const_value(r_val)

                elif rule_type == 'MAP':
                    source_series = None
                    if r_src in src_map: source_series = df_source[src_map[r_src]]
                    else:
                        match = next((c for c in src_map if c.endswith(r_src) and len(c)==6), None)
                        if match: source_series = df_source[src_map[match]]
                    
                    if source_series is not None and r_val:
                        # Default MAP behavior: keep original source value unless translation exists.
                        df_target[target_col] = source_series
                        lookup_dict = self._load_map_file(r_val)
                        if lookup_dict:
                            normalized_source = source_series.astype(str).str.strip().str.upper()
                            mapped_series = normalized_source.map(lookup_dict)
                            # Keep original source value when a key is missing in the translation map.
                            df_target[target_col] = mapped_series.where(mapped_series.notna(), source_series)
                            
                elif rule_type == 'PYTHON':
                    col_name = None
                    if r_src in src_map: col_name = src_map[r_src]
                    else:
                        match = next((c for c in src_map if c.endswith(r_src) and len(c)==6), None)
                        if match: col_name = src_map[match]
                    
                    if col_name:
                        df_target[target_col] = df_source.apply(lambda row: self._execute_python_rule(r_val, row[col_name], row), axis=1)
                    else:
                        df_target[target_col] = df_source.apply(lambda row: self._execute_python_rule(r_val, None, row), axis=1)

            except Exception as e:
                print(f"{Fore.RED}      [RULE ERROR] {target_col}: {e}{Style.RESET_ALL}")
        
        return df_target

class FilterEngine:
    def __init__(self, rules_df):
        if 'RULE_TYPE' in rules_df.columns:
            self.filter_rules = rules_df[rules_df['RULE_TYPE'] == 'FILTER'].copy()
        else:
            self.filter_rules = pd.DataFrame()

    def apply_filters(self, df_source):
        if self.filter_rules.empty: return df_source

        df_filtered = df_source.copy()
        src_map = {c.upper(): c for c in df_source.columns}

        for idx, rule in self.filter_rules.iterrows():
            source_field = str(rule.get('SOURCE_FIELD', '')).strip().upper()
            condition_code = str(rule.get('RULE_VALUE', '')).strip()
            
            col_name = None
            if source_field in src_map: col_name = src_map[source_field]
            else:
                match = next((c for c in src_map if c.endswith(source_field) and len(c)==6), None)
                if match: col_name = src_map[match]

            def check_row(row):
                try:
                    source = row[col_name] if col_name else None
                    local_env = {'source': source, 'row': row}
                    return bool(eval(condition_code, {}, local_env))
                except: return False

            mask = df_filtered.apply(check_row, axis=1)
            df_filtered = df_filtered[mask]
            
        return df_filtered
