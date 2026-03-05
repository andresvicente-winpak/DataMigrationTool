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

    def _normalize_map_part(self, value):
        return str(value).strip().upper()

    def _parse_map_config(self, config_str):
        """
        Supported formats:
        1) file|KEY|VAL
        2) file|KEY1,KEY2,...|VAL
        3) file|KEY1|KEY2|...|VAL   (legacy/alternate shorthand)
        """
        parts = [p.strip() for p in str(config_str).split('|') if str(p).strip()]

        if len(parts) < 3:
            print(f"{Fore.YELLOW}   [MAP WARNING] Invalid MAP format: '{config_str}'. Expected file|KEY|VAL or file|K1,K2|VAL.{Style.RESET_ALL}")
            return None, [], None

        path = parts[0].replace('\\', os.sep)
        val_col = parts[-1]

        if len(parts) == 3:
            key_cols = [k.strip() for k in parts[1].split(',') if k.strip()]
        else:
            # Alternate shorthand: file|K1|K2|...|VAL
            key_cols = [k.strip() for k in parts[1:-1] if k.strip()]

        if not path or not key_cols or not val_col:
            print(f"{Fore.YELLOW}   [MAP WARNING] Incomplete MAP config: '{config_str}'.{Style.RESET_ALL}")
            return None, [], None

        return path, key_cols, val_col

    def _resolve_source_col(self, field_name, src_map):
        f = str(field_name).strip().upper()
        if not f:
            return None
        if f in src_map:
            return src_map[f]

        match = next((c for c in src_map if c.endswith(f) and len(c) == 6), None)
        if match:
            return src_map[match]
        return None

    def _load_map_file(self, config_str):
        if config_str in self._map_cache:
            return self._map_cache[config_str]

        try:
            path, key_cols, val_col = self._parse_map_config(config_str)
            if not path or not key_cols or not val_col:
                return None

            original_path = path
            if not os.path.exists(path):
                if os.path.exists(os.path.join('config', path)):
                    path = os.path.join('config', path)
                elif os.path.exists(os.path.join('raw_data', path)):
                    path = os.path.join('raw_data', path)
                else:
                    print(f"{Fore.YELLOW}   [MAP WARNING] Map file not found. Tried: '{original_path}', 'config/{original_path}', 'raw_data/{original_path}'.{Style.RESET_ALL}")
                    return None

            print(f"{Fore.CYAN}   [MAP DEBUG] Loading map: {path} | keys={key_cols} | value={val_col}{Style.RESET_ALL}")

            if path.endswith('.csv'):
                df = pd.read_csv(path, dtype=str, keep_default_na=False)
            else:
                df = pd.read_excel(path, dtype=str, keep_default_na=False)

            df.columns = [str(c).strip() for c in df.columns]

            if val_col not in df.columns:
                print(f"{Fore.YELLOW}   [MAP WARNING] Value column '{val_col}' not found. Available: {list(df.columns)}{Style.RESET_ALL}")
                return None

            missing_keys = [col for col in key_cols if col not in df.columns]
            if missing_keys:
                print(f"{Fore.YELLOW}   [MAP WARNING] Key column(s) missing in map: {missing_keys}. Available: {list(df.columns)}{Style.RESET_ALL}")
                return None

            for col in key_cols:
                df[col] = df[col].astype(str).map(self._normalize_map_part)

            rows_before = len(df)
            df = df.drop_duplicates(subset=key_cols)
            rows_after = len(df)
            if rows_after < rows_before:
                print(f"{Fore.YELLOW}   [MAP DEBUG] Dropped {rows_before - rows_after} duplicate map row(s) by keys {key_cols}.{Style.RESET_ALL}")

            if len(key_cols) == 1:
                lookup_dict = pd.Series(df[val_col].values, index=df[key_cols[0]]).to_dict()
            else:
                key_series = df[key_cols].astype(str).apply(lambda r: '||'.join(r.values), axis=1)
                lookup_dict = pd.Series(df[val_col].values, index=key_series).to_dict()

            print(f"{Fore.CYAN}   [MAP DEBUG] Loaded {len(lookup_dict)} mapping key(s).{Style.RESET_ALL}")
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
            r_src = str(rule.get('SOURCE_FIELD', '')).strip().upper() if pd.notna(rule.get('SOURCE_FIELD')) else ""
            r_val_raw = rule.get('RULE_VALUE', '')

            if pd.isna(r_val_raw):
                r_val = ""
            else:
                r_val = str(r_val_raw).strip()

            if r_val.endswith('.0'):
                try:
                    r_val = str(int(float(r_val)))
                except Exception:
                    pass

            try:
                if rule_type == 'DIRECT':
                    if r_src in src_map:
                        df_target[target_col] = df_source[src_map[r_src]]
                    else:
                        match = next((c for c in src_map if c.endswith(r_src) and len(c) == 6), None)
                        if match:
                            df_target[target_col] = df_source[src_map[match]]

                elif rule_type == 'CONST':
                    df_target[target_col] = self._normalize_const_value(r_val)

                elif rule_type == 'MAP':
                    if not r_val:
                        print(f"{Fore.YELLOW}   [MAP WARNING] {target_col}: RULE_VALUE is blank.{Style.RESET_ALL}")
                        continue

                    map_path, key_cols, val_col = self._parse_map_config(r_val)
                    if not map_path or not key_cols:
                        print(f"{Fore.YELLOW}   [MAP WARNING] {target_col}: Could not parse map config '{r_val}'.{Style.RESET_ALL}")
                        continue

                    source_fields = [x.strip().upper() for x in r_src.split(',') if x.strip()]
                    if not source_fields:
                        # Enhancement: if SOURCE_FIELD blank, assume key columns are source column names.
                        source_fields = [k.strip().upper() for k in key_cols]

                    print(f"{Fore.CYAN}   [MAP DEBUG] {target_col}: RULE_VALUE='{r_val}' | parsed_path='{map_path}' | keys={key_cols} | value={val_col} | source_fields={source_fields}{Style.RESET_ALL}")

                    lookup_dict = self._load_map_file(r_val)
                    if not lookup_dict:
                        print(f"{Fore.YELLOW}   [MAP WARNING] {target_col}: Failed to build map lookup for '{r_val}'.{Style.RESET_ALL}")
                        continue

                    if len(key_cols) == 1 and len(source_fields) == 1:
                        col_name = self._resolve_source_col(source_fields[0], src_map)
                        if not col_name:
                            print(f"{Fore.YELLOW}   [MAP WARNING] {target_col}: Source column '{source_fields[0]}' not found. Available: {list(src_map.keys())}{Style.RESET_ALL}")
                            continue

                        normalized_source = df_source[col_name].astype(str).map(self._normalize_map_part)
                        mapped_series = normalized_source.map(lookup_dict)
                        miss_count = int(mapped_series.isna().sum())
                        if miss_count:
                            print(f"{Fore.YELLOW}   [MAP DEBUG] {target_col}: {miss_count}/{len(mapped_series)} row(s) had no map hit.{Style.RESET_ALL}")
                        df_target[target_col] = mapped_series
                    else:
                        if len(source_fields) != len(key_cols):
                            print(f"{Fore.YELLOW}   [MAP WARNING] {target_col}: SOURCE_FIELD count ({len(source_fields)}) must match map key count ({len(key_cols)}). SOURCE_FIELD={source_fields}, MAP_KEYS={key_cols}.{Style.RESET_ALL}")
                            continue

                        resolved_cols = [self._resolve_source_col(f, src_map) for f in source_fields]
                        if any(c is None for c in resolved_cols):
                            missing = [source_fields[i] for i, c in enumerate(resolved_cols) if c is None]
                            print(f"{Fore.YELLOW}   [MAP WARNING] {target_col}: Missing source column(s): {missing}. Available: {list(src_map.keys())}. Tip: set SOURCE_FIELD in same order as MAP keys {key_cols}.{Style.RESET_ALL}")
                            continue

                        print(f"{Fore.CYAN}   [MAP DEBUG] {target_col}: Resolved source columns={resolved_cols}{Style.RESET_ALL}")

                        composite_keys = df_source[resolved_cols].astype(str).apply(
                            lambda r: '||'.join([self._normalize_map_part(v) for v in r.values]), axis=1
                        )
                        mapped_series = composite_keys.map(lookup_dict)
                        miss_count = int(mapped_series.isna().sum())
                        if miss_count:
                            print(f"{Fore.YELLOW}   [MAP DEBUG] {target_col}: {miss_count}/{len(mapped_series)} row(s) had no composite map hit.{Style.RESET_ALL}")
                        df_target[target_col] = mapped_series

                elif rule_type == 'PYTHON':
                    col_name = None
                    if r_src in src_map:
                        col_name = src_map[r_src]
                    else:
                        match = next((c for c in src_map if c.endswith(r_src) and len(c) == 6), None)
                        if match:
                            col_name = src_map[match]

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
        if self.filter_rules.empty:
            return df_source

        df_filtered = df_source.copy()
        src_map = {c.upper(): c for c in df_source.columns}

        for idx, rule in self.filter_rules.iterrows():
            source_field = str(rule.get('SOURCE_FIELD', '')).strip().upper()
            condition_code = str(rule.get('RULE_VALUE', '')).strip()

            col_name = None
            if source_field in src_map:
                col_name = src_map[source_field]
            else:
                match = next((c for c in src_map if c.endswith(source_field) and len(c) == 6), None)
                if match:
                    col_name = src_map[match]

            def check_row(row):
                try:
                    source = row[col_name] if col_name else None
                    local_env = {'source': source, 'row': row}
                    return bool(eval(condition_code, {}, local_env))
                except Exception:
                    return False

            mask = df_filtered.apply(check_row, axis=1)
            df_filtered = df_filtered[mask]

        return df_filtered
