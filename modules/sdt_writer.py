import os
import re
import shutil

import pandas as pd
from colorama import Fore, Style
from openpyxl import load_workbook

from modules.transform_engine import TransformEngine


class SDTWriter:
    def __init__(self, output_dir='output'):
        self.output_dir = output_dir
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    @staticmethod
    def _canon_field_name(name):
        """Canonical field name used to match rule targets to template headers."""
        if name is None:
            return ""
        return re.sub(r'[^A-Z0-9]', '', str(name).strip().upper())

    def _transform_data(self, df_legacy, rules_df, valid_fields, sheet_name):
        """Applies rules to legacy data to create the target DataFrame."""
        if 'TARGET_FIELD' not in rules_df.columns:
            return pd.DataFrame(index=df_legacy.index)

        # Build a canonical map from template fields so target names like "SUNO"
        # can still match template headers like "SUNO#".
        valid_field_map = {
            self._canon_field_name(field): field
            for field in valid_fields
            if str(field).strip()
        }

        rules_copy = rules_df.copy()
        rules_copy['TARGET_FIELD_CANON'] = rules_copy['TARGET_FIELD'].apply(self._canon_field_name)

        relevant_rules = rules_copy[rules_copy['TARGET_FIELD_CANON'].isin(valid_field_map)].copy()
        if relevant_rules.empty:
            return pd.DataFrame(index=df_legacy.index)

        # Rewrite target fields to exact header names from the template.
        relevant_rules['TARGET_FIELD'] = relevant_rules['TARGET_FIELD_CANON'].map(valid_field_map)
        relevant_rules.drop(columns=['TARGET_FIELD_CANON'], inplace=True)

        engine = TransformEngine(relevant_rules, lookups={})
        return engine.process(df_legacy)

    @staticmethod
    def _norm_cell(x):
        """Normalize value before writing to Excel."""
        if pd.isna(x):
            return ""

        s = str(x).strip()

        # Convert integers serialized as float text (e.g. "123.0" -> "123")
        if s.endswith('.0'):
            try:
                s = str(int(float(s)))
            except Exception:
                pass

        # Prevent Excel from interpreting incoming text as formulas.
        # This avoids generated workbook repair errors for invalid formulas.
        if s.startswith(('=', '+', '-', '@')):
            return "'" + s

        return s

    def generate_from_template(
        self,
        template_path,
        legacy_data,
        rules_df,
        target_sheets,
        output_filename,
        append_if_exists=False
    ):
        """
        Creates a new SDT file based on the template and populated with data.
        If append_if_exists=True and file exists, it appends to the existing file instead of overwriting.
        """
        output_path = os.path.join(self.output_dir, output_filename)

        try:
            if append_if_exists and os.path.exists(output_path):
                print(f"   -> Appending to existing file: {output_filename}")
            else:
                shutil.copy(template_path, output_path)

            wb = load_workbook(output_path)

            for sheet_name in target_sheets:
                if sheet_name not in wb.sheetnames:
                    print(f"{Fore.YELLOW}   [Warning] Sheet '{sheet_name}' not found in template.{Style.RESET_ALL}")
                    continue

                ws = wb[sheet_name]

                headers = [cell.value for cell in ws[1]]
                valid_fields = [str(h).strip() for h in headers if h]

                df_out = self._transform_data(legacy_data, rules_df, valid_fields, sheet_name)

                df_final = pd.DataFrame(index=df_out.index)
                for field in valid_fields:
                    df_final[field] = df_out[field] if field in df_out.columns else ""

                for row_data in df_final[valid_fields].values.tolist():
                    ws.append([self._norm_cell(x) for x in row_data])

            wb.save(output_path)

            if not append_if_exists:
                print(f"   -> Generated: {output_path}")

        except Exception as e:
            print(f"{Fore.RED}   [SDT WRITER ERROR] {e}{Style.RESET_ALL}")
            raise
