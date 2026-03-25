import pandas as pd
import os
import datetime
from colorama import Fore, Style
from modules.migration_runner import MigrationRunner


class BatchProcessor:
    def __init__(self):
        self.output_dir = 'output'
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def load_batch_file(self, batch_file_path):
        """
        Reads the batch Excel file and returns a clean DataFrame.
        """
        try:
            df_batch = pd.read_excel(batch_file_path)
            df_batch.columns = [str(c).strip().upper() for c in df_batch.columns]
            return df_batch
        except Exception as e:
            print(f"{Fore.RED}Critical Error reading batch file: {e}{Style.RESET_ALL}")
            return None

    def _resolve_path(self, user_path, default_folder='raw_data'):
        """
        Smartly tries to find a file.
        1. SQL query literal
        2. Exact path
        3. Filename under default folder
        """
        path_str = str(user_path).strip()

        if not path_str:
            return None

        if path_str.upper().startswith("SQL:"):
            return path_str

        if os.path.exists(path_str):
            return path_str

        name_only = os.path.basename(path_str)
        candidate = os.path.join(default_folder, name_only)
        if os.path.exists(candidate):
            return candidate

        return None

    def run_batch(self, batch_file_path):
        """
        Executes every enabled batch job by delegating to MigrationRunner.
        This guarantees standard/auto/surgical/batch all share the same coded-rule path.
        """
        df_batch = self.load_batch_file(batch_file_path)
        if df_batch is None:
            return

        print(f"\n{Fore.CYAN}--- BATCH PROCESSOR STARTED ---{Style.RESET_ALL}")
        print(f"Loaded {len(df_batch)} jobs from {os.path.basename(batch_file_path)}")

        runner = MigrationRunner()
        success_count = 0
        fail_count = 0
        log_report = []

        for idx, row in df_batch.iterrows():
            job_id = row.get('JOB_ID', idx + 1)
            enabled_val = str(row.get('ENABLED', 'Y')).strip().upper()
            if enabled_val in ['N', 'NO', '0', 'FALSE']:
                print(f"{Fore.YELLOW}Skipping Job {job_id} (ENABLED={enabled_val}).{Style.RESET_ALL}")
                log_report.append({'Job': job_id, 'Status': 'SKIPPED', 'Reason': 'Disabled'})
                continue

            print(f"\n{Fore.YELLOW}Processing Job {job_id}...{Style.RESET_ALL}")

            try:
                rule_config = str(row.get('RULE_CONFIG', '')).strip()
                if not rule_config:
                    raise ValueError("RULE_CONFIG is required")

                source_path = self._resolve_path(row.get('LEGACY_FILE', ''))
                if not source_path:
                    raise FileNotFoundError(f"Source file not found: {row.get('LEGACY_FILE', '')}")

                scope = str(row.get('SCOPE', 'GLOBAL')).strip().upper() or 'GLOBAL'
                target_sheets_raw = str(row.get('TARGET_SHEETS', '')).strip()
                target_sheets = [s.strip() for s in target_sheets_raw.split(',') if s.strip()] if target_sheets_raw else None

                out_prefix = str(row.get('OUTPUT_PREFIX', 'BATCH')).strip() or 'BATCH'
                out_filename = f"{out_prefix}_{rule_config}.xlsx"

                print(f"      Rules/API: {rule_config}")
                print(f"      Scope: {scope}")
                print(f"      Source: {str(source_path)[:80]}")

                runner.execute_migration(
                    program_name=rule_config,
                    legacy_path=source_path,
                    division=scope,
                    target_sheets=target_sheets,
                    silent=True,
                    output_name_override=out_filename,
                    append_if_exists=False
                )

                print(f"{Fore.GREEN}      [SUCCESS] Job {job_id} Finished.{Style.RESET_ALL}")
                success_count += 1
                log_report.append({'Job': job_id, 'Status': 'SUCCESS', 'File': out_filename})

            except Exception as e:
                print(f"{Fore.RED}      [FAILED] Job {job_id}: {e}{Style.RESET_ALL}")
                fail_count += 1
                log_report.append({'Job': job_id, 'Status': 'FAILED', 'Error': str(e)})

        print(f"\n{Fore.CYAN}--- BATCH RUN COMPLETE ---{Style.RESET_ALL}")
        print(f"Success: {Fore.GREEN}{success_count}{Style.RESET_ALL} | Failed: {Fore.RED}{fail_count}{Style.RESET_ALL}")

        # Disabled per request: do not auto-generate batch execution CSV logs in output/.
        # pd.DataFrame(log_report).to_csv(f"{self.output_dir}/batch_log_{datetime.date.today()}.csv", index=False)
