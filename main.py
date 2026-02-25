import pandas as pd
import os
import sys
import time
import tkinter as tk
import glob
from tkinter import filedialog
from colorama import init, Fore, Style
import datetime

# Initialize colorama
init(autoreset=True)

# --- IMPORTS ---
try:
    import modules.ui as ui
    from modules.audit_manager import AuditManager
    from modules.mco_importer import MCOImporter
    from modules.rule_manager import RuleManager
    from modules.batch_processor import BatchProcessor
    from modules.auto_detector import AutoDetector
    from modules.surgical_extractor import SurgicalExtractor
    from modules.migration_runner import MigrationRunner
    from modules.mco_checker import MCOChecker
    from modules.sdt_utils import SDTUtils
    from modules.extractor import DataExtractor
    from modules.validator_analyzer import ValidatorAnalyzer
    from modules.config_loader import ConfigLoader
except ImportError as e:
    print(f"{Fore.RED}CRITICAL ERROR: Could not import modules.{Style.RESET_ALL}")
    print(f"Details: {e}")
    sys.exit(1)

# --- ACTION CONTROLLERS ---

def action_manual_rule_entry():
    ui.print_header("Interactive Rule Editor")
    RuleManager().interactive_manual_entry()

def action_snapshot_manager():
    ui.print_header("Snapshot / Restore Manager")
    rule_dir = 'config/rules'
    files = glob.glob(os.path.join(rule_dir, "*.xlsx"))
    if not files: print(f"{Fore.RED}No rule files found.{Style.RESET_ALL}"); return
    
    names = [os.path.basename(f) for f in files]
    choice = ui.interactive_list_picker(names, "Select Rule File")
    if not choice: return
    
    program_name = choice.replace('.xlsx', '')
    auditor = AuditManager()
    
    print("\n1. Create Snapshot (Checkpoint)")
    print("2. Restore from Snapshot (Undo)")
    print("0. Cancel")
    opt = input("Select Action: ").strip()
    
    if opt == '1':
        note = input("Enter note: ").strip()
        auditor.create_snapshot(program_name, note if note else "Manual")
    elif opt == '2':
        snapshots = auditor.list_snapshots(program_name)
        if not snapshots: print("No snapshots found."); return
        restore_choice = ui.interactive_list_picker(snapshots, "Select Snapshot to Restore")
        if restore_choice and input(f"{Fore.RED}Confirm overwrite? (y/n): {Style.RESET_ALL}").lower() == 'y':
            auditor.restore_snapshot(program_name, restore_choice)

def action_import_mco_interactive():
    ui.print_header("Import MCO (Create Master Rule Set)")
    mco_path = ui.select_file("SELECT MCO SPECIFICATION EXCEL", [("Excel files", "*.xlsx *.xlsm")])
    if not mco_path: return
    MCOImporter().interactive_import(mco_path)

def action_commit_audit():
    ui.print_header("Commit External Excel Edits")
    rule_dir = 'config/rules'; ui.ensure_folder(rule_dir)
    files = glob.glob(os.path.join(rule_dir, "*.xlsx"))
    if not files: return
    names = [os.path.basename(f) for f in files]
    choice = ui.interactive_list_picker(names, "Select Rule File to Commit")
    if not choice: return
    try:
        auditor = AuditManager()
        auditor.commit_changes(choice.replace('.xlsx', ''))
        print(f"{Fore.GREEN}Done.{Style.RESET_ALL}")
    except Exception as e: print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")

def action_view_history():
    ui.print_header("View Rule History")
    rule_dir = 'config/rules'
    files = glob.glob(os.path.join(rule_dir, "*.xlsx"))
    if not files: print(f"{Fore.RED}No rule files found.{Style.RESET_ALL}"); return
    names = [os.path.basename(f) for f in files]
    choice = ui.interactive_list_picker(names, "Select Rule File")
    if not choice: return
    AuditManager().view_history(choice.replace('.xlsx', ''))
    input(f"\n{Fore.CYAN}Press Enter to return to menu...{Style.RESET_ALL}")

def action_migrate_context_aware():
    ui.print_header("Run Migration")
    rule_dir = 'config/rules'; files = glob.glob(os.path.join(rule_dir, "*.xlsx"))
    if not files: print(f"{Fore.RED}No rule files found.{Style.RESET_ALL}"); return
    names = [os.path.basename(f) for f in files]
    
    rule_choice = ui.interactive_list_picker(names, "STEP 1: Select Master Rule Configuration")
    if not rule_choice: return
    program_name = rule_choice.replace('.xlsx', '')

    legacy_path = ui.select_file(f"STEP 2: SELECT LEGACY SOURCE DATA", [("Excel files", "*.xlsx")])
    if not legacy_path: return
    
    if "MI" in os.path.basename(legacy_path).upper():
        if input(f"{Fore.RED}File looks like SDT. Sure? (y/n): {Style.RESET_ALL}").lower() != 'y': return

    division = input(f"\n{Fore.CYAN}>> Enter Scope (e.g. DIV_US, default GLOBAL): {Style.RESET_ALL}").strip().upper()
    if not division: division = 'GLOBAL'

    runner = MigrationRunner()
    runner.execute_migration(program_name, legacy_path, division=division, silent=False)

def action_auto_detect():
    ui.print_header("Auto-Detect File Load")
    print(f"{Fore.YELLOW}Note: We need the MCO file to learn signatures.{Style.RESET_ALL}")
    mco_path = ui.select_file("Select MCO Specification", [("Excel files", "*.xlsx *.xlsm")])
    if not mco_path: return

    detector = AutoDetector(mco_path)
    detector.learn_signatures()
    runner = MigrationRunner()

    while True:
        print(f"\n{Fore.CYAN}--- READY FOR NEXT FILE ---{Style.RESET_ALL}")
        legacy_path = ui.select_file("Select ANY Movex Data File", [("Excel files", "*.xlsx")])
        if not legacy_path: break

        prefix, mco_sheet, detected_api = detector.identify_file(legacy_path)
        if not prefix:
            print(f"{Fore.RED}Could not identify file signatures.{Style.RESET_ALL}")
            if input("Try another file? (y/n): ").lower() == 'y': continue
            else: break

        print(f"\n{Fore.GREEN}>>> DETECTION SUCCESSFUL!{Style.RESET_ALL}")
        print(f"    Prefix:      {prefix}")
        print(f"    MCO Sheet:   {mco_sheet}")
        
        map_api, map_sdt_file, _ = runner.resolve_from_map_public(mco_sheet, 'MCO_SHEET')
        final_api = map_api if map_api else detected_api
        final_sdt = None
        
        if map_sdt_file:
            full_sdt_path = os.path.join('config/sdt_templates', map_sdt_file)
            if os.path.exists(full_sdt_path): final_sdt = full_sdt_path

        print(f"    Target API:  {final_api}")
        if not final_api or final_api == "Unknown":
            final_api = input(f"{Fore.CYAN}>> Please confirm API Name: {Style.RESET_ALL}").strip().upper()
        
        rule_path = f"config/rules/{final_api}.xlsx"
        if not os.path.exists(rule_path):
            print(f"{Fore.YELLOW}Warning: Rule config {final_api}.xlsx not found. Skipping.{Style.RESET_ALL}")
            continue

        division = input(f"{Fore.CYAN}>> Enter Scope (default GLOBAL): {Style.RESET_ALL}").strip().upper()
        if not division:
            division = 'GLOBAL'

        runner.execute_migration(final_api, legacy_path, division=division, silent=True)
        print(f"\n{Fore.YELLOW}--------------------------------------------------{Style.RESET_ALL}")
        if input(f"{Fore.YELLOW}Process another Movex file? (y/n): {Style.RESET_ALL}").lower() != 'y': break

def action_load_by_id():
    ui.print_header("Load by ID (Delta Load)")
    extractor = SurgicalExtractor()
    objects = extractor.get_available_objects()
    if not objects:
        print(f"{Fore.RED}Configuration missing (surgical_def.csv).{Style.RESET_ALL}")
        return

    obj_type = ui.interactive_list_picker(objects, "Select Business Object")
    if not obj_type: return

    ids_str = input(f"\n{Fore.CYAN}>> Enter IDs (comma separated): {Style.RESET_ALL}").strip()
    if not ids_str: return
    id_list = [x.strip() for x in ids_str.split(',')]

    tasks = extractor.perform_extraction(obj_type, id_list)
    if not tasks:
        print(f"{Fore.YELLOW}No tasks generated. Check IDs or Source Files.{Style.RESET_ALL}"); return

    scope = input(f"\n{Fore.CYAN}>> Enter Scope (default GLOBAL): {Style.RESET_ALL}").strip().upper()
    if not scope: scope = 'GLOBAL'

    runner = MigrationRunner()
    print(f"\n{Fore.GREEN}Starting Execution of {len(tasks)} Surgical Tasks...{Style.RESET_ALL}")
    
    if tasks:
        base_prog = tasks[0]['program_name']
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        master_output_name = f"LOAD_{base_prog}_SURGICAL_{ts}.xlsx"
        print(f"Target Output File: {master_output_name}")
    else: master_output_name = None
    
    first_valid_run = True

    for i, task in enumerate(tasks):
        print(f"\n{Fore.YELLOW}>>> TASK {i+1}/{len(tasks)}: {task['program_name']} ({task['mco_sheet']}){Style.RESET_ALL}")
        _, _, trans_sheet_list = runner.resolve_from_map_public(task['mco_sheet'], 'MCO_SHEET')
        
        targets = None
        if isinstance(trans_sheet_list, list): targets = [x.strip() for x in trans_sheet_list if x]
        elif isinstance(trans_sheet_list, str): targets = [x.strip() for x in trans_sheet_list.split(',') if x.strip()]
        
        if targets:
            print(f"    Target Transaction(s): {targets}")
        else:
            print(f"{Fore.RED}    Warning: No TRANSACTION_SHEET defined. Skipping.{Style.RESET_ALL}")
            continue

        should_append = not first_valid_run

        runner.execute_migration(
            task['program_name'], 
            task['legacy_path'], 
            division=scope, 
            target_sheets=targets, 
            silent=True, 
            output_name_override=master_output_name,
            append_if_exists=should_append
        )

        first_valid_run = False

def action_batch_migration():
    ui.print_header("Run Batch Migration")
    batch_file = ui.select_file("Select Batch Job Definition Excel", [("Excel files", "*.xlsx")])
    if not batch_file: return
    processor = BatchProcessor()
    df_batch = processor.load_batch_file(batch_file)
    if df_batch is None: return
    print(f"\n{Fore.GREEN}Found {len(df_batch)} jobs.{Style.RESET_ALL}")
    print("1. Run All Enabled  2. Select Specific Jobs  0. Cancel")
    choice = input(f"\n{Fore.CYAN}Select Action: {Style.RESET_ALL}").strip()
    if choice == '2':
        options = [f"{row.get('JOB_ID', idx)} | {row.get('RULE_CONFIG')}" for idx, row in df_batch.iterrows()]
        selected_strings = ui.interactive_list_picker(options, "Select Jobs", multi=True)
        if selected_strings:
            indices = [i for i, opt in enumerate(options) if opt in selected_strings]
            print(f"{Fore.YELLOW}Selected-job execution is not supported in CLI mode; running full file instead.{Style.RESET_ALL}")
            processor.run_batch(batch_file)
    elif choice == '1':
        processor.run_batch(batch_file)

def action_check_mco():
    ui.print_header("MCO Health Checker")
    mco_path = ui.select_file("Select MCO File to Validate", [("Excel files", "*.xlsx *.xlsm")])
    if not mco_path: return
    checker = MCOChecker()
    checker.check_file(mco_path)
    input(f"\n{Fore.CYAN}Press Enter to return to menu...{Style.RESET_ALL}")

def action_utilities():
    ui.print_header("Utilities")
    print("1. Copy SDT Sheet Data")
    print("2. Merge SDT Files")
    print("0. Back")
    choice = input(f"\n{Fore.CYAN}>> Selection: {Style.RESET_ALL}").strip()
    if choice == '1': SDTUtils().copy_sdt_sheet_interactive()
    elif choice == '2': SDTUtils().merge_sdt_interactive()

def action_maintenance():
    ui.print_header("System Maintenance")
    print(f"{Fore.RED}1. Hard Reset History{Style.RESET_ALL}"); print("0. Cancel")
    if input(f"\n{Fore.CYAN}>> Selection: {Style.RESET_ALL}").strip() == '1':
        if input(f"{Fore.RED}Type 'DELETE': {Style.RESET_ALL}") == 'DELETE':
            AuditManager().hard_reset(); print(f"\n{Fore.GREEN}Reset Complete.{Style.RESET_ALL}")

def action_analyze_smart():
    ui.print_header("Incremental Analysis (Reverse Engineering)")
    
    # 1. Files
    legacy_path = ui.select_file("Select Legacy (Movex) File", [("Excel files", "*.xlsx")])
    if not legacy_path: return
    legacy_sheet = ui.get_sheet_selection(legacy_path)
    if not legacy_sheet: return
    
    gold_path = ui.select_file("Select Gold Standard (M3) File", [("Excel files", "*.xlsx")])
    if not gold_path: return
    
    # --- UPDATED LOGIC: Select Main Sheet + Merges ---
    print(f"\n{Fore.YELLOW}Select MAIN Sheet (Primary Key holder) in Gold File:{Style.RESET_ALL}")
    gold_sheet = ui.get_sheet_selection(gold_path)
    if not gold_sheet: return

    # Ask for extra sheets to stitch
    extra_sheets = []
    while True:
        ans = input(f"\nAdd another sheet to merge? (y/n): ").lower()
        if ans != 'y': break
        s = ui.get_sheet_selection(gold_path)
        if s and s != gold_sheet and s not in extra_sheets:
            extra_sheets.append(s)
            print(f"   Added {s}")

    try:
        extractor = DataExtractor()
        print("\nLoading files...")
        df_legacy = extractor.load_data(legacy_path, format_type='MOVEX', sheet_name=legacy_sheet)
        
        # STITCH GOLD DATA
        df_gold = extractor.load_sdt_stitched(gold_path, gold_sheet, extra_sheets)

        # Smart Key Defaulting
        key_l = df_legacy.columns[0]
        key_g = df_gold.columns[0]
        for c in df_legacy.columns: 
            if 'ITNO' in c or 'CUNO' in c: key_l = c; break
        for c in df_gold.columns: 
            if 'ITNO' in c or 'CUNO' in c: key_g = c; break
            
        print(f"Suggested Keys: {key_l} -> {key_g}")
        key_legacy = ui.select_columns_interactive(df_legacy, "Select Legacy Key:")
        key_m3 = ui.select_columns_interactive(df_gold, "Select M3 Key:")
        
        # Config Selection
        rule_dir = 'config/rules'; files = glob.glob(os.path.join(rule_dir, "*.xlsx"))
        names = [os.path.basename(f) for f in files]
        rule_choice = ui.interactive_list_picker(names, "Select Target Rule Config")
        if not rule_choice: return
        program_name = rule_choice.replace('.xlsx', '')
        
        ignore_exist = False
        if input("Re-evaluate ALL fields (Ignore existing rules)? (y/n): ").lower() == 'y':
            ignore_exist = True

        existing_targets = []
        if not ignore_exist:
             config = ConfigLoader(program_name)
             rules, _ = config.load_config()
             if not rules.empty: existing_targets = rules['TARGET_FIELD'].tolist()

        validator = ValidatorAnalyzer()
        suggestions = validator.reverse_engineer_rules(
            df_legacy, df_gold, key_legacy, key_m3, 
            existing_targets=existing_targets,
            ignore_existing_config=ignore_exist,
            legacy_sheet_name=legacy_sheet
        )
        
        if suggestions.empty:
            print(f"{Fore.GREEN}No new patterns found.{Style.RESET_ALL}"); return

        print(f"\n{Fore.GREEN}Found {len(suggestions)} patterns.{Style.RESET_ALL}")
        
        output_dir = "output"; ui.ensure_folder(output_dir)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = f"{output_dir}/{program_name}_DRAFT_{ts}.xlsx"
        suggestions.to_excel(out_path, index=False)
        print(f"Draft saved: {out_path}")
        
        print("\nOptions: 1. Merge Now  2. Exit")
        if input("Select: ") == '1':
            # Check for override
            force = input("Force Overwrite existing rules? (y/n): ").lower() == 'y'
            RuleManager().merge_draft_to_production(program_name, suggestions, overwrite=force)

    except Exception as e:
        print(f"{Fore.RED}Analysis failed: {e}{Style.RESET_ALL}")
        import traceback; traceback.print_exc()

# --- MAIN LOOP ---
def main_menu():
    while True:
        print("\n")
        print(f"{Fore.BLUE}=== M3 DATA MIGRATION PLATFORM ==={Style.RESET_ALL}")
        print("1. Interactive Rule Editor")
        print("2. Commit External Excel Edits")
        print("3. Import MCO (Master Init)")
        print("4. Run Migration (Interactive)")
        print("5. Run Batch Migration (Automated)")
        print("6. Auto-Detect File Load")
        print("7. Load by ID (Delta Load)")
        print("8. Snapshot / Restore Manager")
        print("9. View Rule History")
        print("10. Check MCO Health")
        print("11. Utilities")
        print("12. Incremental Analysis (Reverse Engineer)") # <--- Restored
        print("13. System Maintenance")
        print("0. Exit")
        
        choice = input(f"\n{Fore.CYAN}Select an option (0-13): {Style.RESET_ALL}")
        
        options = {
            '1': action_manual_rule_entry,
            '2': action_commit_audit,
            '3': action_import_mco_interactive,
            '4': action_migrate_context_aware,
            '5': action_batch_migration,
            '6': action_auto_detect,
            '7': action_load_by_id,
            '8': action_snapshot_manager,
            '9': action_view_history,
            '10': action_check_mco,
            '11': action_utilities,
            '12': action_analyze_smart, # <--- Mapped
            '13': action_maintenance
        }
        
        if choice in options: options[choice]()
        elif choice in ['0', 'q']: sys.exit(0)
        else: print(f"{Fore.RED}Invalid option.{Style.RESET_ALL}")

if __name__ == "__main__":
    root = tk.Tk(); root.withdraw()
    try: main_menu()
    except KeyboardInterrupt: print("\nOperation cancelled."); sys.exit(0)
