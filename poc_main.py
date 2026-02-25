import pandas as pd
from colorama import init, Fore, Style
from data_loader import select_file, load_and_join
from pattern_hunter import PatternHunter

init(autoreset=True)

def main():
    print(f"\n{Fore.CYAN}=== M3 REVERSE ENGINEER POC (Analysis Only) ==={Style.RESET_ALL}")
    
    print("\n1. Select LEGACY Source File (Movex)...")
    legacy_path = select_file("Select Legacy File")
    if not legacy_path: return print("Cancelled.")

    print("2. Select TARGET File (M3 SDT)...")
    target_path = select_file("Select Target File")
    if not target_path: return print("Cancelled.")

    try:
        # Load
        df, leg_map = load_and_join(legacy_path, target_path)
        
        # Initialize AI
        hunter = PatternHunter(df)
        
        print(f"\n{Fore.YELLOW}{'TARGET':<15} | {'PROB %':<8} | {'TYPE':<8} | {'LOGIC'}{Style.RESET_ALL}")
        print("-" * 100)

        # Analyze each target column
        for col in hunter.tgt_cols:
            clean_name = col.replace('_TGT', '')
            
            # Skip join keys if they are just duplicates
            if clean_name in ['ITNO', 'CUNO', 'CONO']: continue

            res = hunter.analyze_target(col)
            
            # Formatting
            pct = res['Prob']
            color = Fore.WHITE
            if pct == 100.0: color = Fore.GREEN
            elif pct >= 90.0: color = Fore.CYAN
            elif pct < 50.0: color = Fore.LIGHTBLACK_EX
            
            print(f"{color}{clean_name:<15} | {pct:5.1f}%   | {res['Type']:<8} | {res['Logic']}{Style.RESET_ALL}")

    except Exception as e:
        print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
        import traceback; traceback.print_exc()
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()