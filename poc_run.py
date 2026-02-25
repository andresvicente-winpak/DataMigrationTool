import pandas as pd
from colorama import init, Fore, Style
from poc_loader import select_file, load_and_join
from poc_ai import PatternHunter

init(autoreset=True)

def main():
    print(f"\n{Fore.CYAN}=== M3 REVERSE ENGINEER POC ==={Style.RESET_ALL}")
    
    print("\n1. Select LEGACY Source File (Movex)...")
    legacy_path = select_file("Select Legacy File")
    if not legacy_path: return

    print("2. Select TARGET SDT File...")
    target_path = select_file("Select Target File")
    if not target_path: return

    try:
        # Load (Now returns sheet name)
        df, src_cols, tgt_cols, leg_sheet = load_and_join(legacy_path, target_path)
        print(f"\n{Fore.GREEN}   -> Loaded {len(df)} aligned rows.{Style.RESET_ALL}")
        
        # AI
        hunter = PatternHunter(df, src_cols, tgt_cols)
        
        print(f"\n{Fore.YELLOW}{'TARGET':<15} | {'PROB %':<8} | {'TYPE':<8} | {'LOGIC'}{Style.RESET_ALL}")
        print("-" * 100)

        results = []
        for col in tgt_cols:
            clean = col.replace('_TGT', '')
            if clean in ['CONO', 'DIVI', 'RGDT', 'LMDT', 'RGTM', 'CHID']: continue
            
            res = hunter.analyze_target(col)
            if not res: continue

            # Print
            color = Fore.WHITE
            if res['Prob'] == 100.0: color = Fore.GREEN
            elif res['Prob'] > 90.0: color = Fore.CYAN
            elif res['Prob'] < 50.0: color = Fore.LIGHTBLACK_EX
            
            print(f"{color}{clean:<15} | {res['Prob']:5.1f}%   | {res['Type']:<8} | {res['Logic']}{Style.RESET_ALL}")
            
            # Add Metadata for Excel Export
            res['SHEET'] = leg_sheet     # <--- Added Sheet Name
            res['FIELD_NAME'] = clean    # <--- Added Field Name
            
            # Rename internal keys to match your MCO standard if desired
            res['SOURCE_FIELD'] = res.get('Logic', '').replace('Copy ', '') if res['Type'] == 'DIRECT' else ''
            res['RULE_TYPE'] = res['Type']
            res['CONFIDENCE'] = f"{res['Prob']:.1f}%"
            
            results.append(res)

        # Reorder columns for readability
        final_df = pd.DataFrame(results)
        cols_order = ['SHEET', 'FIELD_NAME', 'RULE_TYPE', 'SOURCE_FIELD', 'Logic', 'CONFIDENCE']
        # Ensure cols exist before reordering
        final_cols = [c for c in cols_order if c in final_df.columns] + [c for c in final_df.columns if c not in cols_order]
        final_df = final_df[final_cols]

        # Save
        final_df.to_excel("Reverse_Engineer_Report.xlsx", index=False)
        print(f"\n{Fore.GREEN}Report saved to 'Reverse_Engineer_Report.xlsx'{Style.RESET_ALL}")

    except Exception as e:
        print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
        import traceback; traceback.print_exc()
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()