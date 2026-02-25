import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def select_file(title):
    root = tk.Tk()
    root.withdraw()
    # Make dialog stay on top
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel/CSV", "*.xlsx *.csv")])
    root.destroy()
    return file_path

def normalize_header(h):
    # Strip whitespace and uppercase
    h = str(h).strip().upper()
    # Movex 6-char stripper (MMITNO -> ITNO)
    if len(h) == 6 and h[:2].isalpha():
        return h[2:]
    return h

def load_and_align_data(legacy_path, m3_path):
    print("\n[LOADER] Reading files...")
    
    # 1. Load Legacy
    if legacy_path.endswith('.csv'): df_leg = pd.read_csv(legacy_path, dtype=str)
    else: df_leg = pd.read_excel(legacy_path, dtype=str)
    
    # Normalize Legacy Headers (MMITNO -> ITNO) for easier matching
    # But keep a map to know the original name
    leg_map = {normalize_header(c): c for c in df_leg.columns}
    df_leg.columns = [normalize_header(c) for c in df_leg.columns]

    # 2. Load M3 (Handle SDT structure: Header Row 1, Data Row 4)
    # We assume standard Excel for now, but handle the SDT row skip if needed
    if m3_path.endswith('.csv'): 
        df_m3 = pd.read_csv(m3_path, dtype=str)
    else:
        # Peek to see if it looks like an SDT (Row 1 has headers, Row 2 empty/desc)
        df_peek = pd.read_excel(m3_path, nrows=5, dtype=str)
        if 'MESSAGE' in [str(c).upper() for c in df_peek.columns]:
            # SDT Format: Header row 0, Data starts row 2 (index 2, Excel row 4)
            df_m3 = pd.read_excel(m3_path, header=0, dtype=str)
            df_m3 = df_m3.iloc[2:].reset_index(drop=True)
        else:
            # Standard Flat File
            df_m3 = pd.read_excel(m3_path, dtype=str)

    df_m3.columns = [str(c).strip().upper() for c in df_m3.columns]

    # 3. Find Join Keys
    # Priority keys
    candidates = ['ITNO', 'CUNO', 'SUNO', 'CONO', 'ORDN']
    join_keys = [k for k in candidates if k in df_leg.columns and k in df_m3.columns]
    
    if not join_keys:
        # Fallback: First common column
        common = list(set(df_leg.columns) & set(df_m3.columns))
        if common: join_keys = [common[0]]
        else: raise ValueError("No common keys found between files.")

    print(f"[LOADER] Joining on keys: {join_keys}")
    
    # 4. Merge
    # Inner join to analyze only matching records
    df_combined = pd.merge(df_leg, df_m3, on=join_keys, how='inner', suffixes=('_SRC', '_TGT'))
    
    print(f"[LOADER] Aligned {len(df_combined)} rows.")
    return df_combined, leg_map