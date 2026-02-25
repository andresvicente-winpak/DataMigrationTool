import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def select_file(prompt):
    """Open a dialog box to select an Excel file (xlsx, xls, or xlsb)."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    print(f"Please select the {prompt} file...")
    file_path = filedialog.askopenfilename(
        title=f"Select {prompt} File",
        # Added *.xlsb to the allowed file types
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsb")]
    )
    if not file_path:
        print(f"No file selected for {prompt}. Exiting.")
        exit()
    return file_path

def load_excel_sheet(file_path, file_label):
    """
    Loads an Excel file. If multiple sheets exist, asks user to pick one.
    """
    try:
        # Check extension to determine engine
        # Pandas usually auto-detects, but explicit handling is safer for binary files
        engine = 'pyxlsb' if file_path.lower().endswith('.xlsb') else None
        
        xls = pd.ExcelFile(file_path, engine=engine)
        sheets = xls.sheet_names
        
        selected_sheet = sheets[0] # Default to first
        
        if len(sheets) > 1:
            print(f"\nMultiple sheets found in {file_label}:")
            for i, sheet in enumerate(sheets):
                print(f"{i + 1}: {sheet}")
            
            while True:
                try:
                    choice = int(input(f"Enter the number of the sheet to use for {file_label}: "))
                    if 1 <= choice <= len(sheets):
                        selected_sheet = sheets[choice - 1]
                        break
                    else:
                        print("Invalid number. Try again.")
                except ValueError:
                    print("Please enter a valid number.")
        
        print(f"Loading '{selected_sheet}' from {file_label}...")
        return pd.read_excel(file_path, sheet_name=selected_sheet, engine=engine)
        
    except Exception as e:
        print(f"Error reading {file_label}: {e}")
        exit()

def normalize_key_columns(df, label):
    """
    Standardizes the key column. 
    If MMITNO exists, use it. 
    If ITNO (M3 4-digit style) exists, rename it to MMITNO.
    """
    df.columns = [str(c).strip() for c in df.columns] # Clean whitespace and ensure string headers
    
    if 'MMITNO' in df.columns:
        print(f"Found 'MMITNO' in {label}.")
        return df
    elif 'ITNO' in df.columns:
        print(f"Found 'ITNO' in {label} (M3 Style). Renaming to 'MMITNO'.")
        df.rename(columns={'ITNO': 'MMITNO'}, inplace=True)
        return df
    else:
        print(f"WARNING: Could not find MMITNO or ITNO in {label}. Merge may fail.")
        return df

def main():
    # 1. Select Files
    mitmas_path = select_file("MITMAS (Base File)")
    mitmpr_path = select_file("MITMPR (Merge File)")

    # 2. Load Data (Handle Sheets)
    df_mas = load_excel_sheet(mitmas_path, "MITMAS")
    df_mpr = load_excel_sheet(mitmpr_path, "MITMPR")

    # 3. Normalize Keys (Handle 6-digit vs 4-digit)
    df_mas = normalize_key_columns(df_mas, "MITMAS")
    df_mpr = normalize_key_columns(df_mpr, "MITMPR")

    # Ensure key column is string to avoid mismatch (e.g. 100 vs "100")
    if 'MMITNO' in df_mas.columns:
        df_mas['MMITNO'] = df_mas['MMITNO'].astype(str).str.strip()
    if 'MMITNO' in df_mpr.columns:
        df_mpr['MMITNO'] = df_mpr['MMITNO'].astype(str).str.strip()

    # 4. Filter Columns to Merge
    # We want columns from MPR that are NOT in MAS, plus the key for joining.
    cols_to_add = [col for col in df_mpr.columns if col not in df_mas.columns]
    
    if 'MMITNO' not in df_mas.columns or 'MMITNO' not in df_mpr.columns:
        print("\nCRITICAL ERROR: 'MMITNO' key missing from one of the files. Cannot merge.")
        return

    # Add the key back to the list of columns to fetch
    cols_to_use_mpr = ['MMITNO'] + cols_to_add
    
    # Subset the MPR dataframe to only relevant columns
    df_mpr_clean = df_mpr[cols_to_use_mpr]

    print(f"\nMerging {len(cols_to_add)} new columns from MITMPR into MITMAS...")

    # 5. Merge
    # Left join ensures we keep all rows in MITMAS, only adding data where matches exist
    df_merged = pd.merge(df_mas, df_mpr_clean, on='MMITNO', how='left')

    # 6. Export
    output_filename = "Merged_MITMAS.xlsx"
    print(f"Saving to {output_filename}...")
    df_merged.to_excel(output_filename, index=False)
    print("Done!")

if __name__ == "__main__":
    main()