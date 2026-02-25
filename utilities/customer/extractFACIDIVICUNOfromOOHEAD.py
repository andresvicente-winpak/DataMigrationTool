import pandas as pd
from tkinter import Tk, filedialog, messagebox
from pathlib import Path

def pick_excel_file():
    """Open a file dialog to pick an Excel file and return the path, or None if cancelled."""
    root = Tk()
    root.withdraw()
    root.update()
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )
    root.destroy()
    return Path(file_path) if file_path else None

def main():
    # Pick input file
    input_path = pick_excel_file()
    if not input_path:
        print("No file selected. Exiting.")
        return

    try:
        # Load the only sheet in the workbook, treating everything as text
        xls = pd.ExcelFile(input_path)
        if len(xls.sheet_names) != 1:
            print(f"Warning: {input_path.name} has {len(xls.sheet_names)} sheets. Using the first: {xls.sheet_names[0]}")
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    required_cols = ["OACUNO", "OAFACI", "OADIVI"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        msg = f"Missing required columns in sheet '{sheet_name}': {', '.join(missing)}"
        print(msg)
        try:
            # Try to show a pop-up too
            root = Tk()
            root.withdraw()
            messagebox.showerror("Missing Columns", msg)
            root.destroy()
        except Exception:
            pass
        return

    # Get unique combinations of the three columns
    unique_df = (
        df[required_cols]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    # Build output file name in the same folder: <original>_unique_combos.xlsx
    output_path = input_path.with_name(input_path.stem + "_unique_combos.xlsx")

    try:
        unique_df.to_excel(output_path, index=False)
    except Exception as e:
        print(f"Error writing output Excel file: {e}")
        return

    print(f"Unique combinations saved to: {output_path}")

if __name__ == "__main__":
    main()
