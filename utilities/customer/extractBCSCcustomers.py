import pandas as pd
import tkinter as tk
from tkinter import filedialog

def extract_unique_customers(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Check if required columns exist
    required_cols = ['OKCHAI', 'STATOKCUST', 'OKCUNM', 'OKCUA1']
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        print(f"Missing required columns: {', '.join(missing_cols)}")
        return

    # Extract unique customer IDs from OKCHAI and STATOKCUST
    unique_customers = pd.Series(
        list(df['OKCHAI'].dropna().unique()) + 
        list(df['STATOKCUST'].dropna().unique())
    ).drop_duplicates().reset_index(drop=True)

    # Create DataFrame with mapped OKCUNM and OKCUA1
    customers_df = pd.DataFrame({'OKCUST': unique_customers})

    # Ensure uniqueness before mapping
    okchai_map = df.drop_duplicates(subset='OKCHAI').set_index('OKCHAI')
    statokcust_map = df.drop_duplicates(subset='STATOKCUST').set_index('STATOKCUST')

    # Map customer name (OKCUNM) and customer address (OKCUA1)
    customers_df['OKCUNM'] = customers_df['OKCUST'].map(okchai_map['OKCUNM']).fillna(
                             customers_df['OKCUST'].map(statokcust_map['OKCUNM']))

    customers_df['OKCUA1'] = customers_df['OKCUST'].map(okchai_map['OKCUA1']).fillna(
                             customers_df['OKCUST'].map(statokcust_map['OKCUA1']))

    # Load the original Excel file to add the new sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        customers_df.to_excel(writer, sheet_name='customers', index=False)

    print(f"New sheet 'customers' with mapped values saved to {file_path}")

# Main execution with file selection dialog
if __name__ == "__main__":
    # Hide the main tkinter window
    root = tk.Tk()
    root.withdraw()
    
    # Prompt for the Excel file
    file_path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx")])
    
    if file_path:
        extract_unique_customers(file_path)
    else:
        print("No file selected.")
