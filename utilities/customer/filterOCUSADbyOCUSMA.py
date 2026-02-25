import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def find_column(df, possible_names):
    """
    Searches the DataFrame for the first column name found in the possible_names list.
    Returns the found column name.
    """
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def main():
    # Hide the main Tkinter window
    root = tk.Tk()
    root.withdraw()
    
    # --- Step 1: Prompt for the "Customers" Excel file ---
    customer_file = filedialog.askopenfilename(
        title="Select the 'Customers' Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls *.csv")]
    )
    if not customer_file:
        messagebox.showinfo("Canceled", "No customer file selected. Exiting.")
        sys.exit(0)
    
    # --- Step 2: Prompt for the "Addresses" Excel file ---
    addresses_file = filedialog.askopenfilename(
        title="Select the 'Addresses' Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls *.csv")]
    )
    if not addresses_file:
        messagebox.showinfo("Canceled", "No addresses file selected. Exiting.")
        sys.exit(0)
    
    # Helper to read file (supports excel or csv based on extension)
    def read_file(path):
        if path.lower().endswith('.csv'):
            return pd.read_csv(path, dtype=str).fillna("")
        else:
            return pd.read_excel(path, dtype=str).fillna("")

    # Load each file into a Pandas DataFrame
    try:
        customers = read_file(customer_file)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read the customers file:\n{e}")
        sys.exit(1)
    
    try:
        addresses = read_file(addresses_file)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read the addresses file:\n{e}")
        sys.exit(1)
    
    # --- Step 3: Validate and Find Columns ---
    # Customer file usually has CUNO or OKCUNO
    cust_col = find_column(customers, ["OKCUNO", "CUNO"])
    if not cust_col:
        messagebox.showerror("Error", f"Could not find a Customer ID column in the customer file.\nLooked for: 'OKCUNO', 'CUNO'\nFound: {list(customers.columns)}")
        sys.exit(1)

    # Address file usually has OPCUNO or CUNO
    addr_col = find_column(addresses, ["OPCUNO", "CUNO"])
    if not addr_col:
        messagebox.showerror("Error", f"Could not find a Customer ID column in the addresses file.\nLooked for: 'OPCUNO', 'CUNO'\nFound: {list(addresses.columns)}")
        sys.exit(1)
    
    print(f"Using '{cust_col}' as key for Customers file.")
    print(f"Using '{addr_col}' as key for Addresses file.")

    # --- Step 4: Filter the Addresses rows ---
    # Logic: Normalize IDs to support both Movex (OKxxxx) and M3 (xxxx) formats
    # by stripping 'OK' from the start if it exists.
    
    def normalize_id(val):
        s = str(val).strip()
        if s.startswith("OK"):
            return s[2:]
        return s

    # Create a set of "normalized" valid IDs (e.g., "OK1234" becomes "1234")
    valid_ids = set(customers[cust_col].apply(normalize_id))
    
    # Check if the normalized address ID exists in the valid_ids set
    mask = addresses[addr_col].apply(normalize_id).isin(valid_ids)
    filtered_addresses = addresses[mask]
    
    # --- Step 5: Ask where to save the filtered Addresses table ---
    save_path = filedialog.asksaveasfilename(
        title="Save the filtered Addresses file",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not save_path:
        messagebox.showinfo("Canceled", "No save location selected. Exiting.")
        sys.exit(0)
    
    # --- Step 6: Save the filtered Addresses to Excel ---
    try:
        filtered_addresses.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"Filtered addresses saved to:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Could not save the filtered file:\n{e}")
        sys.exit(1)

if __name__ == "__main__":
    main()