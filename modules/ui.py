import os
import sys
import tkinter as tk
from tkinter import filedialog
from colorama import init, Fore, Style
import pandas as pd
import customtkinter as ctk # Needed for GUI components

# Initialize colorama
init(autoreset=True)

def print_header(title):
    print("\n" + "="*60); print(f"   {title.upper()}"); print("="*60)

def ensure_folder(path):
    if not os.path.exists(path): os.makedirs(path)

def select_file(prompt, filetypes, initialdir=None):
    print(f"\n{Fore.GREEN}--> ACTION REQUIRED: {prompt}{Style.RESET_ALL}")
    try:
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
        
        # FIX: Default to current working directory if not specified
        if not initialdir:
            initialdir = os.getcwd()
            
        f = filedialog.askopenfilename(
            title=prompt, 
            filetypes=filetypes,
            initialdir=initialdir
        )
        root.destroy()
        return f
    except:
        return input(f"Enter path for {prompt}: ").strip()

def select_folder(prompt):
    print(f"\n{Fore.GREEN}--> ACTION REQUIRED: {prompt}{Style.RESET_ALL}")
    try:
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
        f = filedialog.askdirectory(title=prompt, initialdir=os.getcwd())
        root.destroy()
        return f
    except:
        return input(f"Enter directory for {prompt}: ").strip()

def interactive_list_picker(options, title_prompt, multi=False):
    filtered_indices = list(range(len(options)))
    filter_text = ""
    while True:
        print(f"\n{Fore.CYAN}--- {title_prompt} ---{Style.RESET_ALL}")
        if filter_text: print(f"{Fore.YELLOW}[Filter: '{filter_text}'] (Type 'all' to clear){Style.RESET_ALL}")
        
        limit = 20
        count = 0
        display_map = {}
        for i in filtered_indices:
            count += 1
            if count > limit:
                print(f"   ... ({len(filtered_indices) - limit} more matches)")
                break
            print(f"   {count}. {options[i]}")
            display_map[count] = i
        
        if multi: print(f"\n{Fore.GREEN}Type numbers (e.g. 1, 3), 'all', or Filter Text.{Style.RESET_ALL}")
        else: print(f"\n{Fore.GREEN}Type a Number, Filter Text, or '0' to Cancel.{Style.RESET_ALL}")
            
        user_input = input(f"{Fore.CYAN}>> Selection: {Style.RESET_ALL}").strip()
        if not user_input: continue
        
        if user_input == '0': return None

        if multi and user_input.lower() == 'all':
            return [options[i] for i in filtered_indices]

        if multi and ',' in user_input:
            try:
                parts = [int(x.strip()) for x in user_input.split(',')]
                selected = [options[display_map[p]] for p in parts if p in display_map]
                if selected: return selected
            except ValueError: pass

        if user_input.isdigit():
            choice = int(user_input)
            if choice in display_map:
                res = options[display_map[choice]]
                return [res] if multi else res
        
        elif user_input.lower() in ['clear']:
            filter_text = ""; filtered_indices = list(range(len(options)))
        else:
            filter_text = user_input
            filtered_indices = [i for i, opt in enumerate(options) if filter_text.lower() in str(opt).lower()]
            if not filtered_indices: print("No matches.")

def get_sheet_selection(file_path):
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        sheets = xls.sheet_names
        if len(sheets) == 1:
            print(f"   -> Auto-selecting single sheet: '{sheets[0]}'")
            return sheets[0]
        return interactive_list_picker(sheets, f"Select Sheet from {os.path.basename(file_path)}")
    except Exception as e:
        print(f"{Fore.RED}   Warning: Could not read sheets. Defaulting to first.{Style.RESET_ALL}")
        return 0

def select_columns_interactive(df, prompt_text):
    columns = list(df.columns)
    display_limit = min(8, len(columns))
    print(f"\n{Fore.CYAN}   {prompt_text}{Style.RESET_ALL}")
    print(f"   Available Columns (Showing first {display_limit}):")
    for i in range(display_limit):
        print(f"   {i+1}. {columns[i]}")
    if len(columns) > 8: print(f"   ... (and {len(columns)-8} more)")
        
    while True:
        user_input = input(f"{Fore.CYAN}   >> Enter Column Number(s) (comma separated): {Style.RESET_ALL}").strip()
        if not user_input: continue
        selected_cols = []
        try:
            indices = [int(x.strip()) - 1 for x in user_input.split(',')]
            for idx in indices:
                if 0 <= idx < len(columns): selected_cols.append(columns[idx])
            if selected_cols: return selected_cols
        except ValueError: pass

# --- GUI MULTI-SELECT WIDGET ---
class MultiSelectPopup(ctk.CTkToplevel):
    def __init__(self, parent, options, current_selection, callback):
        super().__init__(parent)
        self.title("Select Options")
        self.geometry("300x400")
        self.callback = callback
        
        # Modal behavior
        self.transient(parent)
        self.grab_set()
        
        ctk.CTkLabel(self, text="Select Options:", font=("Arial", 14, "bold")).pack(pady=10)
        
        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.vars = {}
        
        # Clean current selection (handle string "A, B")
        if isinstance(current_selection, str):
            current_set = set([x.strip() for x in current_selection.split(',')])
        else:
            current_set = set(current_selection) if current_selection else set()

        for opt in options:
            is_checked = opt in current_set
            var = ctk.BooleanVar(value=is_checked)
            cb = ctk.CTkCheckBox(scroll, text=opt, variable=var)
            cb.pack(anchor="w", pady=2, padx=5)
            self.vars[opt] = var
            
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkButton(btn_frame, text="Confirm", fg_color="green", command=self._confirm).pack(side="right", fill="x", expand=True, padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", fg_color="gray", command=self.destroy).pack(side="left", fill="x", expand=True, padx=5)

    def _confirm(self):
        selected = [opt for opt, var in self.vars.items() if var.get()]
        self.callback(selected)
        self.destroy()