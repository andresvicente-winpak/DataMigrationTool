from fpdf import FPDF
import datetime
import os

class PDF(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 15)
        self.cell(0, 10, 'M3 Data Migration Platform', 0, 1, 'C')
        self.set_font('Helvetica', 'I', 10)
        self.cell(0, 10, f'User Guide v2.1 - Generated: {datetime.date.today()}', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', 0, 0, 'C')

    def chapter_title(self, num, label):
        self.set_font('Helvetica', 'B', 12)
        self.set_fill_color(230, 230, 230)
        self.cell(0, 8, f'{num}. {label}', 0, 1, 'L', fill=True)
        self.ln(4)

    def chapter_body(self, body):
        self.set_font('Helvetica', '', 11)
        self.multi_cell(0, 6, body)
        self.ln()

    def add_bullet(self, text):
        self.set_font('Helvetica', '', 11)
        self.cell(5)
        self.cell(5, 6, chr(149), 0, 0)
        self.multi_cell(0, 6, text)
    
    def add_code_block(self, text):
        self.set_font('Courier', '', 10)
        self.set_fill_color(245, 245, 245)
        self.multi_cell(0, 5, text, fill=True)
        self.set_font('Helvetica', '', 11) # Reset
        self.ln()

def get_directory_structure(rootdir):
    """
    Creates a nested list of strings representing the folder structure.
    Ignores common junk folders.
    """
    ignore_dirs = ['__pycache__', '.git', '.history', '.snapshots', 'venv', 'output', 'tests', 'raw_data']
    structure_text = ""
    
    for root, dirs, files in os.walk(rootdir):
        # Filter directories in-place
        dirs[:] = [d for d in dirs if d not in ignore_dirs]
        
        level = root.replace(rootdir, '').count(os.sep)
        indent = ' ' * 4 * (level)
        subindent = ' ' * 4 * (level + 1)
        
        folder_name = os.path.basename(root)
        if folder_name == '.': folder_name = "ROOT"
        
        structure_text += f"{indent}{folder_name}/\n"
        
        for f in files:
            if f.endswith('.py') or f.endswith('.csv'):
                structure_text += f"{subindent}{f}\n"
                
    return structure_text

# --- GENERATION START ---

pdf = PDF()
pdf.alias_nb_pages()
pdf.add_page()

# 1. Overview
pdf.chapter_title(1, 'Overview')
pdf.chapter_body(
    "The M3 Data Migration Platform is a Python-based ETL (Extract, Transform, Load) tool designed to "
    "migrate legacy data (Movex, Excel, CSV) into Infor M3. It features a modern GUI, automated rule "
    "detection, and surgical data loading capabilities."
)

# 2. Installation & Requirements
pdf.chapter_title(2, 'Installation & Requirements')
pdf.chapter_body("Ensure you have Python 3.10+ installed. The following libraries are required:")
pdf.add_code_block(
    "pandas\n"
    "openpyxl\n"
    "xlsxwriter\n"
    "colorama\n"
    "customtkinter\n"
    "scikit-learn\n"
    "fpdf2"
)
pdf.chapter_body("To install all dependencies, run:")
pdf.add_code_block("pip install pandas openpyxl xlsxwriter colorama customtkinter scikit-learn fpdf2")

# 3. Directory Structure
pdf.chapter_title(3, 'System Architecture')
pdf.chapter_body("Below is the complete module structure of the application:")
structure = get_directory_structure(".")
pdf.add_code_block(structure)

# 4. Navigation
pdf.add_page() # Force new page for clean start
pdf.chapter_title(4, 'Navigation')
pdf.chapter_body("The application is divided into five main hubs, accessible via the sidebar:")
pdf.add_bullet("Run Migration: The primary interface for executing data loads.")
pdf.add_bullet("Configuration: Tools for importing MCOs and reverse-engineering data.")
pdf.add_bullet("Rules & Admin: Interactive editor for transformation rules.")
pdf.add_bullet("Utilities: Helper tools for Excel manipulation and external scripts.")
pdf.add_bullet("Sync / Merge: Offline synchronization of rules between users.")
pdf.ln()

# 5. Migration Hub
pdf.chapter_title(5, 'Migration Hub')
pdf.set_font('Helvetica', 'B', 11); pdf.cell(0, 6, "A. Standard Migration (Full Load)", 0, 1); pdf.ln(2)
pdf.chapter_body(
    "1. Select Rule Configuration: Choose the API config (e.g., MMS200MI).\n"
    "2. Select Legacy Source File: Browse for your source data (.xlsx).\n"
    "3. Scope: (Optional) Select a specific division (e.g., DIV_US).\n"
    "4. Click RUN MIGRATION."
)

pdf.set_font('Helvetica', 'B', 11); pdf.cell(0, 6, "B. Auto-Detect Migration", 0, 1); pdf.ln(2)
pdf.chapter_body(
    "1. Select Master MCO: Click Browse to pick the MCO file the system should learn from.\n"
    "2. Select Legacy File: Choose the file with unknown headers.\n"
    "3. Click DETECT & RUN."
)

pdf.set_font('Helvetica', 'B', 11); pdf.cell(0, 6, "C. Load by ID (Surgical)", 0, 1); pdf.ln(2)
pdf.chapter_body(
    "1. Select Business Object: Choose the object type (e.g., ITEM, CUSTOMER).\n"
    "2. Enter IDs: Paste a comma-separated list of IDs.\n"
    "3. Click RUN DELTA LOAD."
)

# 6. Configuration Hub
pdf.chapter_title(6, 'Configuration Hub')
pdf.add_bullet("Import MCO: Converts a functional MCO Excel file into a system Rule Configuration.")
pdf.add_bullet("Reverse Engineer: Uses AI to guess transformation rules by comparing a Legacy File against a 'Gold Standard' M3 file.")
pdf.add_bullet("Map Editor: Manage internal CSV mappings (Source files, API templates).")

# 7. Rules & Admin Hub
pdf.chapter_title(7, 'Rules & Admin Hub')
pdf.chapter_body(
    "The Editor allows you to modify rules interactively. The Left Panel lists all fields (supports filtering), "
    "and the Right Panel allows editing context (Global vs Division), overriding rules, or writing Python snippets."
)

# 8. Utilities Hub
pdf.chapter_title(8, 'Utilities Hub')
pdf.add_bullet("Copy Sheet: Copies data from one SDT file to another, automatically mapping columns.")
pdf.add_bullet("Merge Files: Deduplicates and merges rows from a Source file into a Master file.")
pdf.add_bullet("Script Library: A launcher for external Python scripts located in the 'utilities/' folder.")

# Output
output_filename = "M3_Migration_User_Guide.pdf"
pdf.output(output_filename)
print(f"[DOCS] Successfully generated: {output_filename}")