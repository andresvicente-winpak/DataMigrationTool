import os
import subprocess
import sys

# Files to ignore
IGNORE = ['__pycache__', '.git', '.history', '.snapshots', 'output', 'raw_data', 'venv']

def generate_documentation():
    """Runs the PDF generation script if it exists."""
    doc_script = "generate_docs.py"
    if os.path.exists(doc_script):
        print(f"Found {doc_script}. Generating User Guide PDF...")
        try:
            subprocess.run([sys.executable, doc_script], check=True)
            print("[SUCCESS] PDF Generation complete.")
        except Exception as e:
            print(f"[WARNING] Failed to generate PDF: {e}")
            print("Skipping documentation step.")
    else:
        print(f"[INFO] {doc_script} not found. Skipping PDF generation.")

def pack_project():
    print("Starting Project Pack...")
    
    with open("FULL_PROJECT_CONTEXT.txt", "w", encoding="utf-8") as outfile:
        outfile.write("CURRENT PROJECT STATE:\n\n")
        
        # Walk through all files
        for root, dirs, files in os.walk("."):
            # Filter ignored directories
            dirs[:] = [d for d in dirs if d not in IGNORE]
            
            for file in files:
                if file.endswith(".py") or file.endswith(".csv"):
                    path = os.path.join(root, file)
                    outfile.write(f"--- START FILE: {path} ---\n")
                    try:
                        with open(path, "r", encoding="utf-8") as f:
                            outfile.write(f.read())
                    except: outfile.write("[Error reading file]")
                    outfile.write(f"\n--- END FILE: {path} ---\n\n")

    print("Done. Upload 'FULL_PROJECT_CONTEXT.txt' to the AI chat.")

if __name__ == "__main__":
    generate_documentation()
    pack_project()