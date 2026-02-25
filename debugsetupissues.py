import os
import sys

print(f"Current Working Directory: {os.getcwd()}")
print("-" * 30)

# Check 1: Does 'modules' folder exist?
if os.path.isdir('modules'):
    print("[PASS] 'modules' folder found.")
else:
    print("[FAIL] 'modules' folder NOT found in this directory.")

# Check 2: Does '__init__.py' exist?
if os.path.exists('modules/__init__.py'):
    print("[PASS] 'modules/__init__.py' found.")
else:
    print("[FAIL] 'modules/__init__.py' is MISSING. Please create this empty file.")

# Check 3: Check for the specific file
if os.path.exists('modules/config_loader.py'):
    print("[PASS] 'modules/config_loader.py' found.")
else:
    print("[FAIL] 'modules/config_loader.py' is MISSING.")

# Check 4: Check for Libraries
try:
    import pandas
    print("[PASS] pandas library is installed.")
except ImportError:
    print("[FAIL] pandas is NOT installed. Run: pip install pandas")

print("-" * 30)
print("Attempting import now...")
try:
    from modules.config_loader import ConfigLoader
    print("[SUCCESS] Import worked!")
except ImportError as e:
    print(f"[CRITICAL FAILURE] Import failed. Reason: {e}")