import os
import json
import subprocess
import sys
import platform

class ScriptRunner:
    def __init__(self, root_dir="utilities"):
        self.root_dir = root_dir
        self.meta_file = os.path.join(self.root_dir, "scripts_meta.json")
        self.metadata = self._load_metadata()
        
        if not os.path.exists(self.root_dir):
            os.makedirs(self.root_dir)

    def _load_metadata(self):
        if os.path.exists(self.meta_file):
            try:
                with open(self.meta_file, 'r') as f:
                    return json.load(f)
            except: return {}
        return {}

    def _save_metadata(self):
        with open(self.meta_file, 'w') as f:
            json.dump(self.metadata, f, indent=4)

    def scan_scripts(self):
        """
        Returns a dict structure: { 'Category': [ {name, path, desc}, ... ] }
        """
        structure = {}
        if not os.path.exists(self.root_dir): return structure

        for category in os.listdir(self.root_dir):
            cat_path = os.path.join(self.root_dir, category)
            if os.path.isdir(cat_path):
                scripts = []
                for f in os.listdir(cat_path):
                    if f.endswith(".py") and f != "__init__.py":
                        full_path = os.path.join(cat_path, f)
                        script_id = f"{category}/{f}"
                        
                        # Get Description
                        desc = self.metadata.get(script_id, {}).get("description", "")
                        
                        # If no desc in JSON, check file header
                        if not desc:
                            desc = self._read_header_desc(full_path)
                        
                        scripts.append({
                            "id": script_id,
                            "name": f,
                            "path": full_path,
                            "description": desc
                        })
                
                if scripts:
                    structure[category] = scripts
        return structure

    def _read_header_desc(self, file_path):
        """Reads the first few lines to find '# DESCRIPTION: ...'"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for _ in range(5): # Check first 5 lines
                    line = f.readline()
                    if line.upper().startswith("# DESCRIPTION:"):
                        return line.split(':', 1)[1].strip()
        except: pass
        return "No description provided."

    def update_description(self, script_id, new_desc):
        if script_id not in self.metadata:
            self.metadata[script_id] = {}
        self.metadata[script_id]["description"] = new_desc
        self._save_metadata()

    def run_script_in_terminal(self, script_path):
        """
        Launches the script in a new independent OS terminal window.
        """
        abs_path = os.path.abspath(script_path)
        cmd = [sys.executable, abs_path]
        
        system = platform.system()
        
        try:
            if system == "Windows":
                # 'start' command, /wait is optional but we likely want async
                # 'cmd /k' keeps window open after script finishes
                subprocess.Popen(f'start cmd /k "{sys.executable} "{abs_path}""', shell=True)
            elif system == "Darwin": # macOS
                # AppleScript to open Terminal
                subprocess.Popen(['osascript', '-e', 
                    f'tell app "Terminal" to do script "{sys.executable} \'{abs_path}\'"'])
            elif system == "Linux":
                # Try common terminals
                terminals = ['gnome-terminal', 'xterm', 'konsole']
                for term in terminals:
                    if shutil.which(term):
                        if term == 'gnome-terminal':
                            subprocess.Popen([term, '--', sys.executable, abs_path])
                        else:
                            subprocess.Popen([term, '-e', f"{sys.executable} {abs_path}"])
                        break
        except Exception as e:
            return str(e)
        return None