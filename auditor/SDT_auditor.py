import os
import glob
import re
import json
import sqlite3
import hashlib
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import webbrowser
from datetime import datetime

# ==========================================
# CONFIGURATION
# ==========================================
STATUS_COLS = ['MESSAGE', 'RESULT', 'STATUS', '_STATUS', 'RESULTMESSAGE']
KEY_COLUMNS = ['ITNO', 'CUNO', 'SUNO', 'ORNO', 'FACI', 'WHLO', 'ROWREF']
DB_NAME = "migration_history.db"

class MigrationAuditor:
    def __init__(self, target_dir):
        self.target_dir = target_dir
        self.db_path = os.path.join(target_dir, DB_NAME)
        self._init_db()

    def _init_db(self):
        """Create the lightweight database to store history."""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        # Track unique files processed
        c.execute('''CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY,
            filename TEXT,
            base_name TEXT, 
            version_tag TEXT,
            file_hash TEXT,
            run_date DATETIME,
            cono TEXT,
            total_rows INTEGER,
            success_count INTEGER,
            fail_count INTEGER,
            top_error TEXT
        )''')
        # Track detailed sheets within a file
        c.execute('''CREATE TABLE IF NOT EXISTS sheets (
            run_id INTEGER,
            sheet_name TEXT,
            executed BOOLEAN,
            ok_count INTEGER,
            nok_count INTEGER,
            row_data JSON,
            FOREIGN KEY(run_id) REFERENCES runs(id)
        )''')
        conn.commit()
        conn.close()

    def run(self):
        # 1. Scan folder for new/changed files
        excel_files = glob.glob(os.path.join(self.target_dir, "*.xlsx"))
        
        for f in excel_files:
            if "~$" in f: continue
            self._process_file_if_new(f)
            
        # 2. Retrieve history for Dashboard
        return self._get_dashboard_data()

    def _get_file_hash(self, filepath):
        """Generate MD5 hash to detect if file changed."""
        hasher = hashlib.md5()
        with open(filepath, 'rb') as f:
            buf = f.read()
            hasher.update(buf)
        return hasher.hexdigest()

    def _process_file_if_new(self, file_path):
        filename = os.path.basename(file_path)
        file_hash = self._get_file_hash(file_path)
        
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # Check if we already processed this exact file version
        c.execute("SELECT id FROM runs WHERE filename=? AND file_hash=?", (filename, file_hash))
        if c.fetchone():
            conn.close()
            return # Skip, already in DB
            
        print(f"Processing new/changed file: {filename}")
        
        # --- PARSE FILE ---
        # 1. Infer grouping (e.g. "Items_v1" -> Base: "Items")
        base_name = re.sub(r'[_ -]?v\d+.*', '', filename, flags=re.IGNORECASE)
        base_name = os.path.splitext(base_name)[0]
        
        # 2. Get Context (Log)
        log_path = self._find_log_file(file_path)
        context = self._parse_log(log_path)
        
        # 3. Analyze Excel
        run_summary = self._analyze_excel(file_path, context)
        
        # 4. Insert into DB
        c.execute('''INSERT INTO runs 
            (filename, base_name, version_tag, file_hash, run_date, cono, total_rows, success_count, fail_count, top_error)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (filename, base_name, "Latest", file_hash, datetime.now(), context['CONO'], 
             run_summary['Total'], run_summary['OK'], run_summary['NOK'], run_summary['Top_Error']))
        
        run_id = c.lastrowid
        
        for sheet in run_summary['Sheets']:
            c.execute('''INSERT INTO sheets (run_id, sheet_name, executed, ok_count, nok_count, row_data)
                VALUES (?, ?, ?, ?, ?, ?)''',
                (run_id, sheet['Name'], sheet['Executed'], sheet['OK'], sheet['NOK'], json.dumps(sheet['Rows'])))
                
        conn.commit()
        conn.close()

    def _analyze_excel(self, file_path, context):
        xls = pd.ExcelFile(file_path)
        summary = {'OK': 0, 'NOK': 0, 'Total': 0, 'Top_Error': 'None', 'Sheets': []}
        all_errors = []

        for sheet_name in xls.sheet_names:
            if "|L" in sheet_name: continue
            
            was_run = (sheet_name in context["Executed_Sheets"]) if context["Executed_Sheets"] else True
            
            # Read header row 0 (Line 1)
            df = pd.read_excel(xls, sheet_name=sheet_name, header=0, dtype=str)
            df.columns = [str(c).upper().strip().replace('\n','') for c in df.columns]
            
            status_col = next((c for c in STATUS_COLS if c in df.columns), None)
            if not status_col: continue
            
            # Data starts at index 2 (Row 4)
            df_data = df.iloc[2:].copy()
            df_data = df_data[df_data[status_col].notna()]
            
            sheet_ok = 0
            sheet_nok = 0
            row_details = []
            
            for idx, row in df_data.iterrows():
                msg = str(row[status_col])
                line_no = idx + 2
                
                is_success = re.match(r'^OK|^Success', msg, re.IGNORECASE)
                if is_success: sheet_ok += 1
                else: 
                    sheet_nok += 1
                    all_errors.append(msg)
                
                # Only store NOK rows + context to save DB space? 
                # User asked for full log. We store everything for now.
                row_details.append({
                    'Line': line_no,
                    'ID': self._get_row_id(row),
                    'Status': 'OK' if is_success else 'NOK',
                    'Message': msg
                })
            
            if was_run:
                summary['OK'] += sheet_ok
                summary['NOK'] += sheet_nok
                summary['Total'] += (sheet_ok + sheet_nok)

            summary['Sheets'].append({
                'Name': sheet_name,
                'Executed': was_run,
                'OK': sheet_ok,
                'NOK': sheet_nok,
                'Rows': row_details
            })

        if all_errors:
            top = pd.Series(all_errors).mode()[0]
            summary['Top_Error'] = top[:100] + "..." if len(top)>100 else top
            
        return summary

    def _get_row_id(self, row):
        for key in KEY_COLUMNS:
            if key in row and pd.notna(row[key]): return f"{key}: {row[key]}"
        return "Unknown"

    def _find_log_file(self, excel_path):
        directory = os.path.dirname(excel_path)
        basename = os.path.splitext(os.path.basename(excel_path))[0]
        # Remove version suffix like _v1, _v02 for log matching if log naming differs?
        # Usually logs match the excel filename exactly.
        candidates = glob.glob(os.path.join(directory, f"{basename}*.log"))
        return max(candidates, key=os.path.getmtime) if candidates else None

    def _parse_log(self, log_path):
        ctx = {"CONO": "Unknown", "Executed_Sheets": set()}
        if not log_path: return ctx
        try:
            with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                # Find CONO
                match = re.search(r'(?:CONO|Company)\s*[:=]\s*(\d+)', content)
                if match: ctx['CONO'] = match.group(1)
                # Find Sheets
                headers = re.findall(r'((?:API|FNC)_[A-Za-z0-9]+_[A-Za-z0-9]+)', content)
                ctx['Executed_Sheets'].update(headers)
        except: pass
        return ctx

    def _get_dashboard_data(self):
        """Query DB to build hierarchical data (Project -> Versions)."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        
        # Get all runs ordered by Base Name and Date
        c.execute("SELECT * FROM runs ORDER BY base_name, filename DESC")
        runs = c.fetchall()
        
        # Group by Base Name
        projects = {}
        for run in runs:
            base = run['base_name']
            if base not in projects:
                projects[base] = []
            
            # Fetch sheets for this run
            c.execute("SELECT * FROM sheets WHERE run_id=?", (run['id'],))
            sheets = []
            for s in c.fetchall():
                sheets.append({
                    'Name': s['sheet_name'],
                    'Executed': s['executed'],
                    'OK': s['ok_count'],
                    'NOK': s['nok_count'],
                    'Rows': json.loads(s['row_data'])
                })
                
            projects[base].append({
                'Filename': run['filename'],
                'Date': run['run_date'],
                'CONO': run['cono'],
                'Total': run['total_rows'],
                'Success': run['success_count'],
                'Failed': run['fail_count'],
                'Top_Error': run['top_error'],
                'Sheets': sheets
            })
            
        conn.close()
        return projects

# ==========================================
# UI GENERATOR (History Aware)
# ==========================================
def generate_html(projects, output_folder):
    json_data = json.dumps(projects, default=str)
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Migration History Log</title>
        <style>
            :root {{ --primary: #2c3e50; --accent: #3498db; --bg: #ecf0f1; }}
            body {{ font-family: 'Segoe UI', sans-serif; background: var(--bg); margin: 0; padding: 20px; }}
            .container {{ max-width: 1400px; margin: 0 auto; display: flex; gap: 20px; }}
            
            /* SIDEBAR (Project List) */
            .sidebar {{ width: 300px; background: white; border-radius: 8px; padding: 15px; height: 90vh; overflow-y: auto; }}
            .project-item {{ padding: 10px; border-bottom: 1px solid #eee; cursor: pointer; }}
            .project-item:hover {{ background: #f8f9fa; }}
            .project-item.active {{ background: #e8f4fd; border-left: 4px solid var(--accent); }}
            .project-title {{ font-weight: bold; font-size: 0.9em; }}
            .project-meta {{ font-size: 0.8em; color: #7f8c8d; display: flex; justify-content: space-between; }}
            
            /* MAIN CONTENT */
            .main {{ flex: 1; display: flex; flex-direction: column; gap: 20px; }}
            .card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
            
            /* VERSION TIMELINE */
            .timeline {{ display: flex; gap: 10px; overflow-x: auto; padding-bottom: 10px; }}
            .version-chip {{ 
                background: #fff; border: 1px solid #bdc3c7; padding: 8px 12px; border-radius: 20px; 
                cursor: pointer; min-width: 120px; position: relative;
            }}
            .version-chip:hover {{ border-color: var(--accent); }}
            .version-chip.active {{ background: var(--accent); color: white; border-color: var(--accent); }}
            .v-stats {{ font-size: 0.75em; display: block; margin-top: 4px; }}
            
            /* COMPARISON ALERTS */
            .alert-box {{ padding: 10px; background: #fff3cd; color: #856404; border-radius: 4px; margin-bottom: 15px; font-size: 0.9em; display: none; }}

            /* TABLES */
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.9em; }}
            th {{ text-align: left; background: #f8f9fa; padding: 8px; }}
            td {{ padding: 8px; border-bottom: 1px solid #eee; }}
            .ok {{ color: green; }} .nok {{ color: red; }}
            tr.nok-row {{ background: #fff5f5; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="sidebar" id="sidebar"></div>
            <div class="main">
                <div class="card" id="timelineCard" style="display:none">
                    <h2 id="projTitle" style="margin:0 0 15px 0">Select a Project</h2>
                    <div class="timeline" id="timeline"></div>
                    <div id="driftAlert" class="alert-box"></div>
                </div>
                
                <div class="card" id="detailCard" style="display:none; flex:1; overflow-y:auto">
                    <div style="display:flex; justify-content:space-between; align-items:center">
                         <h3 id="verTitle">Version Details</h3>
                         <label><input type="checkbox" id="showOk" onchange="renderGrid()"> Show OK Rows</label>
                    </div>
                    <div id="grid"></div>
                </div>
            </div>
        </div>

        <script>
            const projects = {json_data};
            let currentProj = null;
            let currentVerIdx = 0;

            function init() {{
                const sb = document.getElementById('sidebar');
                Object.keys(projects).forEach(name => {{
                    const runs = projects[name];
                    const latest = runs[0]; // runs are sorted desc
                    
                    const div = document.createElement('div');
                    div.className = 'project-item';
                    div.onclick = () => loadProject(name, div);
                    
                    let statusColor = latest.Failed > 0 ? 'red' : 'green';
                    div.innerHTML = `
                        <div class="project-title">${{name}}</div>
                        <div class="project-meta">
                            <span>Runs: ${{runs.length}}</span>
                            <span style="color:${{statusColor}}">${{latest.Failed}} Err</span>
                        </div>
                    `;
                    sb.appendChild(div);
                }});
            }}

            function loadProject(name, el) {{
                currentProj = projects[name];
                
                // Highlight Sidebar
                document.querySelectorAll('.project-item').forEach(d => d.classList.remove('active'));
                el.classList.add('active');
                
                // Setup Header
                document.getElementById('timelineCard').style.display = 'block';
                document.getElementById('projTitle').innerText = name;
                
                // Build Timeline (Reverse order to show v1 -> v2 -> v3)
                const tl = document.getElementById('timeline');
                tl.innerHTML = '';
                
                // We stored them DESC, so reverse for timeline
                const chronoRuns = [...currentProj].reverse();
                
                chronoRuns.forEach((run, idx) => {{
                    // Calculate drift from previous
                    let drift = 0;
                    if(idx > 0) drift = run.Total - chronoRuns[idx-1].Total;
                    
                    const chip = document.createElement('div');
                    chip.className = 'version-chip';
                    chip.innerHTML = `
                        <strong>${{run.Filename.substring(0,20)}}...</strong>
                        <span class="v-stats">
                            <span style="color:green">✔${{run.Success}}</span> 
                            <span style="color:red">✖${{run.Failed}}</span>
                            ${{drift !== 0 ? '<span style="color:orange; font-weight:bold">('+ (drift>0?'+':'') + drift +')</span>' : ''}}
                        </span>
                    `;
                    chip.onclick = () => loadVersion(run, chip);
                    tl.appendChild(chip);
                    
                    // Auto-select latest
                    if(idx === chronoRuns.length - 1) chip.click();
                }});
            }}

            function loadVersion(run, el) {{
                // Highlight Chip
                document.querySelectorAll('.version-chip').forEach(c => c.classList.remove('active'));
                el.classList.add('active');
                
                document.getElementById('detailCard').style.display = 'block';
                document.getElementById('verTitle').innerText = run.Filename;
                
                // Render Grid
                window.currentRunData = run;
                renderGrid();
            }}

            function renderGrid() {{
                const run = window.currentRunData;
                const showOk = document.getElementById('showOk').checked;
                const container = document.getElementById('grid');
                
                let html = '';
                
                run.Sheets.forEach(sheet => {{
                    if(!sheet.Executed && !showOk) return; // Skip ignored sheets unless showing all
                    
                    html += `<h4 style="margin-top:20px; border-bottom:1px solid #ddd">${{sheet.Name}}</h4>`;
                    html += `<table><thead><tr><th>Line</th><th>ID</th><th>Status</th><th>Message</th></tr></thead><tbody>`;
                    
                    sheet.Rows.forEach(row => {{
                        if(row.Status === 'OK' && !showOk) return;
                        
                        const bg = row.Status === 'OK' ? '' : 'nok-row';
                        html += `<tr class="${{bg}}">
                            <td>${{row.Line}}</td>
                            <td>${{row.ID}}</td>
                            <td class="${{row.Status.toLowerCase()}}"><strong>${{row.Status}}</strong></td>
                            <td>${{row.Message}}</td>
                        </tr>`;
                    }});
                    html += '</tbody></table>';
                }});
                
                container.innerHTML = html;
            }}
            
            init();
        </script>
    </body>
    </html>
    """
    
    output_path = os.path.join(output_folder, "Migration_History.html")
    with open(output_path, "w", encoding='utf-8') as f:
        f.write(html)
    return output_path

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    target = filedialog.askdirectory()
    if target:
        auditor = MigrationAuditor(target)
        data = auditor.run()
        path = generate_html(data, target)
        webbrowser.open(path)