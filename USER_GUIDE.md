
Generate updated User Guide
Apr 13
·
andresvicente-winpak/DataMigrationTool
·
main

Archive

Share

Create PR


hi chat, on this project I have a User Guide, but is outdated as we made changes on the code and add-ons . Could you generate an updated User Guide?
Copy
Summary
Added a new, updated USER_GUIDE.md that documents current platform capabilities, setup requirements, launch modes, and folder conventions for the project’s present architecture.

Documented the full GUI workflow across all hubs and tabs (Migration, Configuration, Rules/Admin, Utilities, and Sync/Merge), including newer add-ons like Script Library and offline rule synchronization/merge behavior.

Added practical operations content: output naming conventions, onboarding checklist, troubleshooting, and guidance on using the Markdown guide when the legacy PDF is outdated.

Testing

✅ git status --short

✅ git add USER_GUIDE.md && git commit -m "Add updated project user guide in Markdown"


File (1)
where should I copy this file?
Copy
Great question — for this project, put the file at the repo root:

✅ USER_GUIDE.md should live in /workspace/DataMigrationTool/USER_GUIDE.md (same level as main.py and gui.py).

Also important: the GUI “User Guide” button currently opens a PDF, not Markdown:

It explicitly looks for M3_Migration_User_Guide.pdf in the root directory.

So best practice is:

Keep USER_GUIDE.md at repo root for editable source docs.

If you want the in-app button to open updated docs, regenerate/update M3_Migration_User_Guide.pdf at repo root too.

(Checked with: nl -ba USER_GUIDE.md | sed -n '1,30p' and nl -ba modules/gui/app.py | sed -n '100,170p'.)




Diff
Logs

USER_GUIDE.md
USER_GUIDE.md
New
+306
-0

# M3 Data Migration Platform — User Guide (Updated)

Last updated: 2026-04-13

## 1) What this tool does

The **M3 Data Migration Platform** helps you transform legacy data (Excel/CSV/SQL sources) into M3 SDT/API-ready Excel outputs using configurable rule files.

It supports:
- Standard full-load migration by object/program mapping.
- Auto-detection of files based on signatures.
- Surgical (Load-by-ID) extraction workflows.
- Batch job execution from a manifest file.
- Rule management (edit/history/tools), including offline sync/merge.
- Utility tools for copying SDT sheets, merging files, and running helper scripts.

---

## 2) Prerequisites

- Python 3.10+
- Recommended dependencies:
  - `pandas`
  - `openpyxl`
  - `xlsxwriter`
  - `colorama`
  - `customtkinter`
  - `scikit-learn`
  - `fpdf2`

Example install:

```bash
pip install pandas openpyxl xlsxwriter colorama customtkinter scikit-learn fpdf2
```

---

## 3) Launching the application

### GUI (recommended)

```bash
python gui.py
```

This opens the desktop application with these hubs:
- Run Migration
- Configuration
- Rules & Admin
- Utilities
- Sync / Merge

### CLI mode (advanced/legacy)

```bash
python main.py
```

CLI mode still exposes many operations (migration, MCO import, snapshots, auto-detect, load by ID, batch, etc.), but the GUI should be preferred for daily usage.

---

## 4) Core folder layout (what matters most)

- `config/rules/` → Rule config workbooks (one per API).
- `config/objects_api.csv` → Main map tying object/program/API/source/template relationships.
- `config/source_map.csv` → Legacy source mapping fallback.
- `config/sdt_templates/` → Output template files used to build migration load files.
- `config/business_units.csv` → Scope values (division/business unit overrides).
- `config/surgical_def.csv` → Definitions used for Load-by-ID extraction.
- `config/db_config.ini` → SQL connection settings for `SQL:...` source entries.
- `output/` → Generated load files.
- `utilities/` → Script Library source for custom helper scripts.

---

## 5) Migration Hub

## 5.1 Standard tab (Full Load)

Use this for normal migrations where object/program mapping is already configured.

Steps:
1. Select **Object**.
2. Select **Program (MCO Sheet)**.
3. Select **Scope** (default `GLOBAL`).
4. Click **RUN MIGRATION**.

Behavior:
- Program/API/source/template details are resolved from mapping files.
- Rules are loaded by API + scope.
- Input is transformed and filtered by configured rules.
- Output is generated from SDT templates into `output/`.

---

## 5.2 Auto-Detect tab

Use when you have a Movex/legacy file and want the system to infer what it is.

Steps:
1. Select MCO specification file.
2. Let the detector learn signatures.
3. Select each source file to classify.
4. Confirm/adjust API if needed.
5. Run migration.

Best use case:
- Intake of mixed unknown files where manual program selection is slow.

---

## 5.3 Load by ID tab (Surgical)

Use this for **delta** or targeted loads (specific IDs only).

Steps:
1. Select business object.
2. Enter IDs (comma-separated).
3. Select scope.
4. Run extraction and migration task sequence.

Behavior:
- Uses `config/surgical_def.csv` to generate task list.
- Can run multiple related programs in one surgical sequence.
- Creates a consolidated output naming pattern such as:
  - `LOAD_<PROGRAM>_SURGICAL_<timestamp>.xlsx`

---

## 5.4 Batch tab

Use for pre-defined job manifests (multiple migration jobs in sequence).

Typical use:
- Scheduled or controlled migration windows.
- Repeatable cutover bundles.

Input:
- Batch workbook (e.g., from `batch_jobs/`) with job lines for API/program/source/scope parameters.

---

## 6) Configuration Hub

## 6.1 Import MCO

Purpose:
- Convert an MCO sheet into a rule workbook (`config/rules/<API>.xlsx`).

Workflow:
1. Browse/select MCO file (`.xlsx`/`.xlsm`).
2. Select worksheet.
3. Enter/confirm API name.
4. Choose overwrite mode:
   - **Unchecked** = merge with existing manual edits.
   - **Checked** = force overwrite.
5. Run import.

Also available:
- **Validate MCO Health** check for duplicate/missing logic issues.

---

## 6.2 Reverse Engineer

Purpose:
- Compare legacy file vs. “gold” M3 output and derive suggested rules.

Current status:
- UI is present, and **Merge Draft** workflow is active.
- **Analyze & Merge** action is currently a placeholder in this codebase.

Recommended usage now:
- Use external/generated draft workbook and apply via **MERGE DRAFT**.

---

## 6.3 Map Editor

Allows direct editing of key CSV maps:
- Objects API Map (merged)
- Source Map
- Surgical Def
- Business Units

Includes:
- Save/Reload
- Add row
- DB Connection popup for `config/db_config.ini`

Tip:
- Keep backup copies before major map edits.

---

## 7) Rules & Admin Hub

## 7.1 Editor

Interactive rules editing for selected API configs.

## 7.2 History

View audit history of rule changes.

## 7.3 Tools

Admin actions such as commit/maintenance/reset workflows.

---

## 8) Utilities Hub

## 8.1 Copy Sheet

Copies rows from one sheet to another inside the same workbook with header-based mapping.

Option:
- Include/Exclude rows marked as `NOK`.

## 8.2 Merge Files

Merges source sheet rows into a master workbook and deduplicates.

## 8.3 Script Library

Scans `utilities/<Category>/*.py` and lets users:
- View scripts by category.
- Edit script descriptions (saved in `utilities/scripts_meta.json`).
- Launch scripts in terminal.

Recommended script header:

```python
# DESCRIPTION: Short explanation of what this script does
```

---

## 9) Sync / Merge Hub (Offline Rule Merge)

Use when receiving rule files from another user/environment.

Workflow:
1. Pick local API rule file.
2. Pick incoming Excel file.
3. Run comparison.
4. Select desired changes.
5. Merge selected.

Behavior:
- Automatically snapshots local file before merge.
- Applies changed/new rules by target field + scope.
- Merges `_Audit_Log` entries and marks imported details.

---

## 10) Output files and naming

Default output folder: `output/`

Typical names:
- Standard: `LOAD_<API>_<YYYYMMDD>.xlsx` (+ version suffix if needed)
- Surgical: `LOAD_<PROGRAM>_SURGICAL_<YYYYMMDD_HHMMSS>.xlsx`

---

## 11) Recommended onboarding checklist

1. Confirm Python and dependencies.
2. Validate/update mapping CSVs in Config → Map Editor.
3. Import or refresh MCO rules for target APIs.
4. Verify SDT templates in `config/sdt_templates/`.
5. Run one Standard migration with a small sample file.
6. Review output workbook sheet structure and row counts.
7. Run a pilot batch only after single-file success.

---

## 12) Troubleshooting

### “No rules found” / missing APIs
- Confirm files exist in `config/rules/*.xlsx`.
- Re-import from MCO if needed.

### “No objects/sources found”
- Check `config/objects_api.csv` and `config/source_map.csv` headers/content.
- Ensure expected columns (e.g., `OBJECTS`, `MCO_SHEET`, `API_NAME`, source column).

### SQL source errors
- Validate `config/db_config.ini` in Map Editor DB popup.
- Test SQL server connectivity/driver availability.

### Merge/sync concerns
- Use Sync Hub compare first; merge only selected changes.
- If needed, recover from snapshot backups in rule snapshot area.

---

## 13) Notes on User Guide PDF

The GUI “User Guide” button opens `M3_Migration_User_Guide.pdf`.

If the PDF is outdated or missing, regenerate documentation assets via project scripts (`generate_docs.py` / packaging flow), or use this Markdown guide as the current operational reference.
