"""
SQL Server Database Compare - Customer Master

Customer Master flow:
1. Read source raw Customer Master from dbo.OCUSMA where OKSTAT = '20'.
2. Apply the migration Rule File to SOURCE only.
3. Read target Customer Master from dbo.OCUSMA.
4. Normalize target OCUSMA column names by removing the first 2 characters for comparison:
   OKCUNO -> CUNO, OKCUNM -> CUNM, etc.
5. Compare using CUNO as the primary key.

Install dependencies:
    pip install pyodbc pandas customtkinter openpyxl

Requires Microsoft ODBC Driver 17 or 18 for SQL Server.
"""

from __future__ import annotations

import json
import os
import re
import threading
from dataclasses import dataclass
from typing import Any, Iterable

import customtkinter as ctk
import pandas as pd
try:
    import pyodbc
except ImportError:  # SQL Server validation UI can be imported without DB drivers installed.
    pyodbc = None
from tkinter import filedialog, messagebox, ttk

from modules.transform_engine import FilterEngine, TransformEngine


APP_CONFIG_FILE = "db_compare_settings.json"


def app_folder() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def app_config_path() -> str:
    return os.path.join(app_folder(), APP_CONFIG_FILE)


def load_app_settings() -> dict[str, Any]:
    path = app_config_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {}


def save_app_settings(settings: dict[str, Any]) -> None:
    with open(app_config_path(), "w", encoding="utf-8") as file:
        json.dump(settings, file, indent=4)


BUSINESS_UNIT_FILTERS = {
    "All": "",
    "Intercompany Sales": "LEFT(OKCUNO, 3) = '000'",
    "ABI": "LEFT(OKCUNO, 1) = '6'",
    "WD": "(LEFT(OKCUNO, 1) = '1' OR LEFT(OKCUNO, 1) = '2')",
    "WEM": "(LEFT(OKCUNO, 2) = '58' OR LEFT(OKCUNO, 2) = '59')",
    "WFI": "LEFT(OKCUNO, 1) = '3'",
    "WHS 115": "TRY_CONVERT(int, OKCUNO) >= 50000 AND TRY_CONVERT(int, OKCUNO) < 55000",
    "WHS 116": "TRY_CONVERT(int, OKCUNO) >= 55000 AND TRY_CONVERT(int, OKCUNO) < 58000",
    "WPP": "LEFT(OKCUNO, 1) = '4'",
}


@dataclass
class SqlServerConfig:
    server: str
    database: str
    auth_type: str = "Windows Authentication"
    username: str = ""
    password: str = ""
    driver: str = "ODBC Driver 17 for SQL Server"
    trust_certificate: bool = True
    encrypt: bool = True

    def connection_string(self) -> str:
        trust = "yes" if self.trust_certificate else "no"
        encrypt = "yes" if self.encrypt else "no"

        base = (
            f"DRIVER={{{self.driver}}};"
            f"SERVER={self.server};"
            f"DATABASE={self.database};"
            f"Encrypt={encrypt};"
            f"TrustServerCertificate={trust};"
        )

        if self.auth_type == "Windows Authentication":
            return base + "Trusted_Connection=yes;"

        return base + f"UID={self.username};PWD={self.password};"


def get_connection(config: SqlServerConfig):
    if pyodbc is None:
        raise ImportError(
            "pyodbc is required for database validation. Install pyodbc and the Microsoft ODBC Driver for SQL Server."
        )
    return pyodbc.connect(config.connection_string(), timeout=15)


def list_tables(config: SqlServerConfig) -> list[str]:
    query = """
        SELECT TABLE_SCHEMA + '.' + TABLE_NAME AS table_name
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_SCHEMA, TABLE_NAME;
    """
    with get_connection(config) as conn:
        rows = conn.cursor().execute(query).fetchall()
    return [row.table_name for row in rows]


def list_target_companies(config: SqlServerConfig) -> list[str]:
    """Return available target company numbers from dbo.OCUSMA.OKCONO."""
    query = """
        SELECT DISTINCT OKCONO
        FROM dbo.OCUSMA
        WHERE OKCONO IS NOT NULL
        ORDER BY OKCONO;
    """
    with get_connection(config) as conn:
        rows = conn.cursor().execute(query).fetchall()

    companies = [str(row.OKCONO).strip() for row in rows if str(row.OKCONO).strip()]
    return ["All"] + companies


def sql_literal(value: str) -> str:
    """Safely quote a simple SQL string literal for generated filter SQL."""
    return str(value).replace("'", "''")


def read_customer_master(config: SqlServerConfig, business_unit: str = "All") -> pd.DataFrame:
    """Read source raw Customer Master from dbo.OCUSMA."""
    where_clauses = ["OKSTAT = '20'"]

    business_unit_filter = BUSINESS_UNIT_FILTERS.get(business_unit, "")
    if business_unit_filter:
        where_clauses.append(f"({business_unit_filter})")

    query = f"""
        SELECT *
        FROM dbo.OCUSMA
        WHERE {' AND '.join(where_clauses)};
    """
    with get_connection(config) as conn:
        return pd.read_sql(query, conn)


def read_target_customer_master(
    config: SqlServerConfig,
    business_unit: str = "All",
    company: str = "All",
) -> pd.DataFrame:
    """Read target Customer Master from dbo.OCUSMA, optionally filtered by OKCONO."""
    where_clauses = []

    company = str(company).strip()
    if company and company.upper() != "ALL":
        where_clauses.append(f"OKCONO = '{sql_literal(company)}'")

    business_unit_filter = BUSINESS_UNIT_FILTERS.get(business_unit, "")
    if business_unit_filter:
        where_clauses.append(f"({business_unit_filter})")

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""
    query = f"SELECT * FROM dbo.OCUSMA {where_sql};"

    with get_connection(config) as conn:
        return pd.read_sql(query, conn)


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [str(col).strip().upper() for col in cleaned.columns]

    for col in cleaned.columns:
        if cleaned[col].dtype == "object":
            cleaned[col] = cleaned[col].map(
                lambda value: value.strip() if isinstance(value, str) else value
            )

    return cleaned.where(pd.notna(cleaned), None)


def normalize_compare_value(value: Any) -> str:
    """Prevent false differences like '100' vs 100, 'ABC ' vs 'ABC', None vs ''."""
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass

    text = str(value).strip()

    if text == "":
        return ""

    # Treat numeric formats as equal: 100, 100.0, '100'.
    try:
        numeric_value = float(text)
        if numeric_value.is_integer():
            return str(int(numeric_value))
        return str(numeric_value)
    except ValueError:
        pass

    # Case-insensitive comparison for text.
    return text.upper()


def clean_rule_field_name(field_name: str) -> str:
    field = str(field_name).strip().upper()
    field = field.replace("[", "").replace("]", "")

    if "." in field:
        field = field.split(".")[-1]

    field = re.sub(r"[^A-Z0-9_]", "", field)
    return field


def load_rules(rule_file_path: str) -> pd.DataFrame:
    excel_file = pd.ExcelFile(rule_file_path)
    sheet_name = "Rules" if "Rules" in excel_file.sheet_names else excel_file.sheet_names[0]

    rules = pd.read_excel(rule_file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("")
    rules.columns = [str(col).strip().upper() for col in rules.columns]

    column_aliases = {
        "SOURCE": "SOURCE_FIELD",
        "SOURCE FIELD": "SOURCE_FIELD",
        "SRC_FIELD": "SOURCE_FIELD",
        "TARGET": "TARGET_FIELD",
        "TARGET FIELD": "TARGET_FIELD",
        "DESTINATION_FIELD": "TARGET_FIELD",
        "DEST_FIELD": "TARGET_FIELD",
        "TYPE": "RULE_TYPE",
        "RULE": "RULE_TYPE",
        "TRANSFORMATION": "RULE_TYPE",
        "VALUE": "RULE_VALUE",
        "LOGIC": "RULE_VALUE",
    }
    rules = rules.rename(columns={col: column_aliases.get(col, col) for col in rules.columns})

    for col in ["SOURCE_FIELD", "TARGET_FIELD", "RULE_TYPE", "RULE_VALUE"]:
        if col not in rules.columns:
            rules[col] = ""

    rules["SOURCE_FIELD"] = rules["SOURCE_FIELD"].map(clean_rule_field_name)
    rules["TARGET_FIELD"] = rules["TARGET_FIELD"].map(clean_rule_field_name)
    rules["RULE_TYPE"] = rules["RULE_TYPE"].str.strip().str.upper()

    rules = rules[rules["TARGET_FIELD"].str.strip() != ""]
    return rules


def load_rule_types(rule_file_path: str) -> list[str]:
    rules = load_rules(rule_file_path)
    rule_types = sorted(
        rule_type
        for rule_type in rules["RULE_TYPE"].dropna().astype(str).str.strip().str.upper().unique()
        if rule_type
    )
    return ["All"] + rule_types




def load_exception_columns(exceptions_file_path: str) -> set[str]:
    """
    Load comparison exception columns from Excel.

    Supported sheets: Exceptions, Columns, or the first sheet.
    Supported column names: COLUMN, COLUMN_NAME, FIELD, FIELD_NAME, TARGET_FIELD.

    Example Excel:
        COLUMN        ACTIVE
        CHID          Y
        LMDT          Y
        RGDT          N

    ACTIVE is optional. If present, only Y/YES/TRUE/1 rows are used.
    Column names are cleaned the same way as Rule File field names, so OKCHID and CHID
    both become comparable after target normalization. CUNO is never ignored.
    """
    if not exceptions_file_path:
        return set()

    excel_file = pd.ExcelFile(exceptions_file_path)
    sheet_name = (
        "Exceptions"
        if "Exceptions" in excel_file.sheet_names
        else "Columns"
        if "Columns" in excel_file.sheet_names
        else excel_file.sheet_names[0]
    )

    exceptions = pd.read_excel(exceptions_file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False).fillna("")
    exceptions.columns = [str(col).strip().upper() for col in exceptions.columns]

    column_aliases = {
        "COLUMN": "COLUMN_NAME",
        "FIELD": "COLUMN_NAME",
        "FIELD_NAME": "COLUMN_NAME",
        "TARGET": "COLUMN_NAME",
        "TARGET_FIELD": "COLUMN_NAME",
        "TARGET COLUMN": "COLUMN_NAME",
        "TARGET_COLUMN": "COLUMN_NAME",
    }
    exceptions = exceptions.rename(columns={col: column_aliases.get(col, col) for col in exceptions.columns})

    if "COLUMN_NAME" not in exceptions.columns:
        raise ValueError(
            "Exceptions file must contain one column named COLUMN, COLUMN_NAME, FIELD, or TARGET_FIELD."
        )

    if "ACTIVE" in exceptions.columns:
        active = exceptions["ACTIVE"].astype(str).str.strip().str.upper()
        exceptions = exceptions[active.isin(["", "Y", "YES", "TRUE", "1", "X"])]

    ignored_columns: set[str] = set()
    for value in exceptions["COLUMN_NAME"].dropna().astype(str):
        column = clean_rule_field_name(value)
        if not column:
            continue

        # Target columns are normalized from OKCUNO -> CUNO before comparison.
        if column.startswith("OK") and len(column) > 2:
            column = column[2:]

        if column != "CUNO":
            ignored_columns.add(column)

    return ignored_columns

def source_value(row: pd.Series, source_field: str) -> Any:
    """
    Rule file removes the first 2 source characters.
    Example: rule field CUNO matches source OCUSMA field OKCUNO.
    """
    field = clean_rule_field_name(source_field)
    if not field:
        return None

    candidates = [field]
    if not field.startswith("OK"):
        candidates.append(f"OK{field}")
    else:
        candidates.append(field[2:])

    for candidate in candidates:
        if candidate in row.index:
            return row.get(candidate)
    return None


def run_python_rule(rule_code: str, source: Any, row: pd.Series) -> Any:
    local_vars = {"source": source, "row": row.to_dict(), "result": None}

    function_code = "def _rule(source, row):\n"
    for line in str(rule_code).splitlines():
        function_code += f"    {line}\n"

    exec(function_code, {}, local_vars)
    return local_vars["_rule"](source, row.to_dict())


def apply_filter_rule(rule_code: str, source: Any, row: pd.Series) -> bool:
    local_vars = {
        "source": source,
        "row": row.to_dict(),
        "val": "" if source is None else str(source).strip(),
    }
    return bool(eval(str(rule_code), {}, local_vars))


def transform_source_with_rules(
    source_df: pd.DataFrame,
    rules: pd.DataFrame,
    selected_rule_type: str = "All",
) -> pd.DataFrame:
    """Apply the shared migration rule engine to source rows before validation.

    This intentionally reuses the same FilterEngine and TransformEngine used by
    migrations, so DIRECT, CONST, MAP, and PYTHON rules behave the same in the
    validation utility. MAP rules can therefore use the existing translation
    workbooks referenced by RULE_VALUE.
    """
    source = normalize_dataframe(source_df)
    working_rules = rules.copy()

    for col in ["TARGET_FIELD", "SOURCE_FIELD", "RULE_TYPE", "RULE_VALUE"]:
        if col not in working_rules.columns:
            working_rules[col] = ""

    working_rules["TARGET_FIELD"] = working_rules["TARGET_FIELD"].map(clean_rule_field_name)
    working_rules["SOURCE_FIELD"] = working_rules["SOURCE_FIELD"].map(clean_rule_field_name)
    working_rules["RULE_TYPE"] = working_rules["RULE_TYPE"].astype(str).str.strip().str.upper()

    selected_rule_type = selected_rule_type.strip().upper()
    if selected_rule_type and selected_rule_type != "ALL":
        keep_key = (
            (working_rules["TARGET_FIELD"].map(clean_rule_field_name) == "CUNO")
            | (working_rules["SOURCE_FIELD"].map(clean_rule_field_name).isin(["CUNO", "OKCUNO"]))
        )
        keep_type = working_rules["RULE_TYPE"] == selected_rule_type
        keep_filters = working_rules["RULE_TYPE"] == "FILTER"
        working_rules = working_rules[keep_type | keep_key | keep_filters].copy()

    filtered_source = FilterEngine(working_rules).apply_filters(source)
    transform_rules = working_rules[
        ~working_rules["RULE_TYPE"].isin(["IGNORE", "TODO", "FILTER"])
        & (working_rules["TARGET_FIELD"] != "_ROW_")
        & (working_rules["TARGET_FIELD"] != "")
    ].copy()

    transformed = TransformEngine(transform_rules, {}).process(filtered_source)

    if "CUNO" not in transformed.columns or transformed["CUNO"].replace("", pd.NA).isna().all():
        transformed["CUNO"] = filtered_source.apply(lambda row: source_value(row, "CUNO"), axis=1)

    return normalize_dataframe(transformed)


def prepare_target_for_rule_comparison(target_df: pd.DataFrame) -> pd.DataFrame:
    """
    Target is dbo.OCUSMA but the Rule File target names are M3 names.
    Normalize target column names only:
        OKCUNO -> CUNO
        OKCUNM -> CUNM
        OKSTAT -> STAT
    """
    target = normalize_dataframe(target_df)
    rename_map = {
        col: col[2:]
        for col in target.columns
        if col.startswith("OK") and len(col) > 2
    }
    return target.rename(columns=rename_map)


def compare_tables(
    source_df: pd.DataFrame,
    target_df: pd.DataFrame,
    primary_key: str = "CUNO",
    ignored_columns: Iterable[str] | None = None,
) -> pd.DataFrame:
    ignored = {col.upper() for col in (ignored_columns or [])}
    pk = primary_key.upper()

    left = normalize_dataframe(source_df)
    right = normalize_dataframe(target_df)

    if pk not in left.columns:
        raise ValueError(f"Primary key {pk} was not found in transformed source table.")
    if pk not in right.columns:
        raise ValueError(f"Primary key {pk} was not found in target table.")

    left = left.drop_duplicates(subset=[pk], keep="first").set_index(pk)
    right = right.drop_duplicates(subset=[pk], keep="first").set_index(pk)

    left_keys = set(left.index)
    right_keys = set(right.index)

    results: list[dict[str, Any]] = []

    for key in sorted(left_keys - right_keys):
        results.append({
            "Issue": "Missing in target",
            "Customer": key,
            "Column": "",
            "Source Value": "Exists",
            "Target Value": "Missing",
        })

    for key in sorted(right_keys - left_keys):
        results.append({
            "Issue": "Missing in source",
            "Customer": key,
            "Column": "",
            "Source Value": "Missing",
            "Target Value": "Exists",
        })

    common_columns = [
        col for col in sorted(set(left.columns) & set(right.columns))
        if col not in ignored and col != pk
    ]

    for key in sorted(left_keys & right_keys):
        for col in common_columns:
            source_raw = left.at[key, col]
            target_raw = right.at[key, col]

            source_value_normalized = normalize_compare_value(source_raw)
            target_value_normalized = normalize_compare_value(target_raw)

            if source_value_normalized != target_value_normalized:
                results.append({
                    "Issue": "Different value",
                    "Customer": key,
                    "Column": col,
                    "Source Value": source_raw,
                    "Target Value": target_raw,
                })

    return pd.DataFrame(
        results,
        columns=["Issue", "Customer", "Column", "Source Value", "Target Value"],
    )


def compare_rule_based_customer_master(
    source_df: pd.DataFrame,
    target_df: pd.DataFrame,
    rules: pd.DataFrame,
    primary_key: str = "CUNO",
    selected_rule_type: str = "All",
    ignored_columns: Iterable[str] | None = None,
) -> pd.DataFrame:
    transformed_source = transform_source_with_rules(source_df, rules, selected_rule_type)

                                                                                  
                                                                   
                            
                                            
                                                          
                                       
                                                             
                                                  
                       
             
             
                                                           

    normalized_target = prepare_target_for_rule_comparison(target_df)
    return compare_tables(
        transformed_source,
        normalized_target,
        primary_key=primary_key,
        ignored_columns=ignored_columns,
    )


class DatabaseCompareHub(ctk.CTkFrame):
    def __init__(self, master=None) -> None:
        super().__init__(master)

        self.results_df = pd.DataFrame()
        self.table_values: list[str] = []

        self.default_source = {
            "authentication": "Windows Authentication",
            "server": "w10bpw02",
            "database": "Movex_Replication",
        }
        self.default_target = {
            "authentication": "Windows Authentication",
            "server": "w10etfsql01",
            "database": "di_trn_staging",
        }
        self.settings = load_app_settings()
        self.default_rule_file = self.settings.get("rule_file_path", self._default_rule_file())
        self.default_exceptions_file = self.settings.get("exceptions_file_path", "")
        self.default_company = self.settings.get("target_company", "All")
        self.default_target_object = "dbo.OCUSMA"
        self.connections_visible = False

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.source_frame = self._server_frame("Source DB", 0)
        self.target_frame = self._server_frame("Target DB", 1)
        self.source_frame.grid_remove()
        self.target_frame.grid_remove()

        controls = ctk.CTkFrame(self)
        controls.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=8)
        controls.grid_columnconfigure(1, weight=1)
        controls.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(controls, text="Module").grid(row=0, column=0, sticky="e", padx=8, pady=8)
        self.module_dropdown = ctk.CTkComboBox(
            controls,
            values=["Customer Master"],
            command=lambda _: self._module_changed(),
        )
        self.module_dropdown.set("Customer Master")
        self.module_dropdown.grid(row=0, column=1, sticky="ew", padx=8, pady=8)

        ctk.CTkButton(
            controls,
            text="Show Connection Settings",
            command=self.toggle_connection_settings,
        ).grid(row=0, column=2, padx=8)

        ctk.CTkLabel(controls, text="Business Unit").grid(row=1, column=0, sticky="e", padx=8)
        self.business_unit_dropdown = ctk.CTkComboBox(
            controls,
            values=list(BUSINESS_UNIT_FILTERS.keys()),
            command=lambda _: self._business_unit_changed(),
        )
        self.business_unit_dropdown.set("All")
        self.business_unit_dropdown.grid(row=1, column=1, sticky="ew", padx=8)

        ctk.CTkLabel(controls, text="Company").grid(row=0, column=3, sticky="e", padx=8)
        self.company_dropdown = ctk.CTkComboBox(
            controls,
            values=[self.default_company] if self.default_company != "All" else ["All"],
            command=lambda _: self._company_changed(),
        )
        self.company_dropdown.set(self.default_company)
        self.company_dropdown.grid(row=0, column=4, sticky="ew", padx=8, pady=8)
        # Companies are loaded automatically at startup.

        ctk.CTkLabel(controls, text="Target Table").grid(row=1, column=2, sticky="e", padx=8)
        self.target_object_entry = ctk.CTkEntry(controls, width=180)
        self.target_object_entry.insert(0, self.default_target_object)
        self.target_object_entry.configure(state="disabled")
        self.target_object_entry.grid(row=1, column=3, sticky="ew", padx=8)

        ctk.CTkLabel(controls, text="Rule File").grid(row=1, column=4, sticky="e", padx=8)
        self.rule_file_entry = ctk.CTkEntry(controls, width=240)
        self.rule_file_entry.insert(0, self.default_rule_file)
        self.rule_file_entry.grid(row=1, column=5, sticky="ew", padx=8)
        ctk.CTkButton(controls, text="Browse", command=self.browse_rule_file).grid(row=1, column=6, padx=8)

        ctk.CTkLabel(controls, text="Rule Type").grid(row=2, column=0, sticky="e", padx=8)
        self.rule_type_dropdown = ctk.CTkComboBox(controls, values=["All"])
        self.rule_type_dropdown.set("All")
        self.rule_type_dropdown.grid(row=2, column=1, sticky="ew", padx=8, pady=8)
        # Rule types are loaded automatically when the rule file is selected and at startup.

        ctk.CTkLabel(controls, text="Exceptions File").grid(row=3, column=0, sticky="e", padx=8)
        self.exceptions_file_entry = ctk.CTkEntry(controls, width=240)
        self.exceptions_file_entry.insert(0, self.default_exceptions_file)
        self.exceptions_file_entry.grid(row=3, column=1, columnspan=3, sticky="ew", padx=8, pady=8)
        ctk.CTkButton(controls, text="Browse", command=self.browse_exceptions_file).grid(row=3, column=4, padx=8)
        ctk.CTkButton(controls, text="Clear Exceptions", command=self.clear_exceptions_file).grid(row=3, column=5, padx=8)

        # Source filter is fixed internally for Customer Master and hidden from the main screen.
        self.table_info_entry = ctk.CTkEntry(controls, width=500)
        self.table_info_entry.insert(0, "dbo.OCUSMA where OKSTAT = '20'")
        self.table_info_entry.configure(state="disabled")
                                                                                              

        ctk.CTkButton(
            controls,
            text="Compare Customer Master",
            command=self.compare_selected_module,
            fg_color="#D97706",
            hover_color="#B45309",
            text_color="white",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=2, column=6, padx=8)
        ctk.CTkButton(controls, text="Export Excel", command=self.export_results).grid(row=2, column=7, padx=8)

        self.status_label = ctk.CTkLabel(self, text="Ready")
        self.status_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=16, pady=4)

        self.summary_label = ctk.CTkLabel(self, text="")
        self.summary_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=16, pady=4)

        table_frame = ctk.CTkFrame(self)
        table_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=12, pady=8)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=y_scroll.set)

        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=x_scroll.set)

        self.load_rule_type_options(show_errors=False)
        self.after(600, self.auto_load_startup_options)

    def _server_frame(self, title: str, column: int) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self)
        frame.grid(row=0, column=column, sticky="ew", padx=12, pady=12)
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 6)
        )

        entries = {}

        ctk.CTkLabel(frame, text="Authentication").grid(row=1, column=0, sticky="e", padx=8, pady=4)
        auth_dropdown = ctk.CTkComboBox(
            frame,
            values=["Windows Authentication", "SQL Server Authentication"],
            command=lambda _: self._toggle_auth_fields(frame),
        )
        auth_dropdown.set("Windows Authentication")
        auth_dropdown.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        entries["authentication"] = auth_dropdown

        labels = ["Server", "Database", "Username", "Password"]
        for row, label in enumerate(labels, start=2):
            ctk.CTkLabel(frame, text=label).grid(row=row, column=0, sticky="e", padx=8, pady=4)
            entry = ctk.CTkEntry(frame)
            if label == "Password":
                entry.configure(show="*")
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
            entries[label.lower()] = entry

        entries["username"].configure(state="disabled", placeholder_text="Not required for Windows Authentication")
        entries["password"].configure(state="disabled", placeholder_text="Not required for Windows Authentication")

        frame.entries = entries  # type: ignore[attr-defined]

        defaults = self.default_source if title == "Source DB" else self.default_target
        entries["authentication"].set(defaults["authentication"])
        entries["server"].insert(0, defaults["server"])
        entries["database"].insert(0, defaults["database"])

        self._toggle_auth_fields(frame)
        return frame

    def _toggle_auth_fields(self, frame: ctk.CTkFrame) -> None:
        entries = frame.entries  # type: ignore[attr-defined]
        auth_type = entries["authentication"].get()

        if auth_type == "Windows Authentication":
            entries["username"].delete(0, "end")
            entries["password"].delete(0, "end")
            entries["username"].configure(state="disabled", placeholder_text="Not required for Windows Authentication")
            entries["password"].configure(state="disabled", placeholder_text="Not required for Windows Authentication")
        else:
            entries["username"].configure(state="normal", placeholder_text="")
            entries["password"].configure(state="normal", placeholder_text="")

    def _config_from_frame(self, frame: ctk.CTkFrame) -> SqlServerConfig:
        entries = frame.entries  # type: ignore[attr-defined]
        return SqlServerConfig(
            server=entries["server"].get().strip(),
            database=entries["database"].get().strip(),
            auth_type=entries["authentication"].get(),
            username=entries["username"].get().strip(),
            password=entries["password"].get(),
        )

    def toggle_connection_settings(self) -> None:
        self.connections_visible = not self.connections_visible
        if self.connections_visible:
            self.source_frame.grid()
            self.target_frame.grid()
        else:
            self.source_frame.grid_remove()
            self.target_frame.grid_remove()

    def load_company_options(self, show_errors: bool = True) -> None:
        try:
            target_config = self._config_from_frame(self.target_frame)
            self._set_status("Loading companies from target dbo.OCUSMA.OKCONO...")
            companies = list_target_companies(target_config)
            self.company_dropdown.configure(values=companies)

            current_value = self.company_dropdown.get().strip() or "All"
            self.company_dropdown.set(current_value if current_value in companies else "All")
            self._company_changed()
            self._set_status(f"Loaded {len(companies) - 1:,} companies from target database.")
        except Exception as exc:
            self.company_dropdown.configure(values=["All"])
            self.company_dropdown.set("All")
            if show_errors:
                messagebox.showerror("Company Load Error", str(exc))
            else:
                self._set_status("Company auto-load skipped. Check target connection settings if needed.")


    def auto_load_startup_options(self) -> None:
        """Load dropdown options automatically when the app opens."""
        self.load_rule_type_options(show_errors=False)
        self.load_company_options(show_errors=False)

    def _company_changed(self) -> None:
        self.settings["target_company"] = self.company_dropdown.get().strip() or "All"
        save_app_settings(self.settings)

    def load_rule_type_options(self, show_errors: bool = True) -> None:
        try:
            rule_file_path = self._rule_file_path()
            rule_types = load_rule_types(rule_file_path)
            self.rule_type_dropdown.configure(values=rule_types)
            current_value = self.rule_type_dropdown.get()
            self.rule_type_dropdown.set(current_value if current_value in rule_types else "All")
            self._set_status(f"Loaded rule types from {os.path.basename(rule_file_path)}.")
        except Exception as exc:
            self.rule_type_dropdown.configure(values=["All"])
            self.rule_type_dropdown.set("All")
            if show_errors:
                messagebox.showerror("Rule Type Error", str(exc))

    def browse_rule_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select rule file",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        self.rule_file_entry.delete(0, "end")
        self.rule_file_entry.insert(0, file_path)

        self.settings["rule_file_path"] = file_path
        save_app_settings(self.settings)
        self.load_rule_type_options(show_errors=True)

    def browse_exceptions_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select comparison exceptions file",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not file_path:
            return

        self.exceptions_file_entry.delete(0, "end")
        self.exceptions_file_entry.insert(0, file_path)

        self.settings["exceptions_file_path"] = file_path
        save_app_settings(self.settings)

    def clear_exceptions_file(self) -> None:
        self.exceptions_file_entry.delete(0, "end")
        self.settings["exceptions_file_path"] = ""
        save_app_settings(self.settings)
        self._set_status("Exceptions file cleared. All common columns will be compared.")

    def _exceptions_file_path(self) -> str:
        exceptions_file = self.exceptions_file_entry.get().strip()
        if not exceptions_file:
            return ""

        if os.path.exists(exceptions_file):
            return exceptions_file

        script_folder = os.path.dirname(os.path.abspath(__file__))
        local_path = os.path.join(script_folder, exceptions_file)
        if os.path.exists(local_path):
            return local_path

        raise FileNotFoundError(
            f"Exceptions file not found: {exceptions_file}. Put it beside this script, use Browse, or clear it."
        )

    def _default_rule_file(self) -> str:
        default_path = os.path.join("config", "rules", "CRS610MI.xlsx")
        if os.path.exists(default_path):
            return default_path

        rule_dir = os.path.join("config", "rules")
        if os.path.isdir(rule_dir):
            candidates = sorted(
                os.path.join(rule_dir, name)
                for name in os.listdir(rule_dir)
                if name.lower().endswith(".xlsx") and not name.startswith("~$")
            )
            if candidates:
                return candidates[0]

        return "CRS610MI.xlsx"

    def _rule_file_path(self) -> str:
        rule_file = self.rule_file_entry.get().strip()
        if os.path.exists(rule_file):
            return rule_file

        search_paths = [
            os.path.join("config", "rules", rule_file),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), rule_file),
        ]
        for path in search_paths:
            if os.path.exists(path):
                return path

        raise FileNotFoundError(
            f"Rule file not found: {rule_file}. Use Browse or place it in config/rules."
        )

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)
        self.update_idletasks()

    def _module_changed(self) -> None:
        selected_module = self.module_dropdown.get()
        if selected_module == "Customer Master":
            self.business_unit_dropdown.set("All")
            self._business_unit_changed()

    def _business_unit_changed(self) -> None:
        business_unit = self.business_unit_dropdown.get()
        business_unit_filter = BUSINESS_UNIT_FILTERS.get(business_unit, "")
        filter_text = "dbo.OCUSMA where OKSTAT = '20'"

        if business_unit_filter:
            filter_text += f" and {business_unit_filter}"

        self.table_info_entry.configure(state="normal")
        self.table_info_entry.delete(0, "end")
        self.table_info_entry.insert(0, filter_text)
        self.table_info_entry.configure(state="disabled")

    def compare_selected_module(self) -> None:
        def worker() -> None:
            try:
                selected_module = self.module_dropdown.get().strip()
                business_unit = self.business_unit_dropdown.get().strip()
                selected_company = self.company_dropdown.get().strip() or "All"
                selected_rule_type = self.rule_type_dropdown.get().strip() or "All"

                source_config = self._config_from_frame(self.source_frame)
                target_config = self._config_from_frame(self.target_frame)

                if selected_module == "Customer Master":
                    rule_file_path = self._rule_file_path()

                    self._set_status(f"Loading rules from {os.path.basename(rule_file_path)}...")
                    rules = load_rules(rule_file_path)

                    exceptions_file_path = self._exceptions_file_path()
                    ignored_columns = set()
                    if exceptions_file_path:
                        self._set_status(f"Loading exceptions from {os.path.basename(exceptions_file_path)}...")
                        ignored_columns = load_exception_columns(exceptions_file_path)

                    self._set_status(f"Reading source Customer Master for Business Unit: {business_unit}...")
                    source_df = read_customer_master(source_config, business_unit)

                    self._set_status(
                        f"Reading target Customer Master for Business Unit: {business_unit}, Company: {selected_company}..."
                    )
                    target_df = read_target_customer_master(target_config, business_unit, selected_company)
                else:
                    raise ValueError(f"Unknown module: {selected_module}")

                self._set_status("Applying Rule File to source and comparing data...")
                self.results_df = compare_rule_based_customer_master(
                    source_df,
                    target_df,
                    rules,
                    primary_key="CUNO",
                    selected_rule_type=selected_rule_type,
                    ignored_columns=ignored_columns,
                )

                self._load_results_into_tree(self.results_df)
                self.summary_label.configure(
                    text=(
                        f"Module: {selected_module}. Business Unit: {business_unit}. "
                        f"Company: {selected_company}. "
                        f"Rule Type: {selected_rule_type}. Rules loaded: {len(rules):,}. "
                        f"Exception columns skipped: {len(ignored_columns):,}. "
                        f"Rows compared: source={len(source_df):,}, target={len(target_df):,}. "
                        f"Issues found: {len(self.results_df):,}."
                    )
                )
                self._set_status("Customer Master comparison complete.")
            except Exception as exc:
                self._set_status("Comparison failed.")
                messagebox.showerror("Compare Error", str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _load_results_into_tree(self, df: pd.DataFrame) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree["columns"] = list(df.columns)

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=180, anchor="w")

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=["" if value is None else value for value in row.tolist()])

    def export_results(self) -> None:
        if self.results_df.empty:
            messagebox.showinfo("No results", "There are no comparison results to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save comparison results",
        )

        if not file_path:
            return

        self.results_df.to_excel(file_path, index=False)
        messagebox.showinfo("Export Complete", f"Saved results to:\n{file_path}")


class CompareApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SQL Server Database Compare")
        self.geometry("1200x720")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        DatabaseCompareHub(self).grid(row=0, column=0, sticky="nsew")


if __name__ == "__main__":
    app = CompareApp()
    app.mainloop()
