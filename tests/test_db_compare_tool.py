import modules.gui.CompareDBTool as compare_tool


class FakeCursor:
    def __init__(self, conn):
        self.conn = conn
        self.rows = []
        self.one = None

    def execute(self, query, *params):
        normalized = " ".join(query.split()).upper()
        self.rows = []
        self.one = None

        if "INFORMATION_SCHEMA.TABLES" in normalized:
            schema, table = params
            self.one = (1,) if (schema.upper(), table.upper()) in self.conn.tables else None
        elif "INFORMATION_SCHEMA.COLUMNS" in normalized:
            schema, table, column = params
            key = (schema.upper(), table.upper(), column.upper())
            self.one = (1,) if key in self.conn.columns else None
        elif "SELECT DISTINCT" in normalized:
            for (schema, table, column), values in self.conn.company_values.items():
                if f"FROM {schema}.{table}".upper() in normalized and column.upper() in normalized:
                    self.rows = [(value,) for value in values]
                    break
        else:
            raise AssertionError(f"Unexpected query: {query}")

        return self

    def fetchone(self):
        return self.one

    def fetchall(self):
        return self.rows


class FakeConnection:
    def __init__(self):
        self.tables = {("DBO", "OCUSMA"), ("DBO", "CIDMAS"), ("DBO", "CIDADR")}
        self.columns = {
            ("DBO", "OCUSMA", "OKCONO"),
            ("DBO", "CIDMAS", "IDCONO"),
            ("DBO", "CIDADR", "SACONO"),
        }
        self.company_values = {
            ("dbo", "OCUSMA", "OKCONO"): ["100", "200"],
            ("dbo", "CIDMAS", "IDCONO"): ["200", "300"],
            ("dbo", "CIDADR", "SACONO"): ["400"],
        }

    def cursor(self):
        return FakeCursor(self)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


def test_list_target_companies_skips_missing_optional_tables(monkeypatch):
    monkeypatch.setattr(compare_tool, "get_connection", lambda config: FakeConnection())

    companies = compare_tool.list_target_companies(compare_tool.SqlServerConfig("server", "database"))

    assert companies == ["All", "100", "200", "300", "400"]


def test_supplier_master_query_uses_cidmas_only():
    query = compare_tool.supplier_master_query(company="20")

    assert "FROM dbo.CIDMAS" in query
    assert "dbo.CIDADR" not in query
    assert "dbo.CIDREF" not in query
    assert "IDCONO = '20'" in query


def test_supplier_address_query_uses_cidadr_only():
    query = compare_tool.supplier_address_query(company="20")

    assert "FROM dbo.CIDADR" in query
    assert "dbo.CIDMAS" not in query
    assert "dbo.CIDREF" not in query
    assert "SACONO = '20'" in query


def test_supplier_compare_remaps_source_suno_before_missing_check(monkeypatch):
    source = compare_tool.pd.DataFrame({"IDSUNO": ["160012"], "IDSUNM": ["ABF Freight"]})
    target = compare_tool.pd.DataFrame({"IDSUNO": ["1202000066"], "IDSUNM": ["ABF Freight"]})
    rules = compare_tool.pd.DataFrame([
        {"SOURCE_FIELD": "IDSUNO", "TARGET_FIELD": "SUNO", "RULE_TYPE": "DIRECT", "RULE_VALUE": ""},
        {"SOURCE_FIELD": "IDSUNM", "TARGET_FIELD": "SUNM", "RULE_TYPE": "DIRECT", "RULE_VALUE": ""},
    ])

    monkeypatch.setattr(
        compare_tool,
        "load_single_column_translation_map",
        lambda file_path, key_column, value_column: {"160012": "1202000066"},
    )

    result = compare_tool.compare_rule_based_customer_master(
        source,
        target,
        rules,
        primary_key="SUNO",
        table_prefixes=("ID",),
        source_key_translation=("translation_tbl/OLD_NEW_SUNO.xlsx", "SUNO", "NEWSUNO"),
    )

    assert result.empty


def test_remap_compare_key_values_keeps_unmapped_suno_values():
    source = compare_tool.pd.DataFrame({"SUNO": ["160012", "160013"], "SUNM": ["ABF", "Other"]})

    result = compare_tool.remap_compare_key_values(source, "SUNO", {"160012": "1202000066"})

    assert result["SUNO"].tolist() == ["1202000066", "160013"]


def test_supplier_address_compare_backfills_address_primary_keys(monkeypatch):
    source = compare_tool.pd.DataFrame({
        "SASUNO": ["160012"],
        "SAADTE": ["1"],
        "SAADID": ["MAIN"],
    })
    target = compare_tool.pd.DataFrame({
        "SASUNO": ["1202000066"],
        "SAADTE": ["1"],
        "SAADID": ["MAIN"],
    })
    rules = compare_tool.pd.DataFrame([
        {"SOURCE_FIELD": "SASUNO", "TARGET_FIELD": "SUNO", "RULE_TYPE": "DIRECT", "RULE_VALUE": ""},
    ])

    monkeypatch.setattr(
        compare_tool,
        "load_single_column_translation_map",
        lambda file_path, key_column, value_column: {"160012": "1202000066"},
    )

    result = compare_tool.compare_rule_based_customer_master(
        source,
        target,
        rules,
        primary_key=["SUNO", "ADTE", "ADID"],
        table_prefixes=("SA",),
        source_key_translation=("translation_tbl/OLD_NEW_SUNO.xlsx", "SUNO", "NEWSUNO"),
    )

    assert result.empty
