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


def test_supplier_master_query_omits_missing_cidref_join():
    query = compare_tool.supplier_master_query(
        company="20",
        include_address=True,
        include_reference=False,
    )

    assert "dbo.CIDADR" in query
    assert "dbo.CIDREF" not in query
    assert "r.*" not in query
    assert "m.IDCONO = '20'" in query
