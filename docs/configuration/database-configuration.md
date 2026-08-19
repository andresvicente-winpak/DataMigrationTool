# Database Configuration

The **Data Migration Tool** supports SQL Server as a source of legacy migration data.

Instead of loading the source from an Excel or CSV file, a migration can execute a SQL query and use the returned records as the source dataset.

---

## SQL Migration Flow

When SQL is configured as the source, the migration flow becomes:

```text
SQL Server
    │
    ▼
SQL Query
    │
    ▼
Data Extractor
    │
    ▼
Source DataFrame
    │
    ▼
FILTER Rules
    │
    ▼
Transformation Rules
    │
    ▼
M3 SDT Output
```

From the Rule Configuration perspective, the resulting SQL data behaves similarly to data loaded from Excel.

---

# Database Configuration File

Database connection information is stored in:

```text
config/db_config.ini
```

The Data Extractor reads this file when a SQL source is detected.

A configuration follows this general structure:

```ini
[DEFAULT]
Driver=ODBC Driver 17 for SQL Server
Server=SERVER_NAME
Database=DATABASE_NAME
Trusted_Connection=yes
```

---

# Configuration Fields

## Driver

`Driver` identifies the ODBC driver used to connect to SQL Server.

Example:

```ini
Driver=ODBC Driver 17 for SQL Server
```

The specified driver must be installed on the computer running the Data Migration Tool.

---

## Server

`Server` identifies the SQL Server instance.

Example:

```ini
Server=SQLSERVER01
```

Depending on the environment, this may also contain a named instance.

---

## Database

`Database` identifies the database containing the migration source data.

Example:

```ini
Database=LegacyDB
```

The SQL query configured for the migration will execute against this database.

---

## Trusted Connection

The current database configuration supports Windows authentication through:

```ini
Trusted_Connection=yes
```

This means SQL Server uses the Windows credentials of the user running the Data Migration Tool.

Conceptually:

```text
Windows User
     │
     ▼
Data Migration Tool
     │
     ▼
SQL Server Authentication
     │
     ▼
Configured Database
```

The user must therefore have the required SQL Server permissions.

---

# Required Python Libraries

SQL connectivity requires:

```text
sqlalchemy
pyodbc
```

If these libraries are not installed, SQL sources cannot be loaded.

They can be installed using:

```bash
pip install sqlalchemy pyodbc
```

---

# SQL Sources

A source is recognized as SQL when it begins with:

```text
SQL:
```

Example:

```text
SQL:SELECT * FROM MITMAS
```

The Data Extractor removes the `SQL:` prefix and executes:

```sql
SELECT * FROM MITMAS
```

against the configured database.

---

# Source Map Example

SQL can be configured directly as the source in the Source Map.

For example:

```csv
MCO_SHEET,SOURCE_FILE
Item Master,SQL:SELECT * FROM MITMAS
```

When **Item Master** is selected during Standard Migration:

```text
Item Master
     │
     ▼
Source Map
     │
     ▼
SQL:SELECT * FROM MITMAS
     │
     ▼
SQL Server
```

The returned records become the source dataset used by the migration rules.

---

# Using SQL Columns in Rules

Columns returned by the SQL query become available to the Rule Configuration.

For example:

```sql
SELECT
    MBITNO,
    MBPUIT,
    MBSTAT
FROM MITMAS
```

produces source columns:

```text
MBITNO
MBPUIT
MBSTAT
```

A rule can then use:

```text
TARGET_FIELD = ITNO
SOURCE_FIELD = MBITNO
RULE_TYPE    = DIRECT
```

A Python rule can also access the returned fields:

```python
itno = str(row.get("MBITNO", "")).strip()
puit = str(row.get("MBPUIT", "")).strip()
```

---

# NULL Values

SQL results may contain `NULL` values.

The Data Extractor normalizes SQL data after loading it so that null-like values can be handled consistently during transformation.

Rules should still handle blank values when the business requirement requires it.

Example:

```python
value = str(row.get("MBPUIT", "")).strip()

if value == "":
    return ""

return value
```

---

# Database Permissions

The Windows account running the Data Migration Tool must have permission to:

- Connect to the SQL Server
- Access the configured database
- Read the tables or views referenced by the query

For migration extraction, read-only database access is normally sufficient when the configured queries only retrieve source data.

!!! warning
    SQL queries used as migration sources should be reviewed carefully. The Data Migration Tool executes the SQL text configured after the `SQL:` prefix.

---

# Recommended SQL Queries

Prefer selecting only the fields required for the migration.

Instead of:

```sql
SELECT *
FROM MITMAS
```

consider:

```sql
SELECT
    MBITNO,
    MBITTY,
    MBPUIT,
    MBSTAT
FROM MITMAS
```

This makes the source definition clearer and reduces unnecessary data retrieval.

When appropriate, SQL can also restrict the migration population:

```sql
SELECT
    MBITNO,
    MBITTY,
    MBPUIT,
    MBSTAT
FROM MITMAS
WHERE MBCONO = 100
```

---

# SQL Filtering vs FILTER Rules

SQL conditions and Data Migration Tool FILTER rules can both restrict the source population, but they operate at different stages.

```text
SQL Server
    │
    │ WHERE condition
    ▼
Extracted Dataset
    │
    │ FILTER rule
    ▼
Migration Dataset
```

A SQL `WHERE` clause limits which records are extracted from the database.

A `FILTER` rule limits which extracted records continue through the migration process.

The migration design should clearly identify where population-selection logic belongs.

---

# Testing the Connection

Before using a SQL source for migration, verify:

1. The ODBC driver is installed.
2. `db_config.ini` exists.
3. The server name is correct.
4. The database name is correct.
5. Windows authentication is permitted.
6. The user has access to the database.
7. `sqlalchemy` and `pyodbc` are installed.
8. The SQL query executes successfully.
9. The returned columns match the Rule Configuration.

---

# Common Errors

## SQLAlchemy or pyodbc Not Installed

The application may report that SQL support is unavailable.

Install:

```bash
pip install sqlalchemy pyodbc
```

---

## Database Configuration Not Found

Verify that the following file exists:

```text
config/db_config.ini
```

---

## Login or Permission Error

Verify that the Windows user running the application has access to the configured SQL Server and database.

---

## Invalid Object Name

If SQL Server reports an invalid table or view, verify:

- Database
- Schema
- Table name
- View name

For example, the fully qualified object may need to be:

```sql
dbo.MITMAS
```

instead of:

```sql
MITMAS
```

depending on the database configuration.

---

## Missing Source Field

If the migration runs but a rule returns blank values, verify that the SQL query actually returns the field expected by `SOURCE_FIELD`.

For example:

```text
Rule expects:

MBPUIT
```

but the query only returns:

```sql
SELECT MBITNO, MBSTAT
FROM MITMAS
```

`MBPUIT` will not be available to the transformation rule.

---

# Security Considerations

Database configuration should be treated as environment-specific configuration.

Avoid placing passwords or other sensitive credentials directly inside Rule Configuration files or SQL queries.

The current trusted-connection configuration allows Windows authentication to be used without storing a database password in the migration rule.

---

# Recommended Practice

For SQL-based migrations:

1. Configure the SQL Server connection.
2. Test database connectivity.
3. Create a focused source query.
4. Verify the returned columns.
5. Add the SQL source to the Source Map.
6. Verify the Rule Configuration dependencies.
7. Run a small test migration.
8. Compare the SQL source records with the generated SDT.
9. Validate record counts and transformed values before production migration.

The complete relationship becomes:

```text
db_config.ini
      │
      ▼
SQL Server
      │
      ▼
SQL Query
      │
      ▼
Source Map
      │
      ▼
Rule Configuration
      │
      ▼
SDT Output
```

---

# Next Step

With the core configuration documented, the next section covers **Run Migration**, including the different migration methods available in the Data Migration Tool.