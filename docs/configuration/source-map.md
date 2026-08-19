# Source Map

The **Source Map** defines where the Data Migration Tool obtains the legacy data for each migration business object.

It connects an MCO business object or worksheet with the source that contains the data to be migrated.

---

## Purpose

The Source Map answers a fundamental question:

> Where does the source data for this migration come from?

Conceptually:

```text
MCO Business Object
        │
        ▼
    Source Map
        │
        ▼
   Source Location
        │
        ▼
Legacy Migration Data
```

This allows the migration process to use a predefined source instead of requiring the user to manually locate the source every time.

---

# Source Map File

The Source Map is stored in:

```text
config/source_map.csv
```

The Data Migration Tool reads this configuration when displaying available sources in the **Standard Migration** screen.

---

# MCO_SHEET

`MCO_SHEET` identifies the migration business object.

For example:

```text
Item Master
Item Facility
Item Warehouse
Supplier
Customer
```

The value should correspond to the MCO business object being configured.

Conceptually:

```text
MCO_SHEET
    │
    ▼
Item Master
```

The application uses this value to locate the associated source.

---

# SOURCE_FILE

`SOURCE_FILE` identifies where the legacy data should be obtained.

Depending on the migration configuration, this may represent a file or a SQL source.

Example:

```text
MCO_SHEET, SOURCE_FILE
Item Master, raw_data/ItemMaster.xlsx
```

This tells the Data Migration Tool:

```text
Item Master
     │
     ▼
raw_data/ItemMaster.xlsx
```

---

# File Sources

A source can point to a legacy Excel or CSV file.

For example:

```text
raw_data/ItemMaster.xlsx
```

or:

```text
raw_data/ItemMaster.csv
```

During migration, the Data Extractor loads the configured source into the migration process.

The source columns then become available to the Rule Configuration.

Example:

```text
Legacy File
│
├── MBITNO
├── MBPUIT
├── MBSTAT
└── ...
       │
       ▼
Transformation Rules
```

---

# SQL Sources

The Data Migration Tool also supports SQL as a source.

A SQL source is identified using the prefix:

```text
SQL:
```

For example:

```text
SQL:SELECT * FROM MITMAS
```

The application recognizes the `SQL:` prefix and executes the remaining text as a SQL query.

Conceptually:

```text
SOURCE_FILE
     │
     ▼
SQL:SELECT ...
     │
     ▼
SQL Server
     │
     ▼
DataFrame
     │
     ▼
Migration Rules
```

!!! important
    SQL sources require the database connection to be configured correctly before the migration can run.

---

# Source Selection in Standard Migration

When the user opens:

**Run Migration → Standard**

the **Select Source Data** dropdown is populated using the MCO sheets defined in:

```text
config/source_map.csv
```

For example:

```text
Select Source Data:

Item Master
Item Facility
Item Warehouse
Supplier
```

When the user selects:

```text
Item Master
```

the application searches the Source Map for that `MCO_SHEET` and retrieves its `SOURCE_FILE`.

The user therefore selects the **business object**, while the application resolves the actual source location.

---

# Example

Consider this Source Map:

```csv
MCO_SHEET,SOURCE_FILE
Item Master,raw_data/ItemMaster.xlsx
Item Facility,raw_data/ItemFacility.xlsx
Supplier,SQL:SELECT * FROM CIDMAS
```

Selecting:

```text
Item Master
```

resolves to:

```text
raw_data/ItemMaster.xlsx
```

Selecting:

```text
Supplier
```

resolves to:

```text
SQL:SELECT * FROM CIDMAS
```

---

# Relationship with the Rule Configuration

The Source Map determines **where the data comes from**.

The Rule Configuration determines **what happens to the data**.

```text
Source Map
    │
    ▼
Legacy Source Data
    │
    ▼
Rule Configuration
    │
    ▼
Transformed M3 Data
```

For example:

```text
Source Map
    ↓
Item Master → ItemMaster.xlsx

Source Data
    ↓
MBITNO = AP32508

Rule
    ↓
TARGET_FIELD = ITNO
SOURCE_FIELD = MBITNO
RULE_TYPE = DIRECT

Output
    ↓
ITNO = AP32508
```

---

# Relationship with the Migration Map

The Source Map and Migration Map perform different functions.

```text
                 MCO Sheet
                    │
          ┌─────────┴─────────┐
          ▼                   ▼
     Source Map          Migration Map
          │                   │
          ▼                   ├── API
    Legacy Source             ├── SDT Template
                              └── Transaction Sheets
```

Together they provide the information required to execute the migration.

---

# Source Field Names

The source data must contain the columns expected by the Rule Configuration.

For example, if a rule contains:

```text
SOURCE_FIELD = MBPUIT
```

the selected source should contain:

```text
MBPUIT
```

If the rule expects one field but the source contains a different field name, the transformation may return a blank or incorrect value.

!!! warning
    Always verify that the Source Map points to the correct dataset before troubleshooting transformation rules.

---

# Updating the Source Map

When configuring a new migration object:

1. Identify the MCO business object.
2. Identify the legacy source.
3. Add the business object to `MCO_SHEET`.
4. Define the source in `SOURCE_FILE`.
5. Save the Source Map.
6. Refresh the source list in Standard Migration.
7. Confirm the new business object appears.
8. Test that the source can be loaded.

---

# Troubleshooting

If a source does not appear in Standard Migration, verify:

- `config/source_map.csv` exists.
- The `MCO_SHEET` column exists.
- The business object has been added to the Source Map.
- The CSV file was saved correctly.

If the source appears but the migration cannot load it, verify:

- `SOURCE_FILE` contains the correct location.
- The source file exists.
- The file is accessible.
- The SQL configuration is correct when using `SQL:`.
- The SQL query is valid.
- The expected source columns exist.

---

# Recommended Practice

The Source Map should represent the approved source for each migration object.

Avoid maintaining multiple ambiguous entries for the same MCO business object unless the migration design specifically requires them.

Before running a migration, the relationship should be clear:

```text
Business Object
      │
      ▼
Approved Source
      │
      ▼
Rule Configuration
      │
      ▼
M3 Output
```

---

# Next Step

After defining the source, configure the **Migration Map**.

The Migration Map determines which M3 API, SDT template, and transaction sheets are associated with the business object.