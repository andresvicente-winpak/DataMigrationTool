# Standard Migration

The **Standard Migration** is the primary method for running a configured migration in the Data Migration Tool.

It allows the user to select a Rule Configuration, choose the corresponding legacy source, optionally select a business-unit scope, and generate the M3 SDT output.

---

## Purpose

Standard Migration should be used when the migration object has already been configured and the source data is known.

The basic process is:

```text
Rule Configuration
        +
Legacy Source
        +
Migration Scope
        │
        ▼
Data Migration Tool
        │
        ▼
Apply FILTER Rules
        │
        ▼
Apply Transformation Rules
        │
        ▼
Populate SDT Template
        │
        ▼
M3 Migration Output
```

---

# Accessing Standard Migration

Navigate to:

**Run Migration → Standard Migration**

The Standard Migration screen provides the options required to execute the migration.

---

# 1. Select Rule Configuration

Select the Rule Configuration associated with the migration.

For example:

```text
MMS200MI
```

The Rule Configuration determines how legacy source fields are transformed into M3 target fields.

It contains information such as:

```text
TARGET_FIELD
SOURCE_FIELD
RULE_TYPE
RULE_VALUE
SCOPE
```

Example:

```text
MBITNO
   │
   │ DIRECT
   ▼
ITNO
```

For more complex transformations:

```text
ITNO ──┐
ITTY ──┼──► PYTHON Rule ──► PLCD
PUIT ──┘
```

---

# 2. Select Source Data

Select the migration business object containing the source data.

For example:

```text
Item Master
Item Facility
Item Warehouse
Supplier
Customer
```

The available selections come from:

```text
config/source_map.csv
```

The selected business object is then resolved to its configured source.

Example:

```text
Item Master
     │
     ▼
Source Map
     │
     ▼
raw_data/ItemMaster.xlsx
```

The source may be:

```text
Excel
CSV
SQL
```

---

# 3. Select Scope

The Scope option determines which version of the transformation rules should be applied.

The default is:

```text
GLOBAL
```

Additional business-unit scopes may also be available.

For example:

```text
GLOBAL
DIV_US
DIV_CA
```

If no business-specific transformation is required, use:

```text
GLOBAL
```

---

## GLOBAL and Scope Overrides

GLOBAL rules provide the default transformation behavior.

For example:

```text
PLCD / GLOBAL
```

A business unit may override that rule:

```text
PLCD / DIV_US
```

When:

```text
Scope = DIV_US
```

the business-specific PLCD rule is used instead of the GLOBAL PLCD rule.

Other fields continue using their GLOBAL rules unless they also have an override.

---

# 4. Run Migration

After selecting:

```text
Rule Configuration
Source Data
Scope
```

click:

**RUN MIGRATION**

The application begins processing the migration.

---

# What Happens During Standard Migration?

The application performs several steps automatically.

```text
START
  │
  ▼
Resolve Source
  │
  ▼
Load Legacy Data
  │
  ▼
Load Rule Configuration
  │
  ▼
Apply FILTER Rules
  │
  ▼
Transform Target Fields
  │
  ▼
Resolve Migration Map
  │
  ├── API
  ├── SDT Template
  └── Transaction Sheets
  │
  ▼
Populate SDT
  │
  ▼
Save Output
```

---

# Loading the Source

The Data Extractor loads the configured source.

For an Excel source:

```text
Excel File
    │
    ▼
DataFrame
```

For CSV:

```text
CSV File
    │
    ▼
DataFrame
```

For SQL:

```text
SQL Query
    │
    ▼
SQL Server
    │
    ▼
DataFrame
```

Once loaded, the source columns become available to the transformation rules.

---

# FILTER Processing

FILTER rules are processed before the target transformations.

For example:

```text
Source
10,000 records
     │
     ▼
FILTER Rules
     │
     ▼
8,450 records
```

Only the remaining records continue through the migration.

If every record is filtered out, the migration cannot continue.

---

# Transformation Processing

The Rule Configuration determines how each M3 target field is populated.

For example:

### DIRECT

```text
MBITNO
   │
   ▼
ITNO
```

### CONST

```text
100
 │
 ▼
CONO
```

### MAP

```text
Legacy Code
    │
    ▼
Mapping
    │
    ▼
M3 Code
```

### PYTHON

```text
Multiple Source Fields
         │
         ▼
    Business Logic
         │
         ▼
     Target Value
```

---

# Migration Map Resolution

The selected MCO business object is used to resolve the corresponding Migration Map configuration.

The Migration Map identifies:

```text
API_NAME
SDT_TEMPLATE
TRANSACTION_SHEET
```

Conceptually:

```text
MCO Business Object
        │
        ▼
Migration Map
        │
        ├── API
        ├── SDT
        └── Transaction Sheets
```

---

# SDT Generation

After transformation, the resulting values are written into the configured M3 SDT template.

```text
Transformed Data
       │
       ▼
SDT Template
       │
       ▼
Transaction Sheet(s)
       │
       ▼
Generated Workbook
```

The original SDT structure is used as the foundation for the generated migration file.

---

# Output

The generated migration file is normally saved under:

```text
output/
```

The System Log provides information about the migration process and generated output.

---

# Monitor the System Log

While the migration runs, review the **System Log**.

The log can provide information about:

```text
Source loading
Rule loading
FILTER processing
Transformation
SDT generation
Warnings
Errors
Output location
```

If the migration fails, the System Log should be one of the first places reviewed.

---

# Example Standard Migration

Suppose you want to migrate Item Master.

Select:

```text
Rule Configuration:
MMS200MI

Source:
Item Master

Scope:
GLOBAL
```

The application resolves:

```text
Item Master
    │
    ├── Source Map
    │      └── Legacy Item Source
    │
    └── Migration Map
           ├── MMS200MI
           ├── SDT Template
           └── Transaction Sheets
```

The rules are then applied and the M3 migration workbook is generated.

---

# Before Running the Migration

Verify:

- [ ] Correct Rule Configuration
- [ ] Correct source
- [ ] Correct scope
- [ ] Source data is available
- [ ] Rule changes have been saved
- [ ] Required mapping files exist
- [ ] SDT template exists
- [ ] Migration Map is configured

---

# After Running the Migration

Do not stop validation at:

```text
Migration Completed Successfully
```

Review the generated output.

Validate:

- [ ] Expected record count
- [ ] Required target fields
- [ ] DIRECT transformations
- [ ] MAP transformations
- [ ] Important PYTHON rules
- [ ] Scope-specific rules
- [ ] SDT transaction sheets
- [ ] Representative business records

---

# Successful Execution vs Successful Migration

These are not necessarily the same thing.

```text
Successful Execution

The application completed
without a technical failure.
```

versus:

```text
Successful Migration

The application completed
        +
The resulting data matches
the approved business requirements.
```

A migration should only be considered ready after the generated output has been validated.

---

# When to Use Another Migration Mode

Use **Auto-Detect Migration** when:

> You have a source file but are unsure which migration configuration should process it.

Use **Load by ID** when:

> You want to test or migrate specific records.

Use **Batch Migration** when:

> Multiple already-configured migrations need to be processed together.

Use **Standard Migration** when:

> You know the migration configuration and want to process its configured source normally.

---

# Recommended Workflow

For a new or modified migration:

```text
Configure
    │
    ▼
Develop Rules
    │
    ▼
Load by ID
    │
    ▼
Validate Test Cases
    │
    ▼
Standard Migration
    │
    ▼
Validate Complete Output
    │
    ▼
Batch Migration
```

This progression makes rule problems easier to identify before processing larger migration populations.