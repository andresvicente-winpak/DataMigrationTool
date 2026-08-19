# Common Errors and Troubleshooting

This section provides guidance for diagnosing common problems when using the Data Migration Tool.

A migration can complete without a technical error while still producing an incorrect business result. Troubleshooting should therefore consider both:

```text
Technical Success
        +
Business Validation
        │
        ▼
Successful Migration
```

---

# Recommended Troubleshooting Approach

When something is wrong, identify where the problem occurs:

```text
Source Data
    │
    ▼
Source Map
    │
    ▼
Migration Map
    │
    ▼
Rule Configuration
    │
    ▼
FILTER Rules
    │
    ▼
Transformation Rules
    │
    ▼
SDT Template
    │
    ▼
Generated Output
```

Start with the source and work toward the output.

---

# Rule Is Not Producing the Expected Value

## Symptom

The migration completes, but a target field contains an unexpected value.

Example:

```text
Expected:

PLCD = 02

Actual:

PLCD = 01
```

## What to Check

### 1. Verify the raw source record

Do not begin by looking only at the generated output.

Check the actual source values required by the rule.

Example:

```text
MBITNO = AP32508
MMITTY = FG
MBPUIT = 1
```

### 2. Identify rule dependencies

For example:

```text
PLCD
 │
 ├── ITNO
 ├── ITTY
 └── PUIT
```

Verify every dependency.

### 3. Review the Rule Configuration

Check:

```text
TARGET_FIELD
SOURCE_FIELD
RULE_TYPE
RULE_VALUE
SCOPE
```

### 4. Verify the selected scope

A scope-specific rule may override the GLOBAL rule.

### 5. Test the record individually

Use **Load by ID** when possible to isolate the problematic record.

---

# Python Rule Returns the Wrong Result

## Symptom

A `PYTHON` rule executes successfully but returns an incorrect value.

## Common Causes

- Wrong source column
- Incorrect condition order
- Unexpected spaces
- Upper/lowercase differences
- Numeric vs string comparison
- Missing dependency
- Blank/null source value

---

## Check Column Names

For example:

```python
row.get("MMITTY", "")
```

is different from:

```python
row.get("MBITTY", "")
```

A small difference in the field name can completely change the result.

---

## Normalize Values

Prefer:

```python
itty = str(row.get("MMITTY", "")).strip().upper()
```

instead of:

```python
itty = row.get("MMITTY")
```

For migration codes, converting to strings before comparison is often safer:

```python
puit = str(row.get("MBPUIT", "")).strip()

if puit == "3":
    return "05"
```

---

# Condition Order Problem

Python rules are evaluated from top to bottom.

Example:

```python
if puit == "3":
    return "05"

if itty == "FG" and puit == "3":
    return "02"
```

For:

```text
ITTY = FG
PUIT = 3
```

the result is:

```text
05
```

The second condition is never reached.

When conditions overlap, verify their priority.

---

# Source Field Is Blank

## Symptom

A target field is unexpectedly blank.

## Check

Verify that the configured `SOURCE_FIELD` actually exists in the selected source.

Example:

```text
Rule expects:

MBPUIT
```

but the source contains:

```text
PUIT
```

The rule may not receive the expected value.

For Python rules that intentionally support both names:

```python
puit = str(
    row.get("PUIT", row.get("MBPUIT", ""))
).strip()
```

---

# Wrong Scope Is Being Used

## Symptom

The rule works under GLOBAL but produces a different result for another migration run.

## Check

Verify the selected:

```text
Scope
```

Example:

```text
GLOBAL
PLCD → 02

DIV_US
PLCD → 03
```

The same source record can produce different results depending on the selected scope.

Also verify the scope defined in the Rule Configuration.

---

# Rule Change Is Not Appearing

## Symptom

A rule was modified, but the migration still appears to use previous logic.

## Check

- Confirm the Rule Configuration was saved.
- Confirm the correct Rule Configuration was selected.
- Confirm the correct scope was selected.
- Confirm you modified the correct target field.
- Review Rule History.
- Verify that another scoped rule is not overriding the changed rule.

When necessary, close and reopen the configuration before testing again.

---

# MAP Rule Returns Blank

## Symptom

A MAP rule does not return a target value.

## Check

Verify:

- Mapping file exists
- Mapping path is correct
- Key column exists
- Value column exists
- Source value exists in the mapping
- Source field is correct

Example:

```text
Mapping:

A → 01
B → 02
C → 03
```

If the source contains:

```text
X
```

there is no configured result.

---

# Too Many Records Are Missing

## Symptom

The generated migration contains fewer records than expected.

## Check FILTER Rules

FILTER rules determine which records continue through the migration.

Remember:

```python
source == "20"
```

means:

> Keep rows where source equals 20.

It does not mean:

> Remove rows where source equals 20.

---

## Compare Record Counts

For example:

```text
Source Records:       10,000
After FILTER Rules:    8,450
Generated Records:     8,450
```

The migration team should understand why:

```text
1,550
```

records were excluded.

---

# All Rows Were Filtered Out

## Symptom

The migration stops because no records remain after filtering.

## Check

Review all `FILTER` rules.

When multiple filters exist:

```text
Original Data
     │
     ▼
Filter 1
     │
     ▼
Filter 2
     │
     ▼
Filter 3
     │
     ▼
0 Records
```

Each filter reduces the remaining population.

Test the filters individually if necessary.

---

# Source Does Not Appear in Standard Migration

## Check

Verify:

```text
config/source_map.csv
```

Confirm:

- The file exists
- `MCO_SHEET` exists
- The expected business object is configured
- The CSV was saved correctly

Then refresh the source list.

---

# Source Cannot Be Loaded

If the source appears but cannot be loaded, verify:

- `SOURCE_FILE` is correct
- File exists
- File is accessible
- File is not corrupted
- File format is supported
- SQL configuration is valid when using `SQL:`

---

# Migration Configuration Cannot Be Found

## Check

Verify:

```text
config/migration_map.csv
```

Confirm the relationship between:

```text
MCO_SHEET
API_NAME
SDT_TEMPLATE
TRANSACTION_SHEET
```

The selected MCO business object must resolve to the correct migration configuration.

---

# SDT Template Cannot Be Found

## Symptom

The migration cannot locate the required SDT workbook.

## Check

Verify that the expected template exists under:

```text
config/sdt_templates/
```

Also check the `SDT_TEMPLATE` value in:

```text
config/migration_map.csv
```

Confirm the filename matches the actual template.

---

# Transaction Sheet Cannot Be Found

## Symptom

The SDT workbook exists, but the migration cannot populate the expected worksheet.

## Check

Verify:

```text
TRANSACTION_SHEET
```

in the Migration Map.

The configured sheet name must correspond to the expected worksheet in the SDT template.

Be especially careful with:

- Spaces
- Spelling
- Multiple transaction sheets

---

# SQL Source Does Not Work

## Check Database Configuration

Verify:

```text
config/db_config.ini
```

Check:

- ODBC driver
- Server
- Database
- Trusted Connection

Also verify that:

```text
sqlalchemy
pyodbc
```

are installed.

---

# SQL Query Works but Rules Return Blank

The SQL query may not be returning all fields required by the Rule Configuration.

Example:

```sql
SELECT
    MBITNO,
    MBSTAT
FROM MITMAS
```

But the rule expects:

```text
MBPUIT
```

`MBPUIT` is not available.

Review the dependencies of the affected rules and compare them with the SQL query output.

---

# Output File Cannot Be Written

## Check

Verify:

- Output directory exists
- User has write permission
- Existing file is not locked
- Excel does not have the output file open
- Sufficient disk space is available

The normal output location is:

```text
output/
```

---

# Migration Runs but Output Is Incorrect

This is one of the most important troubleshooting scenarios.

A message indicating that the migration completed successfully means the application completed the processing workflow.

It does **not** guarantee that every business transformation is correct.

Use:

```text
Source
   │
   ▼
Expected Business Rule
   │
   ▼
Expected Result
   │
   ↕
Actual Output
```

For example:

```text
Source:

ITNO = AP32508
ITTY = FG
PUIT = 1

Requirement:

FG + PUIT 1 → PLCD 02

Expected:

PLCD = 02

Actual:

PLCD = 01
```

This should be treated as a rule/configuration problem even if no application error occurred.

---

# Use Load by ID for Troubleshooting

When investigating a specific record, avoid repeatedly processing the complete dataset.

Use:

**Run Migration → Load by ID**

to isolate representative records.

For example:

```text
AP32508
```

Then review:

```text
Raw Source
    ↓
Rule Dependencies
    ↓
Rule Logic
    ↓
Expected Value
    ↓
Actual Value
```

This is usually much easier than troubleshooting thousands of records at once.

---

# After a Rule Change

Whenever a rule is corrected:

```text
Modify Rule
    │
    ▼
Save
    │
    ▼
Load by ID
    │
    ▼
Test Special Cases
    │
    ▼
Standard Migration
    │
    ▼
Validate Output
```

Do not immediately move from a rule change to a production migration.

---

# Troubleshooting Checklist

When a generated value is incorrect, use this checklist:

- [ ] Am I looking at the correct raw source record?
- [ ] Is the correct Source Map entry being used?
- [ ] Is the correct Migration Map entry being used?
- [ ] Is the correct Rule Configuration selected?
- [ ] Is `TARGET_FIELD` correct?
- [ ] Is `SOURCE_FIELD` correct?
- [ ] Is `RULE_TYPE` correct?
- [ ] Is `RULE_VALUE` correct?
- [ ] Are all rule dependencies available?
- [ ] Is the correct scope selected?
- [ ] Are FILTER rules changing the population?
- [ ] Is the SDT template correct?
- [ ] Is the correct transaction sheet being populated?
- [ ] Does the expected result match the approved business requirement?

---

# Recommended Troubleshooting Order

Avoid changing rules immediately when an output value looks wrong.

Follow this order:

```text
1. Raw Source
      ↓
2. Source Field
      ↓
3. Rule Dependencies
      ↓
4. Business Requirement
      ↓
5. Rule Configuration
      ↓
6. Scope
      ↓
7. Expected Result
      ↓
8. Actual Output
```

Only modify the rule after identifying where the difference originates.

This helps prevent fixing one scenario while unintentionally breaking another.