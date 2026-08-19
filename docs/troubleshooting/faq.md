# Frequently Asked Questions

This page provides quick answers to common questions when working with the Data Migration Tool.

For more detailed troubleshooting, see [Common Errors and Troubleshooting](common-errors.md).

---

## Why is my target field blank?

Check the Rule Configuration first.

Verify:

```text
TARGET_FIELD
SOURCE_FIELD
RULE_TYPE
RULE_VALUE
SCOPE
```

Then confirm that the expected `SOURCE_FIELD` actually exists in the source data.

For a Python rule, also verify every field accessed through `row`.

Example:

```python
puit = str(row.get("MBPUIT", "")).strip()
```

If `MBPUIT` does not exist in the source, the rule receives a blank value.

---

## Why is my Python rule returning the wrong value?

Check:

1. Source values
2. Source column names
3. String versus numeric comparisons
4. Spaces and capitalization
5. Condition order
6. Rule dependencies
7. Selected scope

For example:

```python
puit = str(row.get("MBPUIT", "")).strip()
```

should normally be compared with:

```python
if puit == "3":
```

not:

```python
if puit == 3:
```

---

## Why does condition order matter?

Python rules execute from top to bottom.

Example:

```python
if puit == "3":
    return "05"

if itty == "FG" and puit == "3":
    return "02"
```

The second condition will never execute when `PUIT = 3` because the first condition already returned a result.

Place higher-priority conditions before more general conditions.

---

## Why does the rule work for GLOBAL but not for my division?

A business-unit-specific rule may be overriding the GLOBAL rule.

For example:

```text
PLCD / GLOBAL
PLCD / DIV_US
```

When:

```text
Scope = DIV_US
```

the `DIV_US` rule takes precedence for `PLCD`.

Verify both the selected migration scope and the `SCOPE` value in the Rule Configuration.

---

## When should I use DIRECT?

Use `DIRECT` when the source value can be transferred directly to the target field without additional business logic.

Example:

```text
MBITNO → ITNO
```

---

## When should I use CONST?

Use `CONST` when every processed record should receive the same value.

Example:

```text
CONO = 100
```

---

## When should I use MAP?

Use `MAP` when source values need to be translated using a lookup table.

Example:

```text
A → 01
B → 02
C → 03
```

MAP is generally preferable to writing a long Python rule containing many simple value-to-value conversions.

---

## When should I use PYTHON?

Use `PYTHON` when the target value depends on:

- Multiple fields
- Conditional logic
- String manipulation
- Special business requirements
- More complex transformations

For example:

```text
ITNO ──┐
ITTY ──┼──► PYTHON ──► PLCD
PUIT ──┘
```

---

## When should I use FILTER?

Use `FILTER` when certain source records should not continue through the migration.

Remember that the filter expression represents records to **keep**.

Example:

```python
source == "20"
```

means:

> Keep records where the source value is 20.

---

## Why are records missing from my output?

Check the number of records at each stage:

```text
Source
   │
   ▼
FILTER Rules
   │
   ▼
Remaining Records
   │
   ▼
Output
```

FILTER rules are one of the first things to investigate when the output contains fewer records than expected.

---

## Why did the migration say all rows were filtered out?

Your FILTER rules removed every source record.

Review each filter individually and confirm that the conditions represent records that should be **kept**.

---

## Why doesn't my source appear in Standard Migration?

Check:

```text
config/source_map.csv
```

Verify that the expected business object exists under:

```text
MCO_SHEET
```

Then refresh the source list.

---

## What is the difference between Source Map and Migration Map?

The **Source Map** determines where the data comes from.

```text
MCO Sheet → Legacy Source
```

The **Migration Map** determines which M3 configuration should process it.

```text
MCO Sheet
   │
   ├── API
   ├── SDT Template
   └── Transaction Sheets
```

Together:

```text
Source Map
    │
    ▼
Legacy Data
    │
    ▼
Migration Map
    │
    ▼
M3 Configuration
```

---

## Why can't the tool find my SDT template?

Check:

```text
config/migration_map.csv
```

and verify the configured:

```text
SDT_TEMPLATE
```

Then confirm that the template exists under:

```text
config/sdt_templates/
```

---

## Why can't the tool find the transaction sheet?

Verify:

```text
TRANSACTION_SHEET
```

in the Migration Map.

The configured worksheet must exist in the corresponding SDT template.

Check spelling and spaces carefully.

---

## Can I use Excel, CSV, and SQL as sources?

Yes. The migration source can be configured from supported file sources or SQL.

SQL sources use:

```text
SQL:
```

For example:

```text
SQL:SELECT * FROM MITMAS
```

SQL connectivity also requires the database configuration to be available.

---

## Why does my SQL migration return blank fields?

Confirm that the SQL query returns every source column required by the rules.

For example, if a rule needs:

```text
MBPUIT
```

but the query only returns:

```sql
SELECT MBITNO, MBSTAT
FROM MITMAS
```

then `MBPUIT` is unavailable to the rule.

---

## How can I test one problematic record?

Use:

**Run Migration → Load by ID**

This is useful when troubleshooting records such as:

```text
AP32508
```

Instead of reviewing thousands of rows, you can focus on:

```text
Source Record
     ↓
Rule Dependencies
     ↓
Expected Result
     ↓
Actual Result
```

---

## Should I test a rule before running a complete migration?

Yes.

A recommended workflow is:

```text
Modify Rule
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
Validate Output
```

Use representative records for each important business condition.

---

## When should I use Auto-Detect?

Use Auto-Detect when you have a legacy source file but are unsure which configured migration object should process it.

Auto-Detect analyzes source-field signatures and attempts to identify the appropriate migration context.

The detected configuration should still be verified before migration.

---

## When should I use Batch Migration?

Use Batch Migration after the individual migrations have already been configured and tested.

Recommended progression:

```text
Develop
   ↓
Load by ID
   ↓
Standard Migration
   ↓
Validate
   ↓
Batch Migration
```

Batch Migration should not normally be the first test of a new Rule Configuration.

---

## Should I create a snapshot before changing rules?

For significant changes, yes.

For example:

```text
Current Rules
     │
     ▼
Snapshot
     │
     ▼
Modify Rules
     │
     ▼
Test
```

Snapshots provide a reference to the previous Rule Configuration.

---

## What should I do after Sync / Merge?

Identify which rules changed and retest those fields.

Recommended process:

```text
Merge
  ↓
Identify Changes
  ↓
Load by ID
  ↓
Validate
  ↓
Standard Migration
```

A successful merge does not automatically mean the resulting migration values are correct.

---

## Where are generated migration files saved?

Generated migration files are normally stored under:

```text
output/
```

Check the System Log after migration to confirm the generated output location.

---

## Why can't the output file be saved?

Check whether:

- The file is already open in Excel
- The output folder is accessible
- You have write permission
- An existing output file is locked

Closing the workbook in Excel often resolves file-locking issues.

---

## Does "Migration Completed" mean the migration is correct?

No.

It means the application completed the processing workflow successfully.

You should still validate:

```text
Record Counts
      +
Required Fields
      +
Transformation Rules
      +
Business Requirements
      +
SDT Structure
      │
      ▼
Validated Migration
```

Technical success and business correctness are different.

---

# What Should I Check Before Saying a Migration Is Ready?

Use this final checklist:

- [ ] Correct source selected
- [ ] Correct MCO business object
- [ ] Correct API
- [ ] Correct Rule Configuration
- [ ] Correct scope
- [ ] FILTER population validated
- [ ] Important Python rules tested
- [ ] MAP values validated
- [ ] Record counts validated
- [ ] Required fields populated
- [ ] SDT template correct
- [ ] Transaction sheets correct
- [ ] Representative records validated
- [ ] Output file opens successfully
- [ ] System Log reviewed for warnings or errors

A migration should be considered ready only after both the application execution and the resulting business data have been validated.

---

# Need More Detail?

See:

- [Common Errors and Troubleshooting](common-errors.md)
- [Python Rules](../rules/python.md)
- [MAP Rules](../rules/map.md)
- [FILTER Rules](../rules/filter.md)
- [Load by ID](../migration/load-by-id.md)
- [Rule History](../administration/rules-history.md)