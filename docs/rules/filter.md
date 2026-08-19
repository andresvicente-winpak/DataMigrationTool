# FILTER Rules

A `FILTER` rule determines which source records are allowed to continue through the migration process.

Unlike other rule types, a FILTER rule does not calculate an M3 target value.

Instead, it determines whether a source row should be included or excluded.

---

## Purpose

A source file may contain records that should not be migrated.

For example:

```text
Legacy Source
│
├── Active Item
├── Active Item
├── Obsolete Item
├── Test Item
└── Active Item
```

A FILTER rule can remove records that do not meet the migration criteria.

The result becomes:

```text
Migration Dataset
│
├── Active Item
├── Active Item
└── Active Item
```

---

## When FILTER Rules Are Applied

FILTER rules are applied before target field transformations.

The migration flow is:

```text
Legacy Source Data
        │
        ▼
     FILTER
        │
        ▼
Remaining Records
        │
        ▼
Transformation Rules
        │
        ▼
M3 SDT Output
```

This means that records excluded by a FILTER rule are not processed by the remaining transformation rules.

---

## Basic FILTER Structure

A FILTER rule typically contains:

```text
RULE_TYPE    = FILTER
SOURCE_FIELD = <source column>
RULE_VALUE   = <condition>
```

The condition determines whether a row should remain in the migration dataset.

---

## Simple Example

Assume only active records should be migrated.

Source:

```text
MBSTAT
```

Possible values:

```text
20
90
```

Business requirement:

```text
Keep rows where MBSTAT = 20
```

FILTER rule:

```python
source == "20"
```

If:

```text
MBSTAT = 20
```

the row is kept.

If:

```text
MBSTAT = 90
```

the row is excluded.

---

## Using the Full Row

A FILTER rule can also evaluate other source fields through `row`.

Example:

```python
str(row.get("MMITTY", "")).strip().upper() != "ZZ"
```

This keeps rows where:

```text
MMITTY <> ZZ
```

---

## Multiple Conditions

Filters can evaluate multiple conditions.

Example:

```python
status = str(row.get("MBSTAT", "")).strip()
itty = str(row.get("MMITTY", "")).strip().upper()

status == "20" and itty != "ZZ"
```

This means:

```text
Keep the row when:

MBSTAT = 20

AND

MMITTY <> ZZ
```

---

## Keep Logic vs Exclude Logic

FILTER conditions represent rows that should be kept.

For example:

```python
source == "20"
```

means:

> Keep rows where the source value equals 20.

It does not mean:

> Delete rows where the source equals 20.

This distinction is important when writing filter logic.

---

## Example

Source data:

| ITNO | STAT |
| --- | --- |
| A100 | 20 |
| A200 | 90 |
| A300 | 20 |
| A400 | 90 |

FILTER:

```python
source == "20"
```

Result:

| ITNO | STAT |
| --- | --- |
| A100 | 20 |
| A300 | 20 |

Only the records matching the condition continue into the transformation process.

---

## Multiple FILTER Rules

If multiple FILTER rules exist, each filter is applied to the remaining dataset.

Conceptually:

```text
Original Dataset
      │
      ▼
Filter 1
      │
      ▼
Remaining Rows
      │
      ▼
Filter 2
      │
      ▼
Final Migration Dataset
```

Therefore, filters should be reviewed carefully because each additional filter can reduce the number of migrated records.

---

## Migration Stops When Everything Is Filtered

If FILTER rules remove every record, the migration is aborted.

The System Log may report that all rows were filtered out.

This protects the tool from generating an empty migration file unintentionally.

---

## Recommended Practice

Before creating a FILTER rule:

1. Define the business population that should be migrated.
2. Identify the source fields required to determine eligibility.
3. Write the condition as a **keep condition**.
4. Test the filter against known records.
5. Compare the row count before and after filtering.
6. Review excluded records before production migration.

---

## Validation

Suppose the source contains:

```text
10,000 records
```

After filtering:

```text
8,450 records
```

The migration team should understand why:

```text
1,550 records
```

were excluded.

A successful migration should not only generate a file. The resulting record population should also match the approved migration scope.

---

## Troubleshooting

If too many records are excluded, verify:

- The correct source field is being used.
- String and numeric comparisons are correct.
- Blank values are handled correctly.
- Multiple filters are not unintentionally removing the same population.
- The condition represents rows to **keep**, not rows to delete.

For example:

```python
source != "90"
```

keeps everything except `90`.

While:

```python
source == "90"
```

keeps only `90`.

These two filters produce completely different migration populations.