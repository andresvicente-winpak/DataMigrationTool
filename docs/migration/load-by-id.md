# Load by ID

The **Load by ID** migration mode is used when only specific records need to be migrated instead of processing the entire source dataset.

This is useful for testing, troubleshooting, validation, and targeted migrations.

---

## Purpose

A Standard Migration typically processes the complete configured source:

```text
Complete Source Dataset
        │
        ▼
Transformation
        │
        ▼
M3 Output
```

Load by ID adds a record-selection step:

```text
Complete Source Dataset
        │
        ▼
Select Specific ID(s)
        │
        ▼
Matching Records
        │
        ▼
Transformation
        │
        ▼
M3 Output
```

This allows a smaller population to be processed without changing the original source.

---

# When to Use Load by ID

Load by ID is particularly useful for:

- Testing transformation rules
- Troubleshooting a specific record
- Validating a new rule
- Comparing expected and actual values
- Reprocessing selected records
- Creating a small migration sample
- Investigating source-data issues

For example, instead of migrating thousands of Item Master records, you may want to test:

```text
AP32508
AP32508-2P
RS90001
```

---

# Accessing Load by ID

Navigate to:

**Run Migration → Load by ID**

The Load by ID screen allows the user to configure a targeted migration.

---

# Basic Workflow

A Load by ID migration follows this general process:

```text
Select Migration Configuration
        │
        ▼
Select Source
        │
        ▼
Define Record ID(s)
        │
        ▼
Retrieve Matching Records
        │
        ▼
Apply FILTER Rules
        │
        ▼
Apply Transformation Rules
        │
        ▼
Generate SDT Output
```

The transformation process itself remains based on the same Rule Configuration used by the other migration modes.

---

# Record Identification

The ID represents the source value used to locate the record or records that should be processed.

For Item Master, for example, the identifier may be:

```text
ITNO
```

or the corresponding legacy field:

```text
MBITNO
```

Example source:

| MBITNO | MBITTY | MBPUIT |
| --- | --- | --- |
| AP32508 | FG | 1 |
| AP32508-2P | FG | 3 |
| TEST001 | RM | 2 |

If the requested ID is:

```text
AP32508
```

the targeted dataset becomes:

| MBITNO | MBITTY | MBPUIT |
| --- | --- | --- |
| AP32508 | FG | 1 |

Only the selected record continues through the migration process.

---

# Multiple IDs

Load by ID can be useful when a small group of related records needs to be tested.

For example:

```text
AP32508
AP32508-2P
AP32508-3P
```

Conceptually:

```text
Requested IDs
     │
     ├── AP32508
     ├── AP32508-2P
     └── AP32508-3P
     │
     ▼
Source Lookup
     │
     ▼
Matching Records
     │
     ▼
Migration
```

This provides a controlled test population.

---

# Rule Processing

After the requested records are retrieved, the normal transformation rules are applied.

For example, consider:

```text
MBITNO = AP32508
MMITTY = FG
MBPUIT = 1
```

A PLCD Python rule may evaluate:

```python
itno = str(
    row.get("ITNO", row.get("MBITNO", ""))
).strip().upper()

itty = str(
    row.get("ITTY", row.get("MMITTY", ""))
).strip().upper()

puit = str(
    row.get("PUIT", row.get("MBPUIT", ""))
).strip()

if itno.startswith(("RS9", "ZZ")):
    return "00"

if itty in ("RM", "PK") and puit == "2":
    return "01"

if itty in ("SF", "FG") and puit == "1":
    return "02"

if puit == "3":
    return "05"

return ""
```

For:

```text
ITNO = AP32508
ITTY = FG
PUIT = 1
```

the expected result is:

```text
PLCD = 02
```

This makes Load by ID particularly useful when validating complex rules.

---

# Troubleshooting with Load by ID

Load by ID can help isolate transformation problems.

Instead of reviewing:

```text
50,000 source records
```

you can isolate:

```text
1 source record
```

and compare:

```text
Raw Source
    │
    ▼
Rule Dependencies
    │
    ▼
Business Requirement
    │
    ▼
Expected Result
    │
    ▼
Actual Output
```

---

# Example Troubleshooting Process

Suppose:

```text
ITNO = AP32508
```

produces an unexpected `PLCD`.

First review the complete source record:

```text
MBITNO = AP32508
MMITTY = FG
MBPUIT = 1
```

Then identify the rule dependencies:

```text
PLCD
 │
 ├── ITNO
 ├── ITTY
 └── PUIT
```

Determine the expected result:

```text
ITTY = FG
PUIT = 1

Expected:

PLCD = 02
```

Run only:

```text
AP32508
```

using Load by ID and compare the generated value.

This makes it easier to determine whether the problem comes from:

- Source data
- Source-field naming
- Rule logic
- Rule scope
- Condition priority
- Migration configuration

---

# Load by ID and Scope

The selected migration scope still matters.

For example:

```text
Record:
AP32508

GLOBAL
    ↓
PLCD = 02

DIV_US
    ↓
PLCD = 03
```

The same record may produce different results when a scope-specific override exists.

Always verify the selected scope when using Load by ID for rule testing.

---

# Load by ID vs FILTER

Load by ID and FILTER rules both reduce the migration population, but they serve different purposes.

| Load by ID | FILTER |
| --- | --- |
| User selects specific records | Business logic selects records |
| Usually temporary | Stored in Rule Configuration |
| Useful for testing | Used during normal migration |
| Targets known IDs | Evaluates conditions |

For example:

```text
Load by ID:

Process AP32508 only
```

versus:

```text
FILTER:

Process every record where MBSTAT = 20
```

---

# Load by ID vs Standard Migration

```text
STANDARD

Configured Source
      │
      ▼
All Applicable Records
      │
      ▼
Migration
```

```text
LOAD BY ID

Configured Source
      │
      ▼
Requested IDs
      │
      ▼
Matching Records
      │
      ▼
Migration
```

The main difference is the population being processed.

The transformation rules and M3 output requirements remain the same.

---

# Validation

After running Load by ID, compare the generated output with the original source record.

For each tested field:

```text
Source Value
     +
Business Requirement
     +
Transformation Rule
     │
     ▼
Expected M3 Value
     ↕
Actual M3 Value
```

This is especially useful when developing new `PYTHON` rules.

---

# Recommended Testing Workflow

When developing or changing a transformation rule:

1. Identify representative source records.
2. Include normal cases.
3. Include special cases.
4. Include blank/null cases when applicable.
5. Determine the expected result for each record.
6. Run the selected records using Load by ID.
7. Compare the generated output.
8. Correct the rule if necessary.
9. Repeat the test.
10. Run a larger migration only after the targeted tests pass.

For example:

| Test Record | Scenario | Expected |
| --- | --- | --- |
| AP32508 | FG / PUIT 1 | PLCD 02 |
| TEST001 | PK / PUIT 2 | PLCD 01 |
| RS90001 | RS9 Item | PLCD 00 |
| TEST003 | PUIT 3 | PLCD 05 |

---

# Important Consideration

A successful Load by ID run proves that the selected records were processed successfully.

It does not prove that every possible source-data scenario has been tested.

Use representative records covering all important business-rule conditions before approving a transformation.

---

# Troubleshooting

If a requested record is not found, verify:

- The correct source is selected.
- The correct identifier field is being used.
- The ID exists in the source.
- Leading or trailing spaces are not affecting the comparison.
- The source value has the expected format.
- The database/source connection is available.
- The correct migration configuration is selected.

If the record is found but produces incorrect output, review:

- Rule dependencies
- Source values
- Rule type
- Python condition order
- Selected scope
- Expected business requirement

---

# Next Step

Continue to **Batch Migration** to learn how multiple migration configurations can be processed as part of a larger migration run.