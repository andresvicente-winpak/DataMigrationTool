# Understanding the Rule Configuration

The **Rule Configuration** defines how legacy source data is transformed into the values required by the target M3 SDT.

It is one of the core components of the Data Migration Tool.

A migration may have a valid source file and the correct SDT template, but without the appropriate transformation rules, the target fields cannot be populated correctly.

---

## Role of Rules in a Migration

The migration process can be represented as:

```text
Legacy Source Data
        │
        ▼
┌──────────────────────┐
│  Rule Configuration  │
│                      │
│  TARGET_FIELD        │
│  SOURCE_FIELD        │
│  RULE_TYPE           │
│  RULE_VALUE          │
│  SCOPE               │
└──────────┬───────────┘
           │
           ▼
 Transformation Engine
           │
           ▼
      M3 SDT Output
```

The Rule Configuration tells the Transformation Engine:

- Which M3 field must be populated
- Where the source value comes from
- How the value must be transformed
- Whether special business logic is required
- Whether the rule applies globally or to a specific scope

---

# Rule Configuration Files

Rule Configuration files are stored in:

```text
config/rules/
```

Each M3 program or migration configuration can have its own rule file.

For example:

```text
config/rules/MMS200MI.xlsx
```

A Rule Configuration contains a `Rules` worksheet where the transformation logic is defined.

---

# Rule Structure

Each row in the Rules worksheet represents the transformation logic for a target field.

The main columns are:

| Column | Purpose |
| --- | --- |
| `TARGET_FIELD` | M3 destination field |
| `SOURCE_FIELD` | Field containing the legacy/source value |
| `RULE_TYPE` | Transformation method |
| `RULE_VALUE` | Additional logic or configuration |
| `SCOPE` | Determines where the rule applies |

---

## TARGET_FIELD

`TARGET_FIELD` identifies the field that will be populated in the M3 SDT.

For example:

```text
PLCD
```

The target field must correspond to a field available in the applicable SDT transaction sheet.

A simplified rule could therefore look like:

```text
TARGET_FIELD = PLCD
```

Meaning:

> This rule determines the value that will be written to the M3 field `PLCD`.

---

## SOURCE_FIELD

`SOURCE_FIELD` identifies the legacy field used as the primary source for the transformation.

Example:

```text
SOURCE_FIELD = MBPUIT
```

For a simple direct transformation:

```text
MBPUIT
   │
   ▼
Transformation Rule
   │
   ▼
PUIT
```

However, some rules require more than one source field.

For example, a `PLCD` rule may need information from:

```text
ITNO
ITTY
PUIT
```

In these situations, a Python rule can access additional values from the complete source row.

---

## RULE_TYPE

`RULE_TYPE` determines how the source value will be transformed.

The Data Migration Tool supports different transformation methods.

Common rule types include:

| Rule Type | Purpose |
| --- | --- |
| `DIRECT` | Copy a source value directly |
| `CONST` | Assign a fixed value |
| `MAP` | Translate a source value using a lookup |
| `PYTHON` | Execute custom transformation logic |
| `FILTER` | Determine which source records should be processed |

Each rule type is documented separately.

---

# DIRECT Rules

A `DIRECT` rule copies the value from the source field to the target field.

Example:

```text
TARGET_FIELD = ITNO
SOURCE_FIELD = MBITNO
RULE_TYPE    = DIRECT
```

Result:

```text
MBITNO = AP32508

        ↓

ITNO = AP32508
```

Use `DIRECT` when no transformation is required.

---

# CONST Rules

A `CONST` rule assigns the same value to the target field for every applicable record.

Example:

```text
TARGET_FIELD = CONO
RULE_TYPE    = CONST
RULE_VALUE   = 100
```

Result:

```text
CONO = 100
```

The source record does not need to contain the value.

---

# MAP Rules

A `MAP` rule converts a source value into another value using a predefined mapping.

Conceptually:

```text
Source Value
     │
     ▼
 Lookup Mapping
     │
     ▼
Target Value
```

For example:

```text
Legacy Value    M3 Value
------------    --------
A               01
B               02
C               03
```

If the source contains:

```text
A
```

the resulting target value becomes:

```text
01
```

---

# PYTHON Rules

`PYTHON` rules are used when the transformation requires business logic that cannot be represented by a simple direct mapping, constant, or lookup.

For example:

```python
itno = str(row.get("ITNO", row.get("MBITNO", ""))).strip().upper()
itty = str(row.get("ITTY", row.get("MMITTY", ""))).strip().upper()
puit = str(row.get("PUIT", row.get("MBPUIT", ""))).strip()

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

This type of rule can evaluate multiple fields from the same source row before determining the target value.

Python rules are particularly useful when the transformation contains:

- Multiple conditions
- IF/ELSE logic
- Dependencies on several source columns
- String manipulation
- Numeric calculations
- Business-specific transformation logic

!!! warning
    Python rules should be tested carefully. An incorrect condition may generate valid-looking output containing incorrect M3 values.

---

# FILTER Rules

A `FILTER` rule controls whether source records should participate in the migration.

Instead of calculating a target value, the rule determines whether a row should be included or excluded.

Conceptually:

```text
Source Records
      │
      ▼
  FILTER Rules
      │
      ├── Keep
      │
      └── Exclude
      │
      ▼
Transformation
```

Filters are applied before the remaining transformation rules are written to the output.

---

# Rule Scope

Rules can also be controlled using `SCOPE`.

The default scope is:

```text
GLOBAL
```

A GLOBAL rule represents the default transformation.

Additional scopes can be used when a business unit requires different transformation logic.

For example:

```text
TARGET_FIELD    SCOPE
------------    -------
PLCD            GLOBAL
PLCD            DIV_US
```

This allows the same target field to have different transformation logic depending on the selected migration scope.

---

# Why Rule Order and Dependencies Matter

Some target fields depend on other source values to determine their result.

For example:

```text
             ┌── ITNO
             │
Source Row ──┼── ITTY
             │
             └── PUIT
                  │
                  ▼
              PLCD Rule
                  │
                  ▼
                PLCD
```

Because of these dependencies, it is important to understand which source columns are required by each rule.

A rule should never assume that another target field has already been calculated unless the transformation process explicitly supports that dependency.

Whenever possible, rules should reference the original source row values.

---

# Example: PLCD

Consider the following source record:

```text
MBITNO  = AP32508
MMITTY  = FG
MBPUIT  = 1
```

A PLCD business rule may state:

```text
If ITNO starts with RS9 or ZZ
    → PLCD = 00

If ITTY is RM or PK and PUIT = 2
    → PLCD = 01

If ITTY is SF or FG and PUIT = 1
    → PLCD = 02

If PUIT = 3
    → PLCD = 05
```

For this record:

```text
ITTY = FG
PUIT = 1
```

Therefore:

```text
PLCD = 02
```

This demonstrates why some target fields cannot simply copy one source column.

---

# Recommended Rule Design

When creating a new rule:

1. Identify the M3 `TARGET_FIELD`.
2. Identify the primary `SOURCE_FIELD`.
3. Determine whether additional source fields are required.
4. Select the simplest appropriate `RULE_TYPE`.
5. Define the transformation logic.
6. Determine the appropriate `SCOPE`.
7. Test the rule against representative source records.
8. Validate the generated SDT output.

Whenever possible, prefer simpler rule types.

For example:

```text
DIRECT
   ↓
CONST
   ↓
MAP
   ↓
PYTHON
```

Use a Python rule when the business requirement actually requires conditional or custom logic.

---

# Rule Validation

After creating or modifying a rule, always validate the resulting migration output.

Do not validate only that the migration completed successfully.

A successful execution means that the application was able to process the migration. It does **not automatically mean that every transformed value is correct**.

Validation should compare:

```text
Source Data
     +
Business Requirement
     +
Rule Configuration
     ↓
Expected M3 Value
     ↕
Actual Generated Value
```

Any difference between the expected and actual value should be investigated before the generated file is used for migration.