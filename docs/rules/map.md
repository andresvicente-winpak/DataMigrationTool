# MAP Rules

A `MAP` rule is used when a source value must be translated into a different target value using a lookup table.

This is useful when legacy codes do not match the values expected by M3.

---

## When to Use a MAP Rule

Use a MAP rule when:

- The source value must be converted into another code
- There are many value-to-value translations
- The mapping is easier to maintain in a table than in Python logic
- Business users may need to review or update the mapping

Example:

```text
Legacy Value    M3 Value
------------    --------
A               01
B               02
C               03
```

If the legacy source contains:

```text
A
```

the MAP rule returns:

```text
01
```

---

## Basic MAP Structure

A typical MAP rule contains:

```text
TARGET_FIELD = <M3 field>
SOURCE_FIELD = <legacy source field>
RULE_TYPE    = MAP
RULE_VALUE   = <mapping configuration>
```

The `SOURCE_FIELD` provides the value that must be translated.

The `RULE_VALUE` identifies the lookup definition used by the transformation engine.

---

## External Mapping Files

The Data Migration Tool can load mapping values from an external Excel or CSV file.

The mapping configuration uses this format:

```text
path|key_column|value_column
```

Example:

```text
config/maps/item_status.xlsx|LEGACY_STATUS|M3_STATUS
```

This means:

```text
File:
config/maps/item_status.xlsx

Lookup key:
LEGACY_STATUS

Returned value:
M3_STATUS
```

---

## Example Mapping File

Example Excel file:

| LEGACY_STATUS | M3_STATUS |
| --- | --- |
| A | 10 |
| B | 20 |
| C | 30 |

Rule:

```text
TARGET_FIELD = STAT
SOURCE_FIELD = MBSTAT
RULE_TYPE    = MAP
RULE_VALUE   = config/maps/item_status.xlsx|LEGACY_STATUS|M3_STATUS
```

Source:

```text
MBSTAT = B
```

Result:

```text
STAT = 20
```

---

## How MAP Processing Works

The transformation flow is:

```text
SOURCE_FIELD
     │
     ▼
Normalize Source Value
     │
     ▼
Mapping Table
     │
     ▼
TARGET Value
```

The tool normalizes the source lookup value before matching.

For example:

```text
" a "
```

is treated as:

```text
"A"
```

This helps reduce problems caused by spaces or capitalization differences.

---

## Duplicate Mapping Keys

The mapping file should contain one target value per source key.

Avoid configurations like:

```text
SOURCE_KEY    TARGET_VALUE
----------    ------------
A             01
A             02
```

A duplicate key creates ambiguity.

The mapping file should instead contain:

```text
A → 01
```

or:

```text
A → 02
```

depending on the approved business requirement.

---

## Missing Mapping Values

If the source value does not exist in the lookup table, the result may be blank.

For example:

```text
Source:

MBSTAT = X
```

Mapping:

```text
A → 10
B → 20
C → 30
```

There is no mapping for:

```text
X
```

Therefore the target value cannot be resolved.

!!! warning
    Before using a MAP rule in production migration, verify that all expected source values exist in the mapping table.

---

## When MAP is Better Than PYTHON

Consider this requirement:

```text
A → 10
B → 20
C → 30
D → 40
E → 50
F → 60
```

This could be written in Python:

```python
if source == "A":
    return "10"
if source == "B":
    return "20"
if source == "C":
    return "30"
```

However, this becomes difficult to maintain when many values are involved.

A MAP rule is preferable because the values can be managed in a table.

Use:

```text
MAP
```

for lookup-style transformations.

Use:

```text
PYTHON
```

when the target depends on conditions or multiple fields.

---

## Recommended Practice

When creating a MAP rule:

1. Identify the source field.
2. List all known source values.
3. Define the corresponding M3 values.
4. Store the mapping in a controlled Excel or CSV file.
5. Configure the MAP rule.
6. Test every expected source value.
7. Review unmapped values before migration.

---

## Troubleshooting

If a MAP rule returns blank or incorrect values, verify:

- The mapping file exists.
- The file path is correct.
- The key column exists.
- The value column exists.
- The source field contains the expected values.
- The mapping table does not contain unexpected duplicates.
- The source value exists in the mapping table.

---

## Example Validation

Source data:

| MBSTAT |
| --- |
| A |
| B |
| C |

Expected M3 output:

| STAT |
| --- |
| 10 |
| 20 |
| 30 |

The generated migration output should be compared against this expected mapping before the file is accepted.