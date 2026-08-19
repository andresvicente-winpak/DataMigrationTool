# Python Rules

A `PYTHON` rule is used when a target field requires transformation logic that cannot be handled by a simple `DIRECT`, `CONST`, or `MAP` rule.

Python rules allow the Data Migration Tool to evaluate values from the complete source row and apply custom business logic before generating the M3 target value.

---

## When to Use a Python Rule

Use a Python rule when the target value depends on:

- Multiple source fields
- Conditional logic
- Multiple business conditions
- String manipulation
- Numeric calculations
- Blank or null handling
- A combination of the above

For example:

```text
ITNO ──┐
       │
ITTY ──┼──► Python Rule ──► PLCD
       │
PUIT ──┘
```

In this case, `PLCD` cannot be determined from only one source field.

---

## Basic Structure

A Python rule normally reads values, evaluates the business conditions, and returns the resulting target value.

Example:

```python
puit = str(row.get("PUIT", row.get("MBPUIT", ""))).strip()

if puit == "1":
    return "00"

if puit == "2":
    return "11"

if puit == "3":
    return "12"

return ""
```

The final `return` determines the value written to the target field.

---

## Available Variables

### `source`

`source` represents the value obtained from the field defined in `SOURCE_FIELD`.

Example:

```python
if source is None:
    return ""

return str(source).strip()
```

For simple transformations, using `source` may be sufficient.

---

### `row`

`row` provides access to the complete source record.

For example:

```python
itno = row.get("MBITNO", "")
puit = row.get("MBPUIT", "")
```

This makes it possible for one target rule to depend on multiple source columns.

Example:

```python
itty = str(row.get("MMITTY", "")).strip().upper()
puit = str(row.get("MBPUIT", "")).strip()

if itty == "FG" and puit == "1":
    return "02"

return ""
```

---

## Handling Alternative Column Names

Source files may not always use the same column naming convention.

For example, a field could appear as:

```text
PUIT
```

or:

```text
MBPUIT
```

A Python rule can support both:

```python
puit = str(
    row.get("PUIT", row.get("MBPUIT", ""))
).strip()
```

The rule first searches for `PUIT`.

If it does not exist, it searches for `MBPUIT`.

If neither exists, it returns an empty string.

This pattern can also be used for other fields:

```python
itno = str(
    row.get("ITNO", row.get("MBITNO", ""))
).strip().upper()

itty = str(
    row.get("ITTY", row.get("MMITTY", ""))
).strip().upper()
```

---

## Normalize Values Before Comparing

Values should normally be cleaned before conditions are evaluated.

For text fields:

```python
value = str(row.get("FIELD", "")).strip().upper()
```

This performs three operations:

```text
Convert to string
       ↓
Remove spaces
       ↓
Convert to uppercase
```

For example:

```text
" fg "
```

becomes:

```text
"FG"
```

This helps prevent conditions from failing because of capitalization or unwanted spaces.

---

## Example: PLCD

Consider the following business requirement:

```text
If ITNO starts with RS9 or ZZ
    → 00

If ITTY is RM or PK and PUIT = 2
    → 01

If ITTY is SF or FG and PUIT = 1
    → 02

If PUIT = 3
    → 05
```

The Python rule can be written as:

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

---

## Condition Order Matters

Python rules are evaluated from top to bottom.

Consider:

```python
if puit == "3":
    return "05"

if itty == "FG" and puit == "3":
    return "02"
```

If:

```text
ITTY = FG
PUIT = 3
```

the result will be:

```text
05
```

The second condition will never be evaluated because the first condition already returned a value.

Therefore, condition priority must match the business requirement.

!!! warning
    When multiple conditions can match the same record, place the highest-priority condition first.

---

## Blank Values

Python rules should explicitly consider blank values when necessary.

Example:

```python
value = str(source).strip() if source is not None else ""

if value == "":
    return ""
```

This prevents blank values from accidentally being interpreted as meaningful data.

---

## Returning the Original Source Value

Sometimes a transformation should apply only under certain conditions and otherwise preserve the source value.

Example:

```python
value = str(source).strip() if source is not None else ""

if value == "PQ":
    return "QC"

return value
```

This means:

```text
PQ → QC
ON → ON
MB → MB
```

---

## Returning a Constant from a Condition

Python rules can also return constants based on another field.

Example requirement:

```text
If PUIT = 1 → NOR
If PUIT = 2 → INV
If PUIT = 3 → D22
```

Rule:

```python
puit = str(
    row.get("PUIT", row.get("MBPUIT", ""))
).strip()

if puit == "1":
    return "NOR"

if puit == "2":
    return "INV"

if puit == "3":
    return "D22"

return ""
```

---

## Multiple Conditions

Conditions can be combined using `and`.

Example:

```python
itty = str(row.get("MMITTY", "")).strip().upper()
eoqm = str(row.get("MBEOQM", "")).strip()

if itty == "FG" and eoqm == "12":
    return "010"

return "020"
```

This represents:

```text
ITTY = FG
   AND
EOQM = 12
    │
    ▼
   010

Otherwise
    │
    ▼
   020
```

---

## Testing a Python Rule

A Python rule should always be tested using known source records.

For example:

| ITNO | ITTY | PUIT | Expected PLCD |
| --- | --- | ---: | --- |
| RS90001 | RM | 2 | 00 |
| AP32508 | FG | 1 | 02 |
| TEST01 | PK | 2 | 01 |
| TEST02 | FG | 3 | 05 |

After running the migration, compare the generated output against the expected results.

---

## Troubleshooting

When a Python rule produces an unexpected value, verify the following:

### 1. Source column name

Confirm that the rule is reading the correct source column.

For example:

```text
MMITTY
```

is different from:

```text
MBITTY
```

A rule referencing the wrong column may return a blank value without producing an obvious error.

---

### 2. Actual source value

Check the raw source record.

Do not assume the value in the generated output represents what existed in the source.

---

### 3. Data type

Values loaded from Excel may be represented differently than expected.

For migration codes, converting values to strings before comparison is generally recommended.

```python
puit = str(row.get("MBPUIT", "")).strip()
```

Then compare:

```python
if puit == "3":
```

rather than:

```python
if puit == 3:
```

---

### 4. Condition order

Verify that an earlier condition is not returning a value before the intended condition is reached.

---

### 5. Dependencies

Identify every source field required by the rule.

For example:

```text
PLCD
 │
 ├── ITNO
 ├── ITTY
 └── PUIT
```

When troubleshooting `PLCD`, all three source values should be reviewed.

---

## Recommended Practice

Keep Python rules focused on the business requirement.

Prefer:

```python
puit = str(row.get("MBPUIT", "")).strip()

if puit == "1":
    return "00"

if puit == "2":
    return "11"

if puit == "3":
    return "12"

return ""
```

over unnecessarily complex logic.

A rule should be understandable by another migration team member who needs to maintain it later.

---

## Next Steps

See:

- [Understanding the Rule Configuration](rule-file.md)
- [MAP Rules](map.md)
- [FILTER Rules](filter.md)
- Rule Troubleshooting