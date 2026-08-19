# Business Units and Rule Scope

The **Business Unit / Scope** functionality allows the Data Migration Tool to apply different transformation rules depending on the organization, division, or migration context being processed.

The default scope is:

```text
GLOBAL
```

GLOBAL rules define the standard transformation behavior.

When a business unit requires different logic, a more specific scope can override the GLOBAL rule.

---

## Why Scope Is Needed

Most transformation rules may be common across the organization.

For example:

```text
ITNO
MBITNO → ITNO
```

may work the same way everywhere.

However, some M3 fields may require different transformation logic depending on the business unit.

Instead of creating completely separate Rule Configuration files, the Data Migration Tool allows multiple versions of a rule to exist under different scopes.

Conceptually:

```text
                  TARGET_FIELD
                       │
                       ▼
                     PLCD
                       │
             ┌─────────┴─────────┐
             ▼                   ▼
          GLOBAL             DIVISION
           Rule                 Rule
                                 │
                                 ▼
                             Override
```

---

# GLOBAL Scope

`GLOBAL` represents the default transformation rule.

Example:

| TARGET_FIELD | SOURCE_FIELD | RULE_TYPE | SCOPE |
| --- | --- | --- | --- |
| ITNO | MBITNO | DIRECT | GLOBAL |
| PUIT | MBPUIT | DIRECT | GLOBAL |

When the migration is executed using:

```text
Scope = GLOBAL
```

the GLOBAL rules are used.

---

# Business-Specific Scope

A business-specific scope can provide different logic for a target field.

Example:

| TARGET_FIELD | RULE_TYPE | SCOPE |
| --- | --- | --- |
| PLCD | PYTHON | GLOBAL |
| PLCD | PYTHON | DIV_US |

In this example, two rules exist for:

```text
PLCD
```

The GLOBAL rule provides the default logic.

The `DIV_US` rule provides different logic when that scope is selected.

---

# How Overrides Work

When a specific scope is selected, the Data Migration Tool loads:

```text
GLOBAL Rules
      +
Selected Scope Rules
```

If both contain a rule for the same `TARGET_FIELD`, the specific scope takes precedence.

For example:

```text
GLOBAL
PLCD → Rule A

DIV_US
PLCD → Rule B
```

Running:

```text
Scope = GLOBAL
```

uses:

```text
PLCD → Rule A
```

Running:

```text
Scope = DIV_US
```

uses:

```text
PLCD → Rule B
```

Other GLOBAL fields without a `DIV_US` override continue using their GLOBAL rules.

---

# Example

Assume the Rule Configuration contains:

| TARGET_FIELD | RULE_TYPE | RULE_VALUE | SCOPE |
| --- | --- | --- | --- |
| CONO | CONST | 100 | GLOBAL |
| ITNO | DIRECT | | GLOBAL |
| PLCD | PYTHON | Global PLCD logic | GLOBAL |
| PLCD | PYTHON | US PLCD logic | DIV_US |

When running:

```text
Scope = DIV_US
```

the effective configuration becomes:

```text
CONO → GLOBAL rule
ITNO → GLOBAL rule
PLCD → DIV_US rule
```

Only `PLCD` is overridden.

---

# Selecting Scope During Migration

Navigate to:

**Run Migration → Standard**

The Standard Migration screen includes:

```text
3. Scope (Optional)
```

The default selection is:

```text
GLOBAL
```

Configured business units are added to the available options.

For example:

```text
GLOBAL
DIV_US
DIV_CA
```

Select the appropriate scope before clicking:

**RUN MIGRATION**

---

# Business Unit Configuration

Available business units are loaded from:

```text
config/business_units.csv
```

The application reads the configured `UNIT` values and adds them to the Scope dropdown.

Conceptually:

```text
business_units.csv
       │
       ▼
Configured Units
       │
       ▼
Scope Dropdown
```

---

# Scope and Rule Configuration

Scope is stored with the transformation rule.

Example:

```text
TARGET_FIELD = PLCD
SOURCE_FIELD = MBPUIT
RULE_TYPE    = PYTHON
SCOPE        = DIV_US
```

This means the rule applies when:

```text
DIV_US
```

is selected for the migration.

---

# When to Create an Override

Create a scope-specific rule only when the business requirement actually differs from the GLOBAL rule.

For example:

```text
GLOBAL Requirement

PUIT = 1 → PLCD = 02
```

Suppose one division requires:

```text
DIV_US Requirement

PUIT = 1 → PLCD = 03
```

Then an override is appropriate:

```text
PLCD / GLOBAL
        │
        └── Default logic

PLCD / DIV_US
        │
        └── Division-specific logic
```

---

# Avoid Duplicating GLOBAL Rules

Do not create scope-specific copies of every GLOBAL rule when the transformation is identical.

For example, if:

```text
ITNO = MBITNO
```

works globally, there is no need to create:

```text
ITNO / GLOBAL
ITNO / DIV_US
ITNO / DIV_CA
```

with identical logic.

Keep:

```text
ITNO / GLOBAL
```

and create overrides only where required.

This keeps the Rule Configuration easier to maintain.

---

# Scope Selection Is Important

Selecting the wrong scope can change the resulting migration values.

For example:

```text
Source
PUIT = 1

GLOBAL Rule
    ↓
PLCD = 02

DIV_US Rule
    ↓
PLCD = 03
```

The same source record can therefore produce different output depending on the selected scope.

!!! warning
    Always verify the selected migration scope before generating a production load file.

---

# Testing Scope Overrides

When creating an override, test both configurations.

For example:

### Test 1 — GLOBAL

```text
Scope: GLOBAL
Source PUIT: 1
Expected PLCD: 02
```

### Test 2 — DIV_US

```text
Scope: DIV_US
Source PUIT: 1
Expected PLCD: 03
```

This verifies that:

1. The GLOBAL rule still works.
2. The override is applied only to the intended scope.

---

# Recommended Practice

When using scopes:

1. Define the standard transformation as `GLOBAL`.
2. Identify genuine business-unit differences.
3. Create overrides only for those differences.
4. Keep the same `TARGET_FIELD` for the overridden rule.
5. Clearly document why the override exists.
6. Test both GLOBAL and specific-scope behavior.
7. Verify the selected scope before production migration.

A good configuration should look conceptually like:

```text
GLOBAL
│
├── ITNO
├── ITTY
├── PUIT
├── PLCD
├── STAT
└── ...

DIV_US
│
└── PLCD      ← Override only

DIV_CA
│
├── PLCD      ← Override
└── STAT      ← Override
```

Most rules remain GLOBAL while only genuine business differences are overridden.

---

# Troubleshooting

If a scope-specific rule is not being applied, verify:

- The business unit exists in `config/business_units.csv`.
- The correct scope was selected before running the migration.
- The Rule Configuration contains the expected `SCOPE`.
- The `TARGET_FIELD` matches the GLOBAL target field being overridden.
- The scope name matches exactly.
- The Rule Configuration was saved before running the migration.

If the migration behaves correctly under GLOBAL but incorrectly under another scope, compare the GLOBAL and scope-specific rules for the affected target field.

---

# Next Step

The next configuration topic is **Database Configuration**, which explains how SQL Server connections are configured when SQL is used as a migration source.