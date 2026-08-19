# MCO Import

The **MCO Import** function is used to create or update migration configuration from an MCO workbook.

The MCO provides the field-level information required to establish the relationship between legacy source fields and the corresponding M3 target fields.

---

## Purpose

When preparing a new migration object, the MCO acts as one of the primary specifications for the migration.

The import process helps convert information from the MCO into configuration that can be used by the Data Migration Tool.

Conceptually:

```text
MCO Workbook
     │
     ▼
MCO Import
     │
     ▼
Analyze MCO Fields
     │
     ▼
Migration Configuration
     │
     ├── Target Fields
     ├── Source Fields
     └── Initial Rule Information
```

This reduces the need to manually create the initial Rule Configuration from scratch.

---

# Accessing MCO Import

From the application, navigate to:

**Configuration → Import MCO**

Select the MCO workbook that contains the migration specification.

---

# MCO Workbook

An MCO workbook may contain multiple worksheets representing different migration objects.

For example:

```text
MCO Workbook
│
├── Item Master
├── Item Facility
├── Item Warehouse
├── Supplier
└── Customer
```

Each worksheet can contain information about the source and target fields required for that migration object.

---

# Field Identification

During MCO processing, the tool looks for information identifying the M3 target field and the corresponding legacy source field.

For example:

```text
M3 Field       Source
--------       --------
ITNO           MBITNO
ITTY           MMITTY
PUIT           MBPUIT
```

This information provides the foundation for creating transformation rules.

---

# Source Field Prefixes

Legacy Movex fields frequently contain prefixes that identify the source table or business object.

For example:

```text
MBITNO
MMITTY
OKCUNO
IDSUNO
```

The application can analyze these field names and learn source prefixes.

Example:

```text
MB → Item-related source
MM → Item-related source
OK → Customer-related source
```

These signatures can later assist features such as **Auto-Detect Migration**.

---

# API Identification

When API information is available in the MCO, the tool can identify M3 API names using patterns such as:

```text
MMS200MI
CRS610MI
CRS620MI
```

This information helps establish the relationship:

```text
MCO Business Object
        │
        ▼
      M3 API
        │
        ▼
Rule Configuration
```

---

# Rule Configuration

The MCO provides the starting point for creating transformation rules.

However, importing the MCO does **not mean that all migration rules are complete**.

Some fields may require:

```text
DIRECT
CONST
MAP
PYTHON
FILTER
```

depending on the business requirement.

For example, an MCO may establish:

```text
TARGET_FIELD = ITNO
SOURCE_FIELD = MBITNO
```

A simple field may then use:

```text
RULE_TYPE = DIRECT
```

But another field such as:

```text
PLCD
```

may depend on several source fields and require a `PYTHON` rule.

---

# MCO Import vs Rule Development

These are two separate activities.

```text
MCO Import
    │
    ▼
Initial Configuration
    │
    ▼
Rule Review
    │
    ▼
Business Logic Development
    │
    ▼
Testing
    │
    ▼
Approved Rule Configuration
```

!!! important
    MCO Import should be considered the starting point of Rule Configuration, not the end of rule development.

Rules must still be reviewed against the approved migration requirements.

---

# After Importing an MCO

After the import process is complete:

1. Review the generated migration configuration.
2. Verify the M3 target fields.
3. Verify the identified source fields.
4. Confirm the associated M3 API.
5. Review the generated Rule Configuration.
6. Identify fields requiring additional business logic.
7. Create any required MAP or PYTHON rules.
8. Configure the source and migration mappings.
9. Run a test migration.
10. Validate the generated SDT output.

---

# Example

Suppose the MCO contains:

```text
M3 FIELD     SOURCE
--------     -------
ITNO         MBITNO
PUIT         MBPUIT
```

The initial configuration may establish:

```text
ITNO ← MBITNO
PUIT ← MBPUIT
```

For `ITNO`, a direct transformation may be sufficient:

```text
TARGET_FIELD = ITNO
SOURCE_FIELD = MBITNO
RULE_TYPE    = DIRECT
```

For another target field, the MCO information may only provide part of what is required.

The migration team must then add the appropriate business transformation logic.

---

# Relationship with Auto-Detect

The Data Migration Tool can also analyze MCO worksheets to learn field signatures.

For example:

```text
MMITTY
```

produces the prefix:

```text
MM
```

When a legacy file is later analyzed, the Auto-Detect process can compare its headers against known MCO prefixes.

Conceptually:

```text
MCO
 │
 ▼
Learn Prefixes
 │
 ▼
MM → Item Master
OK → Customer
ID → Supplier
 │
 ▼
Legacy File Headers
 │
 ▼
Identify Migration Object
```

This allows the application to assist with identifying the appropriate migration configuration automatically.

---

# Validation

After MCO Import, verify that:

- The correct MCO worksheet was processed.
- Target fields correspond to the expected M3 fields.
- Source fields correspond to the correct legacy fields.
- The correct API is associated with the migration object.
- Rule Configuration was created correctly.
- Fields requiring custom business logic have been identified.

The MCO provides the migration specification, while the Rule Configuration defines how that specification is executed by the Data Migration Tool.