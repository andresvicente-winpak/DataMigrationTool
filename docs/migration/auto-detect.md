# Auto-Detect Migration

The **Auto-Detect Migration** function analyzes a legacy source file and attempts to determine which migration business object and configuration should be used.

This is useful when the user has a source file but does not already know which MCO migration object it belongs to.

---

## Purpose

In Standard Migration, the user already knows the migration configuration:

```text
User
 │
 ├── Select Rule Configuration
 ├── Select Source
 └── Select Scope
```

Auto-Detect changes the beginning of this process.

Instead of manually identifying the migration object, the application analyzes the source file structure.

```text
Legacy Source File
        │
        ▼
Analyze Column Headers
        │
        ▼
Compare with Known MCO Signatures
        │
        ▼
Identify Migration Object
        │
        ▼
Resolve Migration Configuration
        │
        ▼
Run Migration
```

---

# How Auto-Detect Works

The Data Migration Tool learns source-field signatures from the configured MCO workbook.

Legacy fields frequently contain prefixes identifying their source table.

For example:

```text
MBITNO
MBPUIT
MBSTAT
```

share the prefix:

```text
MB
```

Another source may contain:

```text
MMITNO
MMITTY
```

with the prefix:

```text
MM
```

The Auto-Detect process uses these patterns to help identify the migration object.

---

# MCO Signatures

During MCO analysis, the application examines source-field definitions and extracts recognizable prefixes.

Conceptually:

```text
MCO
 │
 ├── Item Master
 │      ├── MBITNO
 │      ├── MBPUIT
 │      └── MBSTAT
 │
 ├── Item Facility
 │      ├── MMITNO
 │      └── MMITTY
 │
 ▼
Known Signatures

Item Master    → MB
Item Facility  → MM
```

These signatures can then be compared against an unknown legacy source file.

---

# Source File Analysis

Suppose a source file contains:

```text
MBITNO
MBITTY
MBPUIT
MBSTAT
MBCONO
```

The application analyzes the column headers and extracts their prefixes.

In this example:

```text
MB
```

appears repeatedly.

The Auto-Detect process compares that information with the known MCO signatures.

```text
Source File
    │
    ▼
MBITNO
MBITTY
MBPUIT
MBSTAT
    │
    ▼
Detected Prefix
    │
    ▼
MB
    │
    ▼
Known MCO Signatures
    │
    ▼
Potential Migration Match
```

---

# API Detection

The MCO analysis can also identify M3 API names when they are available in the MCO specification.

Examples include:

```text
MMS200MI
CRS610MI
CRS620MI
```

This information helps connect the detected business object with the appropriate migration configuration.

---

# Running Auto-Detect

Navigate to:

**Run Migration → Auto-Detect**

Select the legacy source file that should be analyzed.

The application will inspect the file headers and compare them against the known migration signatures.

---

# Detection Result

When a suitable match is identified, the application can use the detected MCO context to resolve the corresponding migration configuration.

Conceptually:

```text
Legacy File
    │
    ▼
Auto-Detect
    │
    ▼
MCO Business Object
    │
    ▼
Migration Map
    │
    ├── API
    ├── SDT Template
    └── Transaction Sheets
    │
    ▼
Rule Configuration
```

The normal migration process can then continue.

---

# Auto-Detect Does Not Replace Rules

Auto-Detect determines **which migration configuration should be used**.

It does not determine the business transformation logic for individual target fields.

For example:

```text
Auto-Detect
    │
    ▼
"This appears to be Item Master"
    │
    ▼
MMS200MI Rule Configuration
    │
    ▼
DIRECT / CONST / MAP / PYTHON Rules
```

The Rule Configuration remains responsible for transforming the source data.

!!! important
    Auto-Detect helps identify the migration object. It does not replace a complete and validated Rule Configuration.

---

# Example

Suppose the user receives:

```text
Legacy_Item_Data.xlsx
```

The user may not know which configured migration should process the file.

The file contains:

```text
MBITNO
MBITTY
MBPUIT
MBSTAT
```

Auto-Detect analyzes the headers:

```text
MBITNO ─┐
MBITTY ─┤
MBPUIT ─┼──► MB
MBSTAT ─┘
```

The application compares:

```text
MB
```

against its known MCO signatures.

If a matching migration object is identified, that information can be used to determine the corresponding API and migration configuration.

---

# Why Auto-Detect Is Useful

Auto-Detect is particularly useful when:

- The source filename does not clearly identify the migration object.
- Multiple migration files are being reviewed.
- The user is unfamiliar with the legacy source tables.
- Source files use recognizable Movex field prefixes.
- The migration configuration has already been established.

Instead of relying only on filenames, the application analyzes the actual source structure.

---

# Detection Depends on Configuration

Auto-Detect relies on information learned from the MCO configuration.

Therefore, detection quality depends on the available MCO signatures.

If a migration object has not been configured correctly, Auto-Detect may not be able to identify it.

```text
Good MCO Configuration
        │
        ▼
Good Signatures
        │
        ▼
Better Detection
```

---

# Ambiguous Matches

Some source files may contain fields associated with multiple source tables.

For example:

```text
MBITNO
MBPUIT
MMITNO
MMITTY
```

In this situation, more than one migration object may appear relevant.

Auto-Detect should therefore be treated as an identification aid rather than a replacement for migration validation.

Always confirm that the detected migration object matches the intended business requirement.

---

# Validation Before Migration

After Auto-Detect identifies a migration configuration, verify:

1. The detected MCO business object is correct.
2. The expected API is selected.
3. The correct Rule Configuration exists.
4. The correct SDT template is configured.
5. The source fields required by the rules are present.
6. The appropriate migration scope is selected.
7. The generated output is validated.

---

# Troubleshooting

If Auto-Detect cannot identify the source, verify:

- The source file contains column headers.
- The headers correspond to fields defined in the MCO.
- The MCO has been imported/configured.
- Source-field prefixes exist in the MCO specification.
- The migration object has a valid configuration.
- The source file is actually associated with a configured migration.

If the wrong migration object is detected, compare the source headers against the signatures of the possible MCO objects.

---

# Standard vs Auto-Detect

The main difference is how the migration configuration is identified.

```text
STANDARD MIGRATION

User selects configuration
        │
        ▼
Migration
```

```text
AUTO-DETECT MIGRATION

Source File
    │
    ▼
Analyze Headers
    │
    ▼
Identify Configuration
    │
    ▼
Migration
```

After the configuration has been identified, the same core migration concepts still apply:

```text
Source
  +
Configuration
  +
Rules
  +
SDT Template
  │
  ▼
M3 Output
```

---

# Recommended Practice

Use Auto-Detect to assist with identifying unknown source files, but always verify the detected migration configuration before using the resulting file for an M3 load.

A successful detection means:

> The source structure appears to match a known migration object.

It does not automatically mean:

> The source data and transformation rules have been fully validated.

---

# Next Step

Continue to **Load by ID** to learn how a migration can retrieve and process specific records using identifiers instead of processing an entire source dataset.