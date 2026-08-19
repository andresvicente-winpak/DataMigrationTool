# Configuration Overview

The **Configuration** area defines how the Data Migration Tool connects legacy source data with the appropriate M3 migration configuration.

Configuration provides the information the application needs to determine:

- Where the source data comes from
- Which MCO business object is being migrated
- Which M3 API is associated with the migration
- Which Rule Configuration should be used
- Which SDT template should receive the transformed data
- Which transaction sheets are required

---

## Configuration in the Migration Process

Configuration connects the major components of a migration.

```text
Legacy Source
      │
      ▼
Source Configuration
      │
      ▼
MCO / Business Object
      │
      ├─────────────► Rule Configuration
      │
      ▼
M3 API
      │
      ▼
SDT Template
      │
      ▼
Transaction Sheet(s)
      │
      ▼
Migration Output
```

Without the correct configuration, the application may not be able to determine which source, rules, API, or SDT template should be used.

---

# Main Configuration Components

The Data Migration Tool uses several configuration components.

## MCO Specification

The MCO specification provides information about the migration business objects and their source-to-target field relationships.

The tool can analyze an MCO workbook and use that information to help create migration configurations.

---

## Source Map

The Source Map defines where the legacy data for a migration comes from.

It connects an MCO sheet or business object with its corresponding source.

The source may be:

- Excel
- CSV
- SQL

Conceptually:

```text
MCO Sheet
    │
    ▼
Source Map
    │
    ▼
Legacy Source
```

---

## Migration Map

The Migration Map connects the migration business object with the M3 configuration required to generate the output.

It can identify information such as:

```text
MCO Sheet
    │
    ├──► API
    ├──► SDT Template
    └──► Transaction Sheet(s)
```

This allows the Migration Runner to automatically determine which SDT structure should be used.

---

## Rule Configuration

The Rule Configuration contains the transformation logic.

For example:

```text
Legacy Source
      │
      ▼
Rule Configuration
      │
      ▼
M3 Target Values
```

Rule Configuration files are normally stored under:

```text
config/rules/
```

See [Understanding the Rule Configuration](../rules/rule-file.md) for more information.

---

## SDT Templates

SDT templates define the M3 output structure.

Templates are normally stored under:

```text
config/sdt_templates/
```

The Migration Map determines which template should be associated with a migration.

---

## Business Units and Scope

Business-unit configuration allows transformation rules to vary when required.

The default rule scope is:

```text
GLOBAL
```

Additional scopes can be configured for business-specific requirements.

For example:

```text
GLOBAL
DIV_US
DIV_CA
```

---

## Database Configuration

The Data Migration Tool can also use SQL as a source.

Database connection information is maintained separately from the transformation rules.

A SQL source can then be used by the extraction process instead of an Excel or CSV file.

---

# Why Configuration Matters

Consider a migration for Item Master.

The tool needs to understand the complete relationship:

```text
Item Master
    │
    ├── Where is the source data?
    │
    ├── Which M3 API applies?
    │
    ├── Which rules should be used?
    │
    ├── Which SDT template applies?
    │
    └── Which transaction sheets should be populated?
```

Configuration provides these relationships so they do not need to be manually specified every time a migration is executed.

---

# Configuration vs Rules

Configuration and Rules serve different purposes.

| Configuration | Rules |
| --- | --- |
| Determines **what migration setup to use** | Determines **how data is transformed** |
| Identifies source | Reads source values |
| Identifies API | Calculates M3 values |
| Identifies SDT template | Applies business logic |
| Identifies transaction sheets | Populates target fields |

Both are required for a successful migration.

```text
Configuration
      +
Rules
      +
Source Data
      │
      ▼
Migration Engine
      │
      ▼
M3 SDT Output
```

---

# Recommended Configuration Workflow

When preparing a new migration object:

1. Review the MCO specification.
2. Identify the legacy source.
3. Configure the Source Map.
4. Identify the M3 API.
5. Configure the Migration Map.
6. Verify the SDT template.
7. Create or import the Rule Configuration.
8. Review transformation rules.
9. Test the migration.
10. Validate the generated output.

---

# Next Steps

Continue with:

- MCO Import
- Source Map
- Migration Map
- Database Configuration