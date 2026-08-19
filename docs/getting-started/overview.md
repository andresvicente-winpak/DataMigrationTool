# Data Migration Tool Overview

The **Data Migration Tool** is a Python-based application designed to support the transformation of legacy data into formats required for migration into **Infor M3**.

The tool provides a graphical interface for configuring migration processes, applying transformation rules, generating M3 SDT files, and managing migration configurations.

---

## Purpose

During a data migration, source data does not always match the structure or values expected by M3.

The Data Migration Tool provides a configurable transformation layer between the **legacy source data** and the **M3 target format**.

A simplified migration flow is:

```text
Legacy Source Data
        │
        ▼
Rule Configuration
        │
        ▼
Transformation Engine
        │
        ▼
SDT Template
        │
        ▼
M3 Load File
```

Instead of hard-coding every migration transformation into the application, transformation logic can be maintained through **Rule Configuration files**.

---

## Main Application Areas

The application is organized into five main areas.

### Run Migration

The **Run Migration** area is used to execute migration processes.

The available migration methods include:

- Standard Migration
- Auto-Detect
- Load by ID
- Batch Migration

Each method provides a different way of selecting and processing source data.

---

### Configuration

The **Configuration** area contains tools used to prepare and maintain the migration environment.

Configuration includes items such as:

- MCO specifications
- Source mappings
- Migration mappings
- Business units
- Database configuration
- Rule configuration generation

---

### Rules & Admin

The **Rules & Admin** area is used to maintain the transformation logic used during migration.

Rule configurations determine how source fields are converted into M3 target fields.

For example:

```text
Source Data

MBITNO = AP32508
MBPUIT = 1
MMITTY = FG

        │
        ▼

Transformation Rules

        │
        ▼

M3 Output

ITNO = AP32508
PUIT = 1
PLCD = 02
```

Rules may perform direct mappings, assign constants, perform lookups, apply filters, or execute custom Python logic.

---

### Utilities

The **Utilities** area provides supporting tools for migration activities.

Current utilities include:

- Copy Sheet
- Merge Files
- Script Library

These tools can be used independently from the main migration process.

---

### Sync / Merge

The **Sync / Merge** area allows Rule Configuration changes from another user or file to be compared with the local configuration.

Detected differences can be reviewed before selected changes are merged into the local Rule Configuration.

---

## Migration Components

Several components work together during a migration.

### Source Data

The source data contains the legacy information that needs to be migrated.

Sources may include Excel or CSV files, and configured processes may also use SQL data sources.

### Rule Configuration

The Rule Configuration defines how source data should be transformed.

A rule normally identifies:

- Target field
- Source field
- Rule type
- Rule value or transformation logic
- Scope

### SDT Template

The SDT template defines the M3 structure that must be populated.

The tool applies the configured transformation rules and writes the resulting values into the appropriate SDT transaction sheets.

### Output

The result of the migration process is an M3 load file generated in the required SDT structure.

---

## Next Step

Continue to [Standard Migration](../migration/standard-migration.md) to learn how to execute a standard migration.