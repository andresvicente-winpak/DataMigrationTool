# Batch Migration

The **Batch Migration** function is used to execute multiple migration jobs as part of a single migration process.

Instead of running each migration configuration individually, Batch Migration allows a group of migration jobs to be processed together.

---

## Purpose

During a migration cycle, several related M3 objects may need to be generated.

Running them individually would require:

```text
Migration 1
    ↓
Wait
    ↓
Migration 2
    ↓
Wait
    ↓
Migration 3
```

Batch Migration provides a more efficient workflow:

```text
Batch
 │
 ├── Migration 1
 ├── Migration 2
 ├── Migration 3
 └── Migration 4
        │
        ▼
 Execute Jobs
        │
        ▼
Generated Migration Files
```

---

# When to Use Batch Migration

Batch Migration is useful when:

- Multiple migration objects need to be processed.
- Several configured migrations belong to the same migration cycle.
- The migration configurations have already been tested individually.
- Multiple M3 load files need to be regenerated.
- A repeatable group of migration jobs is required.

---

# Before Running a Batch

Each migration included in the batch should already have a valid configuration.

Verify:

```text
Source Map
    ✓

Migration Map
    ✓

Rule Configuration
    ✓

SDT Template
    ✓

Transaction Sheets
    ✓
```

Batch Migration should normally be used **after the individual migrations have been tested successfully**.

!!! important
    Batch processing does not replace migration validation. Each migration configuration should be tested before it becomes part of a larger batch.

---

# Batch Migration Flow

The general process is:

```text
Batch Configuration
        │
        ▼
Load Migration Jobs
        │
        ▼
Process Job 1
        │
        ▼
Process Job 2
        │
        ▼
Process Job 3
        │
        ▼
Generate Output Files
```

Each job uses the same core migration engine as an individual migration.

---

# Individual Job Processing

For each migration job, the Data Migration Tool performs the normal migration workflow:

```text
Migration Job
     │
     ▼
Resolve Source
     │
     ▼
Load Source Data
     │
     ▼
Load Rule Configuration
     │
     ▼
Apply FILTER Rules
     │
     ▼
Apply Transformation Rules
     │
     ▼
Resolve SDT Configuration
     │
     ▼
Generate Output
```

The batch functionality coordinates multiple executions of this process.

---

# Example Batch

A migration cycle could contain:

```text
Item Master
Item Facility
Item Warehouse
```

The batch would conceptually execute:

```text
Batch Migration
│
├── Item Master
│      └── MMS200MI
│
├── Item Facility
│      └── Configured API
│
└── Item Warehouse
       └── Configured API
```

Each migration can have its own:

- Source
- Rule Configuration
- API
- SDT Template
- Transaction sheets

---

# Migration Dependencies

The order of migration objects may be important.

For example:

```text
Item Master
     │
     ▼
Item Facility
     │
     ▼
Item Warehouse
```

If one migration object logically depends on another, the batch should respect the approved migration sequence.

!!! warning
    Do not assume that migration jobs can always be executed in any order. Follow the migration sequence defined for the M3 implementation.

---

# Rule Configurations

Each migration job continues to use its own Rule Configuration.

For example:

```text
Batch
 │
 ├── Migration A
 │      └── Rules A
 │
 ├── Migration B
 │      └── Rules B
 │
 └── Migration C
        └── Rules C
```

Changes to one Rule Configuration therefore affect the corresponding migration when the batch is executed.

---

# Scope

Where applicable, the migration scope must also be considered.

For example:

```text
GLOBAL
DIV_US
DIV_CA
```

Scope-specific rules may produce different results from the same source data.

Before running a production batch, verify that the appropriate scope is being used for the migration jobs.

---

# Output Files

Each migration job generates its corresponding M3 load output.

Conceptually:

```text
Batch Run
 │
 ├── Item Master
 │      └── LOAD_....xlsx
 │
 ├── Item Facility
 │      └── LOAD_....xlsx
 │
 └── Item Warehouse
        └── LOAD_....xlsx
```

Generated files are stored in the configured output location, normally:

```text
output/
```

---

# Monitoring the Batch

During execution, monitor the **System Log**.

The log provides information about:

- Migration jobs being processed
- Source loading
- Rule processing
- Generated files
- Warnings
- Errors

A batch containing several jobs may take longer than an individual migration.

---

# Error Handling

If a migration job encounters an error, review the System Log to determine:

```text
Which migration failed?
        │
        ▼
At what stage?
        │
        ├── Source
        ├── Configuration
        ├── Rules
        ├── SDT
        └── Output
```

Do not assume that all generated files are valid simply because the batch process started successfully.

---

# Recommended Batch Workflow

Before a production batch:

1. Verify the Source Map.
2. Verify the Migration Map.
3. Review recent Rule Configuration changes.
4. Test modified migrations individually.
5. Confirm the migration sequence.
6. Confirm the appropriate scope.
7. Run the batch.
8. Review the System Log.
9. Verify all expected output files were generated.
10. Validate record counts.
11. Validate transformed values.
12. Confirm the files are ready for the next migration stage.

---

# Validation

Batch validation should occur at two levels.

## Batch-Level Validation

Confirm that every expected migration was processed.

For example:

```text
Expected Jobs: 5
Generated Files: 5
```

Investigate any difference.

---

## Migration-Level Validation

Each generated file must still be validated individually.

For example:

```text
Item Master
│
├── Record count
├── Required fields
├── Rule results
└── SDT structure

Item Facility
│
├── Record count
├── Required fields
├── Rule results
└── SDT structure
```

A successful batch execution does not automatically mean that the business transformation results are correct.

---

# Batch vs Standard Migration

Standard Migration processes one migration configuration:

```text
Configuration
      │
      ▼
Migration
      │
      ▼
Output
```

Batch Migration coordinates multiple migration jobs:

```text
Batch
 │
 ├── Configuration A → Output A
 ├── Configuration B → Output B
 └── Configuration C → Output C
```

The underlying transformation concepts remain the same.

---

# When Not to Use Batch Migration

Avoid using Batch Migration as the first test of a new migration configuration.

Instead:

```text
Develop Rules
     │
     ▼
Load by ID
     │
     ▼
Standard Migration
     │
     ▼
Validate Output
     │
     ▼
Batch Migration
```

This makes troubleshooting significantly easier.

---

# Troubleshooting

If a batch migration fails, determine which job caused the problem.

Then test that migration individually using **Standard Migration**.

Verify:

- Source availability
- Source Map
- Migration Map
- Rule Configuration
- Rule Scope
- SDT template
- Transaction sheets
- Output permissions

For rule-specific problems, use **Load by ID** to isolate representative records.

---

# Recommended Migration Progression

A useful progression for developing a migration is:

```text
Configure Migration
       │
       ▼
Develop Rules
       │
       ▼
Load by ID
       │
       ▼
Standard Migration
       │
       ▼
Validate
       │
       ▼
Batch Migration
```

This keeps Batch Migration focused on executing configurations that have already been tested and approved.