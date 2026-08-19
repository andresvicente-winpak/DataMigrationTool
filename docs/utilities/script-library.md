# Script Library

The **Script Library** provides access to reusable scripts and utilities that support data migration activities.

These scripts are intended for tasks that do not necessarily belong to the standard migration workflow but are useful during data preparation, analysis, validation, or troubleshooting.

---

## Purpose

Not every migration task requires a complete migration process.

Sometimes a specific operation is needed:

```text
Migration Task
      │
      ▼
Reusable Script
      │
      ▼
Processed Result
```

The Script Library provides a central location for these supporting tools.

---

# Accessing the Script Library

Navigate to:

**Utilities → Script Library**

The available scripts can then be reviewed and executed according to their intended purpose.

---

# Why Use a Script Library?

During migration development, small utilities are often created to solve specific problems.

Without a central library, these scripts can become scattered across:

```text
Personal folders
Downloads
Temporary directories
Test folders
Desktop
```

A Script Library provides a more controlled approach:

```text
Data Migration Tool
        │
        ▼
   Script Library
        │
        ├── Script A
        ├── Script B
        ├── Script C
        └── Script D
```

This makes useful migration utilities easier to find and reuse.

---

# Typical Script Uses

Scripts may support activities such as:

- Data preparation
- File manipulation
- Data validation
- Migration analysis
- Data comparison
- Troubleshooting
- Repetitive migration tasks

The exact functionality depends on the scripts currently available in the library.

---

# Script Library vs Migration Rules

Scripts and migration rules serve different purposes.

## Migration Rules

Migration rules determine how source records are transformed into M3 values.

For example:

```text
MBPUIT
   │
   ▼
PYTHON Rule
   │
   ▼
PLCD
```

Rules belong to the migration transformation process.

---

## Scripts

Scripts perform supporting operations outside the normal field-level transformation process.

```text
Source / Migration File
          │
          ▼
        Script
          │
          ▼
Processed / Analyzed Result
```

A reusable script should not replace a Rule Configuration when the logic belongs to the normal migration transformation.

---

# Before Running a Script

Before executing a script, understand:

1. What the script does.
2. What input it requires.
3. What output it produces.
4. Whether it modifies an existing file.
5. Whether the operation can be reversed.
6. Whether the result needs additional validation.

!!! warning
    Do not run an unfamiliar script against production migration files without first understanding its purpose and expected result.

---

# Input Files

Some scripts may require an input file.

For example:

```text
Input File
    │
    ▼
Script
    │
    ▼
Output File
```

Always verify that the selected input corresponds to the purpose of the script.

Preserving the original file before running a modification script is recommended.

---

# Output Validation

The output of a script should be validated before being used in a migration.

For example:

```text
Original Data
      │
      ▼
Run Script
      │
      ▼
Generated Result
      │
      ▼
Validate
      │
      ▼
Use in Migration
```

Validation may include:

- Row counts
- Column names
- Business keys
- Expected values
- File structure
- Duplicate records

---

# Script Documentation

Each reusable script should ideally document:

| Information | Description |
| --- | --- |
| Name | Script name |
| Purpose | What problem it solves |
| Input | Required file or data |
| Output | What the script generates |
| Changes Data | Whether source data is modified |
| Usage | How to run it |
| Validation | What should be checked afterward |

For example:

```text
Script: Example Utility

Purpose:
Prepare source data for migration.

Input:
Legacy Excel file.

Output:
Processed Excel file.

Validation:
Compare source and output record counts.
```

---

# Adding New Scripts

When a reusable migration utility is created, consider adding it to the Script Library instead of keeping it as an isolated personal script.

Before adding a script:

1. Confirm that the operation is reusable.
2. Give the script a descriptive name.
3. Define its expected input.
4. Define its expected output.
5. Add appropriate error handling.
6. Test the script.
7. Document its purpose.
8. Define how its result should be validated.

---

# Recommended Practice

Use the Script Library for **supporting migration utilities**.

Use the Rule Configuration for **data transformation logic that belongs to the migration itself**.

A useful distinction is:

```text
Does this determine an M3 target value?
              │
       ┌──────┴──────┐
       │             │
      Yes            No
       │             │
       ▼             ▼
Migration Rule    Utility Script
```

This helps keep migration logic centralized and maintainable.

---

# Troubleshooting

If a script fails:

- Verify the correct script was selected.
- Verify the required input was provided.
- Check that input files exist.
- Check that files are not locked.
- Verify read/write permissions.
- Review any error message produced by the script.

If the script completes but the result is unexpected, do not immediately use the output for migration.

Compare:

```text
Expected Result
       ↕
Actual Result
```

and investigate the difference first.

---

# Utilities Summary

The Utilities section provides supporting functionality around the core migration process:

```text
Utilities
│
├── Copy Sheet
│
├── Merge Files
│
└── Script Library
```

These tools assist with preparing and managing migration data but remain separate from the core Rule Configuration and transformation process.