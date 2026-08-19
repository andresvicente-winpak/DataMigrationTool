# Merge Files Utility

The **Merge Files** utility is used to combine migration-related files into a single consolidated file.

This utility supports file preparation and consolidation activities and operates separately from the main migration transformation process.

---

## Purpose

During migration activities, data may be distributed across multiple files that need to be combined.

Conceptually:

```text
File 1 ──┐
File 2 ──┼──► Merge Files ──► Combined File
File 3 ──┘
```

The Merge Files utility provides a controlled way to perform this consolidation from within the Data Migration Tool.

---

# Accessing the Utility

Navigate to:

**Utilities → Merge Files**

Use the interface to select the files that should participate in the merge operation.

---

# Typical Use Cases

Merge Files can be useful when:

- Migration data has been delivered in multiple files
- Multiple extracts belong to the same dataset
- Migration results need to be consolidated
- Test files need to be combined
- Supporting migration files need to be prepared

---

# Basic Workflow

The general workflow is:

```text
Select Files
     │
     ▼
Review Selection
     │
     ▼
Merge
     │
     ▼
Combined File
     │
     ▼
Validate Result
```

---

# Selecting Files

Select the files that should be included in the merge.

Before proceeding, verify that the selected files belong to the same intended dataset or migration activity.

For example:

```text
ItemMaster_Part1.xlsx
ItemMaster_Part2.xlsx
ItemMaster_Part3.xlsx
```

can represent different portions of the same source dataset.

---

# Data Compatibility

Files being combined should have compatible structures.

For example:

```text
File 1

ITNO | ITTY | PUIT
```

```text
File 2

ITNO | ITTY | PUIT
```

These structures can logically be consolidated because the columns represent the same data.

Be careful when files contain different structures:

```text
File 1

ITNO | ITTY | PUIT
```

```text
File 2

CUNO | CUNM | STAT
```

These represent different datasets and should not normally be treated as one migration source.

---

# Validation After Merge

Always review the resulting file after a merge operation.

Verify:

- Expected files were included
- Column structure is correct
- Records were not unintentionally omitted
- Unexpected duplicate records were not introduced
- The resulting file can be opened successfully

---

# Record Count Validation

Record counts provide a useful first validation.

For example:

```text
File 1     5,000 rows
File 2     4,000 rows
File 3     3,000 rows
           ──────────
Expected  12,000 rows
```

The merged result should contain the expected population, accounting for any headers or other structural behavior of the merge process.

!!! important
    Record count validation does not guarantee that the data is correct, but unexpected differences should always be investigated.

---

# Duplicate Records

Combining multiple files can potentially introduce duplicate business records.

For example:

```text
File 1
AP32508

File 2
AP32508
```

After consolidation:

```text
AP32508
AP32508
```

Whether this is valid depends on the source and migration requirement.

The merge operation should therefore be followed by business-key validation when duplicate records are a concern.

---

# Merge Files vs Migration

The Merge Files utility does not replace the migration transformation process.

```text
Merge Files
     │
     ▼
Combined Dataset
     │
     ▼
Migration Engine
     │
     ▼
Rule Configuration
     │
     ▼
M3 Output
```

The merge prepares or consolidates data.

The Migration Engine applies transformation rules.

---

# Recommended Practice

Before merging:

1. Confirm all files belong to the intended dataset.
2. Review their column structures.
3. Record the source row counts.
4. Perform the merge.
5. Review the resulting structure.
6. Validate the resulting row count.
7. Check important business keys for unexpected duplicates.
8. Preserve the original files.

---

# Troubleshooting

If a merge operation fails, verify:

- All selected files exist
- The files are accessible
- The files use supported formats
- Files are not locked by another application
- The user has permission to read and write the required locations

If the merge completes but the result is unexpected, compare:

```text
Individual Source Files
          │
          ▼
Expected Structure / Count
          │
          ▼
Merged Result
```

This helps identify which input file introduced the difference.

---

# Important Consideration

The exact merge behavior depends on the implementation of the utility.

Before relying on the result for production migration, verify how the utility handles:

- Column differences
- Worksheet selection
- Headers
- Duplicate records
- Blank rows

The resulting file should always be validated before it becomes a migration source.

---

# Next Step

Continue to **Script Library** to learn about the reusable migration utilities and scripts available within the Data Migration Tool.