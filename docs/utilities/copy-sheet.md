# Copy Sheet Utility

The **Copy Sheet** utility is used to copy worksheet data between Excel workbooks as part of migration preparation or file-management activities.

This utility operates independently from the main migration process.

---

## Purpose

Migration work often requires moving data from one workbook into another while preparing source files, templates, or supporting migration files.

The Copy Sheet utility simplifies this process.

```text
Source Workbook
      │
      ▼
Select Worksheet
      │
      ▼
Copy Sheet
      │
      ▼
Destination Workbook
```

---

# Accessing the Utility

Navigate to:

**Utilities → Copy Sheet**

The utility provides the controls required to select the source and destination Excel files.

---

# Typical Use Cases

Copy Sheet can be useful when:

- Preparing migration workbooks
- Moving data between Excel files
- Consolidating migration information
- Reusing an existing worksheet
- Preparing supporting files for testing
- Avoiding manual Excel copy-and-paste operations

---

# Basic Workflow

The general process is:

```text
Select Source File
       │
       ▼
Select Source Sheet
       │
       ▼
Select Destination
       │
       ▼
Copy
       │
       ▼
Review Result
```

---

# Source Workbook

Select the Excel workbook containing the worksheet that should be copied.

For example:

```text
Legacy_Item_Data.xlsx
```

The workbook may contain several worksheets:

```text
Legacy_Item_Data.xlsx
│
├── Item Master
├── Item Facility
└── Item Warehouse
```

Select the worksheet required for the operation.

---

# Destination Workbook

Select the workbook that should receive the copied worksheet.

For example:

```text
Migration_Workbook.xlsx
```

The selected sheet is then copied into the destination workbook according to the utility's configured behavior.

---

# Validation

After copying a worksheet, verify:

- The expected worksheet exists in the destination workbook.
- Column headers were preserved.
- The expected number of records was copied.
- Important values were preserved.
- The destination workbook remains valid.

For migration-related data, comparing row counts is recommended:

```text
Source Sheet
12,450 rows

Destination Sheet
12,450 rows
```

Unexpected differences should be investigated.

---

# Important Consideration

The Copy Sheet utility performs a file-preparation operation.

It does **not** apply migration transformation rules.

```text
Copy Sheet
    │
    └── Moves worksheet data

Migration Engine
    │
    └── Applies transformation rules
```

These are separate processes.

---

# Recommended Practice

When using Copy Sheet:

1. Confirm the correct source workbook.
2. Confirm the correct source worksheet.
3. Confirm the destination workbook.
4. Perform the copy.
5. Review the resulting workbook.
6. Validate row counts when applicable.
7. Preserve the original source file when possible.

---

# Troubleshooting

If the operation fails, verify:

- The source workbook exists.
- The destination workbook exists or is valid for the requested operation.
- The selected worksheet exists.
- The workbook is not locked by another application.
- The user has permission to access the files.
- The files are valid Excel workbooks.

If Excel has the destination workbook open, close it and retry the operation.

---

# Relationship to Migration

Copy Sheet can support migration preparation, but it is not part of the transformation pipeline.

```text
Preparation Utilities
        │
        ▼
Prepared Source Files
        │
        ▼
Data Migration Tool
        │
        ▼
Transformation Rules
        │
        ▼
M3 SDT Output
```

Use this utility when workbook preparation is required before or after the main migration process.