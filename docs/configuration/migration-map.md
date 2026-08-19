# Migration Map

The **Migration Map** defines the relationship between a migration business object and the M3 configuration required to generate its output.

While the **Source Map** determines where the legacy data comes from, the Migration Map determines where that data is going and which M3 structure should be used.

---

## Purpose

The Migration Map connects the major M3 components of a migration:

```text
MCO Business Object
        │
        ▼
   Migration Map
        │
        ├── M3 API
        │
        ├── SDT Template
        │
        └── Transaction Sheet(s)
```

This allows the Data Migration Tool to automatically resolve the correct output configuration during migration.

---

# Migration Map File

The Migration Map is stored in:

```text
config/migration_map.csv
```

The Migration Runner reads this file when determining the API, SDT template, and transaction sheets associated with a migration.

---

# Main Fields

The current Migration Map uses the following key fields:

| Field | Purpose |
| --- | --- |
| `MCO_SHEET` | Identifies the migration business object |
| `API_NAME` | Identifies the M3 API / Rule Configuration |
| `SDT_TEMPLATE` | Identifies the SDT workbook used for output |
| `TRANSACTION_SHEET` | Identifies the SDT sheet or sheets to populate |

---

# MCO_SHEET

`MCO_SHEET` identifies the business object from the MCO.

Examples may include:

```text
Item Master
Item Facility
Item Warehouse
Supplier
Customer
```

This is an important lookup key because multiple migration objects may use related APIs but require different SDT configurations.

For example:

```text
Item Master
      │
      ▼
Migration Map
```

The tool can then resolve the exact configuration associated with **Item Master**.

---

# API_NAME

`API_NAME` identifies the M3 API associated with the migration.

Example:

```text
MMS200MI
```

The API is also associated with the Rule Configuration used by the migration.

For example:

```text
MCO Sheet
   │
   ▼
Item Master
   │
   ▼
MMS200MI
   │
   ▼
config/rules/MMS200MI.xlsx
```

---

# SDT_TEMPLATE

`SDT_TEMPLATE` identifies the M3 SDT workbook that should be used to generate the migration output.

SDT templates are normally stored under:

```text
config/sdt_templates/
```

For example:

```text
config/sdt_templates/MMS200MI.xlsx
```

The template provides the required M3 workbook structure.

The Data Migration Tool does not create the SDT structure from scratch during each migration. Instead, it populates the appropriate fields in the configured template.

---

# TRANSACTION_SHEET

`TRANSACTION_SHEET` identifies the SDT worksheet or worksheets that should receive the transformed data.

A migration may require one transaction sheet:

```text
AddItemBasic
```

or multiple transaction sheets:

```text
AddItemBasic,UpdItemBasic
```

Multiple transaction sheets are stored as a comma-separated list.

The Migration Runner converts this value into the list of sheets that should be processed.

---

# Complete Relationship

A Migration Map entry can be understood as:

```text
MCO_SHEET
    │
    ▼
Business Object
    │
    ├── API_NAME
    │
    ├── SDT_TEMPLATE
    │
    └── TRANSACTION_SHEET
```

For example:

```text
Item Master
    │
    ├── MMS200MI
    ├── MMS200MI SDT Template
    └── Required Transaction Sheet(s)
```

---

# Relationship with Source Map

The Source Map and Migration Map work together.

```text
                   MCO_SHEET
                       │
            ┌──────────┴──────────┐
            │                     │
            ▼                     ▼
       Source Map            Migration Map
            │                     │
            ▼                     ├── API
      Legacy Source               ├── SDT Template
                                  └── Transaction Sheets
            │                     │
            └──────────┬──────────┘
                       ▼
                 Migration Runner
```

The **Source Map** answers:

> Where does the data come from?

The **Migration Map** answers:

> Which M3 configuration should process it?

---

# Relationship with Rules

The API resolved through the Migration Map is connected to the Rule Configuration.

For example:

```text
MCO_SHEET
    │
    ▼
Item Master
    │
    ▼
API_NAME
    │
    ▼
MMS200MI
    │
    ▼
Rule Configuration
    │
    ▼
config/rules/MMS200MI.xlsx
```

The Rule Configuration then determines how each source field is transformed.

---

# Standard Migration Flow

When Standard Migration is executed, the process can be represented as:

```text
User selects MCO Source
        │
        ▼
Resolve Source Map
        │
        ▼
Load Legacy Source
        │
        ▼
Resolve Migration Map
        │
        ├── API
        ├── SDT Template
        └── Transaction Sheets
        │
        ▼
Load Rule Configuration
        │
        ▼
Apply Transformation Rules
        │
        ▼
Generate SDT Output
```

---

# Why MCO_SHEET Is Important

The Data Migration Tool can resolve migration information using `MCO_SHEET`.

This is important because using only the API name may be ambiguous.

For example, two MCO business objects could potentially reference the same API but require different migration definitions.

Conceptually:

```text
Business Object A ──┐
                    ├── Same API
Business Object B ──┘
```

But they may require:

```text
Different SDT Templates
or
Different Transaction Sheets
```

Using the MCO context allows the application to select the configuration belonging to the actual business object.

---

# SDT Template Resolution

During migration, the application first attempts to resolve the configured SDT template from the Migration Map.

If the configured template exists under:

```text
config/sdt_templates/
```

that template is used.

The Migration Runner also contains fallback logic that can search for a template beginning with the API name.

However, the preferred configuration is an explicit Migration Map entry.

!!! tip
    Maintain the Migration Map correctly rather than relying on fallback template detection.

---

# Example Migration Map

A simplified configuration could look like:

```csv
MCO_SHEET,API_NAME,SDT_TEMPLATE,TRANSACTION_SHEET
Item Master,MMS200MI,MMS200MI.xlsx,AddItemBasic
```

This means:

```text
Item Master
     │
     ├── API: MMS200MI
     │
     ├── Template: MMS200MI.xlsx
     │
     └── Sheet: AddItemBasic
```

When **Item Master** is selected, the application has the information required to determine the M3 output configuration.

---

# Multiple Transaction Sheets

Some migrations require more than one SDT transaction sheet.

Example:

```csv
TRANSACTION_SHEET
AddItemBasic,UpdItemBasic
```

The application separates the value by commas:

```text
AddItemBasic
UpdItemBasic
```

and processes the required sheets.

This allows one migration configuration to populate multiple sections of an SDT workbook when required.

---

# Creating a New Migration Map Entry

When configuring a new migration object:

1. Identify the MCO business object.
2. Identify the M3 API.
3. Identify the correct SDT template.
4. Identify the required transaction sheet or sheets.
5. Add the configuration to `migration_map.csv`.
6. Verify that the SDT template exists.
7. Verify that the corresponding Rule Configuration exists.
8. Test the migration.

---

# Validation

Before running a production migration, verify the complete relationship:

```text
MCO Sheet
    ↓
Source Map
    ↓
Correct Legacy Source

MCO Sheet
    ↓
Migration Map
    ↓
Correct API
    ↓
Correct Rule Configuration
    ↓
Correct SDT Template
    ↓
Correct Transaction Sheet(s)
```

A mistake in the Migration Map can cause the migration to use the wrong M3 structure even when the transformation rules themselves are correct.

---

# Troubleshooting

If the migration cannot determine the SDT template, verify:

- `config/migration_map.csv` exists.
- `MCO_SHEET` matches the selected business object.
- `API_NAME` is correct.
- `SDT_TEMPLATE` contains the correct filename.
- The template exists under `config/sdt_templates/`.
- `TRANSACTION_SHEET` contains valid SDT worksheet names.

If the wrong migration configuration is selected, verify the `MCO_SHEET` relationship first.

---

# Recommended Practice

Treat `MCO_SHEET` as the business-object identity connecting the migration configuration.

A well-defined configuration should make this relationship clear:

```text
MCO Business Object
        │
        ├── Source
        ├── API
        ├── Rules
        ├── SDT Template
        └── Transaction Sheets
```

This provides a consistent and repeatable migration setup.

---

# Next Step

After configuring the Migration Map, review **Business Units and Scope** to understand how GLOBAL rules and business-specific overrides are applied.