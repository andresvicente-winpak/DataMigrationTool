# Standard Migration

Standard Migration is used to execute a migration using an existing **Rule Configuration** and a source that has already been defined in the application's configuration.

This is the primary migration method when the source, M3 API, transformation rules, and SDT template are already configured.

---

## Standard Migration Screen

Navigate to:

**Run Migration → Standard**

The Standard Migration screen contains three main selections:

1. Rule Configuration
2. Source Data
3. Scope

After these selections have been made, the migration can be started using **RUN MIGRATION**.

---

## 1. Select Rule Configuration

Select the Rule Configuration that should be used for the migration.

Example:

```text
MMS200MI
```

Rule Configuration files are stored in:

```text
config/rules/
```

For example:

```text
config/rules/MMS200MI.xlsx
```

The selected configuration determines the transformation rules that will be applied to the source data.

---

## 2. Select Source Data

Select the source dataset from the **Source Data** dropdown.

The available sources are obtained from:

```text
config/source_map.csv
```

The selected entry identifies the source associated with an MCO sheet.

The source may point to a file or another configured source such as SQL.

!!! note
    The source must already exist in the Source Map before it can be selected from Standard Migration.

---

## 3. Select Scope

The default scope is:

```text
GLOBAL
```

A different scope can be selected when business-unit-specific transformation rules are required.

For example:

```text
GLOBAL
DIV_US
DIV_CA
```

When a specific scope is selected, the Rule Configuration loader determines which rules apply to that scope.

Specific scope rules can override GLOBAL rules for the same target field.

---

## 4. Run Migration

After selecting the Rule Configuration, source data, and scope, click:

**RUN MIGRATION**

The application starts the migration process.

A simplified execution flow is:

```text
Selected Source
      │
      ▼
Load Source Data
      │
      ▼
Load Rule Configuration
      │
      ▼
Apply Filters
      │
      ▼
Apply Transformation Rules
      │
      ▼
Resolve SDT Template
      │
      ▼
Generate Output File
```

Progress and errors are displayed in the **System Log** at the bottom of the application.

---

## SDT Template Resolution

The application uses the migration configuration to determine which SDT template and transaction sheets are required.

SDT templates are normally located under:

```text
config/sdt_templates/
```

The relationship between the source, API, SDT template, and transaction sheets is controlled by the migration configuration.

---

## Output File

Generated migration files are written to:

```text
output/
```

A typical automatically generated filename follows a structure similar to:

```text
LOAD_MMS200MI_20260819.xlsx
```

If a file with the same name already exists, the application can generate a versioned filename such as:

```text
LOAD_MMS200MI_20260819_v1.xlsx
```

---

## Example

Assume the following configuration:

```text
Rule Configuration: MMS200MI
Source Data:       Item Master
Scope:             GLOBAL
```

The application will:

1. Locate the configured source for **Item Master**.
2. Resolve the migration mapping associated with that source.
3. Load the `MMS200MI` Rule Configuration.
4. Load the legacy source data.
5. Apply configured filters and transformation rules.
6. Determine the appropriate SDT template and transaction sheets.
7. Generate the resulting M3 load file in the `output` directory.

---

## Troubleshooting

If the migration does not run, verify:

- A valid Rule Configuration is selected.
- The source exists in `config/source_map.csv`.
- The corresponding migration mapping exists.
- The Rule Configuration file exists.
- The required SDT template exists.
- The source file or SQL source is accessible.

Check the **System Log** for detailed messages generated during processing.