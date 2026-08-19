# Rule Configuration Snapshots

A **Snapshot** preserves a copy of a Rule Configuration at a specific point in time.

Snapshots provide a safety mechanism when rules are being developed, modified, or tested.

---

## Purpose

Rule Configurations can contain many transformation rules.

A change to one rule may affect the generated migration output.

Before significant modifications, preserving the current configuration provides a known reference point.

```text
Current Rules
     │
     ▼
Create Snapshot
     │
     ▼
Modify Rules
     │
     ▼
Test Changes
```

If necessary, the previous configuration can then be reviewed.

---

# When to Create a Snapshot

Snapshots are especially useful before:

- Major rule changes
- Business-unit overrides
- Large MAP updates
- Complex Python-rule modifications
- Rule Configuration restructuring
- Testing significant migration changes

For example:

```text
Working PLCD Rule
       │
       ▼
Create Snapshot
       │
       ▼
Modify PLCD Logic
       │
       ▼
Test
```

---

# Snapshots vs Rule History

These concepts are related but serve different purposes.

| Rule History | Snapshot |
| --- | --- |
| Tracks changes | Preserves a configuration state |
| Useful for investigation | Useful as a recovery/reference point |
| Shows evolution | Represents a point in time |

Together they provide better control over Rule Configuration changes.

---

# Example

Suppose `MMS200MI.xlsx` contains validated rules.

Before implementing several new requirements:

```text
MMS200MI.xlsx
      │
      ▼
Snapshot
      │
      ▼
Modify Rules
      │
      ├── PLCD
      ├── PUIT
      └── ORTY
```

If unexpected results appear, the previous snapshot provides a reference for comparing the configuration before and after the modifications.

---

# Snapshot Validation

A snapshot should not automatically be considered a valid production configuration simply because it exists.

It represents:

> The Rule Configuration at a particular point in time.

The migration status at that point should still be understood.

For example:

```text
Snapshot A
    └── Development

Snapshot B
    └── Tested

Snapshot C
    └── Production Candidate
```

The team should know which configuration was actually validated.

---

# Recommended Practice

Create meaningful snapshots around significant migration milestones.

For example:

```text
Initial Configuration
        │
        ▼
Rules Developed
        │
        ▼
Snapshot
        │
        ▼
Business Testing
        │
        ▼
Rule Corrections
        │
        ▼
Snapshot
        │
        ▼
Final Validation
```

Avoid relying on memory to determine which version contained a particular transformation.

---

# Troubleshooting with Snapshots

If a previously working migration begins producing different results:

1. Identify the affected target fields.
2. Review recent Rule History.
3. Compare the current configuration with a previous snapshot.
4. Identify changed rules.
5. Review the associated business requirement.
6. Test representative records.
7. Correct the current configuration if necessary.

---

# Important Consideration

Snapshots should support controlled rule development, not replace testing.

The preferred workflow remains:

```text
Snapshot
   │
   ▼
Modify
   │
   ▼
Test
   │
   ▼
Validate
   │
   ▼
Approve
```

---

# Next Step

Continue to **Sync / Merge** to understand how Rule Configuration changes can be compared and merged between different copies of the configuration.