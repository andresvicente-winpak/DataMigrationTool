# Sync / Merge Rule Configurations

The **Sync / Merge** functionality is used to compare two Rule Configuration files and selectively merge changes into the local configuration.

This is useful when multiple users are working on migration rules independently.

---

## Purpose

During migration development, two users may modify different copies of the same Rule Configuration.

For example:

```text
Original MMS200MI
       │
       ├───────────────┐
       ▼               ▼
   User A           User B
       │               │
Modify PLCD        Modify PUIT
Modify ORTY        Modify EOQM
       │               │
       └───────┬───────┘
               ▼
          Sync / Merge
               │
               ▼
     Combined Configuration
```

Sync / Merge provides a controlled method for reviewing those differences before modifying the local Rule Configuration.

---

# Accessing Sync / Merge

Navigate to:

**Sync / Merge**

from the main application.

The Sync / Merge interface allows a local Rule Configuration to be compared with another Rule Configuration.

---

# Local and Incoming Configuration

The process involves two configurations.

### Local Configuration

The Rule Configuration currently maintained by the user.

For example:

```text
config/rules/MMS200MI.xlsx
```

### Incoming Configuration

Another copy containing changes that may need to be incorporated.

For example:

```text
User B
   │
   ▼
MMS200MI.xlsx
```

The incoming file should be reviewed before its changes are applied locally.

---

# Comparison Process

The application compares the Rule Configurations and identifies differences.

Conceptually:

```text
Local Rules
     │
     │
     ├──── Compare ──── Incoming Rules
     │                       │
     └──────────┬────────────┘
                ▼
           Differences
                │
                ▼
             Review
```

This is safer than simply replacing the local Rule Configuration with another user's file.

---

# What Should Be Compared

Important rule properties include:

```text
TARGET_FIELD
SOURCE_FIELD
RULE_TYPE
RULE_VALUE
SCOPE
```

For example:

### Local

```text
TARGET_FIELD = PLCD
SOURCE_FIELD = MBPUIT
RULE_TYPE    = PYTHON
SCOPE        = GLOBAL
```

### Incoming

```text
TARGET_FIELD = PLCD
SOURCE_FIELD = MBPUIT
RULE_TYPE    = PYTHON
SCOPE        = DIV_US
```

Although the target field is the same, the scope is different and should be reviewed carefully.

---

# Example Rule Difference

Suppose the local PLCD rule contains:

```python
if puit == "1":
    return "01"
```

The incoming configuration contains:

```python
if itty in ("SF", "FG") and puit == "1":
    return "02"
```

Sync / Merge should allow this difference to be identified before the local configuration is changed.

The reviewer should determine:

```text
Which rule represents the approved requirement?
```

before accepting the incoming change.

---

# Selective Merge

Not every incoming change needs to be accepted.

For example:

```text
Incoming Changes

PLCD   ✓ Accept
PUIT   ✓ Accept
ORTY   ✗ Reject
EOQM   ✓ Accept
```

The purpose of Sync / Merge is to allow changes to be reviewed individually rather than blindly replacing the entire configuration.

---

# Conflicting Changes

A conflict can occur when both configurations modify the same rule differently.

Example:

```text
               PLCD
                 │
        ┌────────┴────────┐
        ▼                 ▼
      Local            Incoming
        │                 │
   PLCD = 02          PLCD = 03
        │                 │
        └────────┬────────┘
                 ▼
              Conflict
```

A conflict requires review.

The correct result should be determined from the approved migration business requirement, not simply from which file is newer.

---

# Scope Conflicts

Scope should always be considered during comparison.

These may represent two separate valid rules:

```text
PLCD / GLOBAL
PLCD / DIV_US
```

while these may represent competing versions:

```text
PLCD / GLOBAL
PLCD / GLOBAL
```

The combination of:

```text
TARGET_FIELD + SCOPE
```

is therefore important when reviewing differences.

---

# Recommended Merge Workflow

Use the following process when receiving another Rule Configuration:

1. Identify the incoming Rule Configuration.
2. Open Sync / Merge.
3. Compare it with the local configuration.
4. Review the detected differences.
5. Identify changed target fields.
6. Review changes to `SOURCE_FIELD`.
7. Review changes to `RULE_TYPE`.
8. Review changes to `RULE_VALUE`.
9. Review changes to `SCOPE`.
10. Select only the approved changes.
11. Merge them into the local configuration.
12. Review the resulting Rule Configuration.
13. Test the affected migrations.

---

# After a Merge

A successful merge means:

> The selected configuration changes were incorporated.

It does **not** mean:

> The resulting migration output has been validated.

After merging:

```text
Merge
 │
 ▼
Identify Changed Rules
 │
 ▼
Determine Dependencies
 │
 ▼
Load by ID
 │
 ▼
Validate Results
 │
 ▼
Standard Migration
 │
 ▼
Validate Complete Output
```

---

# Example

Suppose another user sends an updated `MMS200MI` Rule Configuration.

The comparison identifies:

| Target Field | Local | Incoming | Action |
| --- | --- | --- | --- |
| ITNO | DIRECT | DIRECT | No change |
| PLCD | Python v1 | Python v2 | Review |
| PUIT | DIRECT | Python | Review |
| ORTY | Python v1 | Python v1 | No change |

Only `PLCD` and `PUIT` require attention.

After reviewing the requirements, you may decide:

```text
PLCD → Accept Incoming
PUIT → Keep Local
```

The resulting configuration contains the approved combination of both files.

---

# Multiple Users

When several users work on Rule Configurations, avoid workflows such as:

```text
User A edits file
      ↓
User B replaces file
      ↓
User A changes disappear
```

Instead use:

```text
User A Changes
      │
      ├──────┐
      │      │
User B Changes
      │      │
      └──┬───┘
         ▼
     Compare
         │
         ▼
      Review
         │
         ▼
       Merge
```

This reduces the risk of accidentally overwriting valid migration logic.

---

# Snapshots Before Merge

Before performing a significant merge, preserving the current Rule Configuration is recommended.

```text
Current Configuration
        │
        ▼
     Snapshot
        │
        ▼
   Sync / Merge
        │
        ▼
      Testing
```

This provides a reference point if unexpected results appear after the merge.

---

# Validation After Merge

Focus testing on the fields that changed.

For example:

```text
Changed:

PLCD
PUIT
ORTY
```

Identify representative records for those rules and test them using **Load by ID**.

Then perform a Standard Migration to verify the complete output.

---

# Troubleshooting

If unexpected results appear after a merge:

1. Identify the affected target field.
2. Review the current rule.
3. Review Rule History.
4. Compare with the pre-merge configuration or snapshot.
5. Check whether the rule came from the incoming configuration.
6. Verify the selected scope.
7. Test representative source records.
8. Compare expected and actual results.

---

# Recommended Practice

Treat Sync / Merge as a **review process**, not simply a file-copy operation.

The preferred workflow is:

```text
Compare
   │
   ▼
Understand
   │
   ▼
Select
   │
   ▼
Merge
   │
   ▼
Test
   │
   ▼
Validate
```

Every accepted rule change should have a clear migration or business reason.

---

# Rules & Admin Workflow

The administrative features work together:

```text
Rule Development
      │
      ▼
Rule History
      │
      ▼
Snapshots
      │
      ▼
Sync / Merge
      │
      ▼
Testing
      │
      ▼
Validated Configuration
```

These controls help maintain reliable Rule Configurations as migration development progresses.