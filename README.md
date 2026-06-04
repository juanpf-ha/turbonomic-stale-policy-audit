# Turbonomic Stale Policy Audit

A non-invasive Python script that connects to the Turbonomic REST API to identify unused, orphaned, or obsolete policies across any environment — VMware on-premises, Azure, AWS, GCP, or hybrid.

---

## What It Does

The script queries the following Turbonomic API endpoints:

- `/api/v3/policies` — automation policies (resize, move, suspend actions)
- `/api/v3/settingspolicies` — settings policies (entity configuration, cloud tier exclusions)
- `/api/v3/actions` — recent actions generated per policy
- `/api/v3/audit` — audit log of platform activity
- `/api/v3/targets` — connected targets (vCenter, Azure, AWS, GCP)
- `/api/v3/search` — entity inventory counts

Each policy is classified into one of three categories:

| Classification | Meaning |
|---|---|
| **DELETE** | Safe to remove — disabled test policies, orphaned Azure/GCP resources, auto-generated policies from disconnected targets |
| **REVIEW** | Requires manual verification — empty scope, no recent actions, absent from audit log |
| **KEEP** | Retain — Turbonomic built-in defaults |

Output is an Excel workbook with four sheets: **Summary**, **DELETE**, **REVIEW**, and **All Policies**.

---

## Requirements

### Turbonomic Access

- **Minimum role required:** `Observer` (read-only)
- The script performs **no write operations** — it only reads policy metadata, action history, audit logs, and inventory counts
- Compatible with **Turbonomic 8.x** (on-premises or SaaS)
- No agents, no probes, no configuration changes required

### Python Environment

- Python 3.8 or later
- Library: `openpyxl`

```bash
pip3 install openpyxl
```

---

## Usage

```bash
python3 stale_policies.py \
  --host https://your-turbonomic-instance.com \
  --user observer_user \
  --password YourPassword \
  --days 90 \
  --output client_name_stale_policies.xlsx
```

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `--host` | Yes | — | Turbonomic base URL |
| `--user` | Yes | — | API username (Observer role minimum) |
| `--password` | Yes | — | Password |
| `--days` | No | `90` | Inactivity threshold in days |
| `--output` | No | `vmware_stale_policies.xlsx` | Output filename |
| `--verbose` | No | off | Print each policy name during processing |

---

## Output

The script exports an Excel file to the current directory and copies it to `~/Downloads`.

### Sheet: Summary
High-level counts with environment context: connected VMware targets, entity inventory, and policy audit results by classification.

### Sheet: DELETE
Policies safe to remove with justification. Typical findings:
- Disabled test/sandbox policies created by individual users
- Auto-generated Cloud Tier Exclusion policies from disconnected Azure or AWS targets
- Orphaned AzureScaleSet and AvailabilitySet policies pointing to deleted resources

### Sheet: REVIEW
Policies requiring a human decision before acting:
- Policies with empty scope (no entities assigned)
- Policies absent from the audit log for the entire analysis period
- Disabled policies that may be intentional (maintenance windows, seasonal schedules)

### Sheet: All Policies
Complete list color-coded by classification: red = DELETE, yellow = REVIEW, green = KEEP.

---

## What "Non-Invasive" Means

| Action | Performed |
|---|---|
| Read policy list | Yes |
| Read action history | Yes |
| Read audit log | Yes |
| Read target status | Yes |
| Read entity counts | Yes |
| Create policies | **No** |
| Modify policies | **No** |
| Delete policies | **No** |
| Execute actions | **No** |

The script never modifies the Turbonomic environment. All deletions identified must be performed manually by an authorized administrator after reviewing the report.

---

## Customizing Classification Patterns

Edit these lists at the top of the script before running in a new client environment:

```python
TEST_NAME_PATTERNS = [
    "test", "demo", "temp", "tmp", "sandbox",
    "poc", "pilot", "trial", "dev-", "qa-", "staging-"
]

AUTOGEN_PATTERNS = [
    "cloud compute tier exclusion policy",
    "consistent scaling policy",
    "rds tier exclusion",
    "rds performance insights",
    "azure app service plan"
]

ORPHAN_PATTERNS = [
    "availabilityset::",
    "azurescaleset::"
]
```

---

## Tested On

- Turbonomic 8.x on-premises and SaaS (IBM Turbonomic)
- Python 3.10, 3.11, 3.13
- macOS, Linux

---

## License

MIT
