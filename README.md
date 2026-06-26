# Turbonomic Stale Policy Audit

Python script to audit Turbonomic policies and generate an Excel report focused on stale, risky, duplicated, empty-scope, or review-worthy policies.

The script is designed for VMware on-premises Turbonomic environments, but it also reports non-VMware policy rows separately when returned by the API.

The script is **read-only**. It does not modify, delete, create, enable, disable, or execute Turbonomic policies or actions.

---

## Purpose

This tool helps review Turbonomic policies in a safer and more structured way.

Main goals:

* Inventory Turbonomic policies.
* Identify policies that require manual review.
* Detect policies with empty scopes or empty groups.
* Compare custom settings policies against default settings policies.
* Identify action modes such as `AUTOMATIC`, `MANUAL`, `RECOMMEND`, `DISABLED`, or `EXTERNAL_APPROVAL`.
* Detect possible policy conflicts.
* Use audit logs when available.
* Save historical snapshots between executions.
* Export a simplified Excel workbook for operational review.

The script is intentionally conservative. It does **not** aggressively classify policies as deletion candidates unless there is strong evidence.

---

## Tested Environment

Tested against an on-premises Turbonomic instance with a VMware vCenter target.

Example validated inventory output:

```text
VirtualMachine: 8612
PhysicalMachine: 365
Storage: 922
Cluster: 42
VirtualDataCenter: 46
Total VMware entities: 9987
```

Example policy output:

```text
Total settings policies: 78
VMware-specific: 47

Total policies: 137
VMware-specific by entityType: 0
Other entity types: 137
```

Example audit log output with administrative privileges:

```text
Admin audit logs downloaded: /opt/turbonomic/turbo-audit/auditlogs/auditlog.tar.gz
Parsed audit text files: 1
Audit matches: 40
```

---

## Required Roles and Permissions

### Basic Policy Audit

Minimum role validated for policy read access:

```text
OPERATIONAL_OBSERVER
```

A plain `Observer` role may authenticate successfully but can return `HTTP 403` when trying to read policy endpoints.

The user must be able to read:

```text
/api/v3/targets
/api/v3/search
/api/v3/policies
/api/v3/settingspolicies
/api/v3/actions
```

Recommended minimum for normal report generation:

```text
Operational Observer
```

### Admin Audit Logs

Audit log analysis uses:

```text
/api/v3/admin/auditlogs?days=N
```

This endpoint is under the Admin API and returns a compressed audit log file.

Minimum validated role for admin audit log download:

```text
ADMINISTRATOR
```

If the user does not have sufficient privileges, the script can still generate the policy report, but audit log analysis will be skipped or marked as unavailable.

Typical error without sufficient privileges:

```text
HTTP 403 Access is denied
```

---

## API Endpoints Used

The script uses these Turbonomic API endpoints:

```text
POST /api/v3/login
GET  /api/v3/targets
GET  /api/v3/search
GET  /api/v3/policies
GET  /api/v3/settingspolicies
GET  /api/v3/actions
GET  /api/v3/admin/auditlogs
```

Endpoint usage:

| Endpoint                                      | Purpose                                   |
| --------------------------------------------- | ----------------------------------------- |
| `/api/v3/login`                               | Authenticate to Turbonomic                |
| `/api/v3/targets`                             | Detect VMware/vCenter targets             |
| `/api/v3/search`                              | Count VMware inventory entities           |
| `/api/v3/settingspolicies`                    | Read automation/settings policies         |
| `/api/v3/settingspolicies?only_defaults=true` | Read default automation/settings policies |
| `/api/v3/policies`                            | Read placement/policy endpoint results    |
| `/api/v3/actions`                             | Best-effort action check                  |
| `/api/v3/admin/auditlogs?days=N`              | Optional admin audit log download         |

Important note:

```text
/api/v3/audit
```

is not used. In the tested environment it returned:

```text
HTTP 404
Resource: /audit not found
```

The valid audit log endpoint was:

```text
/api/v3/admin/auditlogs
```

---

## VMware Target Detection

Turbonomic may report VMware vCenter targets as:

```text
category = Hypervisor
type     = vCenter
status   = Discovered
```

The script detects VMware/vCenter targets using a combination of:

* target category
* target type
* display name
* address/name fields

Accepted target states include:

```text
Validated
Discovered
```

A `Discovered` target should still be validated in the Turbonomic UI before assuming full target health.

---

## VMware Inventory Counting

The script counts VMware entities through:

```text
/api/v3/search?types=<EntityType>&limit=500&cursor=<N>
```

The tested Turbonomic environment paginates search results using a numeric `cursor`.

Example:

```text
cursor=0
cursor=500
cursor=1000
cursor=1500
...
```

The script keeps requesting pages until the last page returns fewer than the page size.

Default inventory parameters:

```text
--inventory-page-size 500
--inventory-max-pages 1000
```

This avoids the previous issue where only the first page of 500 entities was counted.

VMware-related entity types currently counted:

```text
VirtualMachine
PhysicalMachine
Storage
Datacenter
Cluster
VirtualDataCenter
ResourcePool
```

Note: in many Turbonomic VMware environments, `Datacenter` may return zero while `VirtualDataCenter` returns the expected VMware datacenter objects.

---

## Installation

Clone the repository:

```bash
git clone https://github.com/juanpf-ha/turbonomic-stale-policy-audit.git
cd turbonomic-stale-policy-audit
```

Install dependencies:

```bash
pip3 install requests openpyxl urllib3
```

Optional:

```bash
chmod +x stale_policies.py
```

---

## Usage

### Basic Run

```bash
python3 stale_policies.py \
  --host https://turbonomic.example.com/ \
  --user operational_observer_user \
  --password 'PASSWORD' \
  --days 90 \
  --output /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx
```

### Run with Admin Audit Logs

Requires an account with administrative privileges.

```bash
python3 stale_policies.py \
  --host https://turbonomic.example.com/ \
  --user admin_user \
  --password 'PASSWORD' \
  --days 90 \
  --use-admin-auditlogs \
  --admin-audit-days 90 \
  --output /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx \
  --verbose
```

### Run with Detail Sheets

By default, the Excel report is simplified.

To include technical detail sheets:

```bash
python3 stale_policies.py \
  --host https://turbonomic.example.com/ \
  --user admin_user \
  --password 'PASSWORD' \
  --days 90 \
  --use-admin-auditlogs \
  --admin-audit-days 90 \
  --include-detail-sheets \
  --output /opt/turbonomic/turbo-audit/vmware_stale_policies_detail.xlsx \
  --verbose
```

---

## Parameters

| Parameter                        | Required | Default                      | Description                                              |
| -------------------------------- | -------: | ---------------------------- | -------------------------------------------------------- |
| `--host`                         |      Yes | N/A                          | Turbonomic base URL                                      |
| `--user`                         |      Yes | N/A                          | Turbonomic username                                      |
| `--password`                     |      Yes | N/A                          | Turbonomic password                                      |
| `--days`                         |       No | `90`                         | Inactivity threshold                                     |
| `--output`                       |       No | `vmware_stale_policies.xlsx` | Excel output path                                        |
| `--verbose`                      |       No | disabled                     | Print detailed execution information                     |
| `--snapshot-dir`                 |       No | `snapshots`                  | Directory for policy snapshots                           |
| `--no-snapshot`                  |       No | disabled                     | Disable snapshot creation                                |
| `--skip-group-check`             |       No | disabled                     | Skip group/scope validation                              |
| `--skip-action-check`            |       No | disabled                     | Skip action checks                                       |
| `--trust-env-proxy`              |       No | disabled                     | Allow Python requests to use proxy environment variables |
| `--use-admin-auditlogs`          |       No | disabled                     | Enable admin audit log download and parsing              |
| `--admin-audit-days`             |       No | same as `--days`             | Number of days to request from admin audit logs          |
| `--audit-artifact-dir`           |       No | `auditlogs`                  | Directory to store downloaded audit logs                 |
| `--max-audit-matches-per-policy` |       No | script default               | Maximum audit matches stored per policy                  |
| `--inventory-page-size`          |       No | `500`                        | Page size for inventory search                           |
| `--inventory-max-pages`          |       No | `1000`                       | Max pages per entity type                                |
| `--include-detail-sheets`        |       No | disabled                     | Include technical Excel detail sheets                    |

---

## Excel Output

The script generates an Excel workbook.

Default simplified workbook includes:

| Sheet                | Description                                 |
| -------------------- | ------------------------------------------- |
| `Summary`            | General execution summary                   |
| `REVIEW`             | Policies that require manual review         |
| `KEEP`               | Policies that should likely be retained     |
| `INFO`               | Informational rows, usually non-actionable  |
| `All Policies`       | Consolidated simplified policy view         |
| `VMware Environment` | vCenter target and VMware inventory context |

Sheets with no rows are not created.

For example, if there are no `CANDIDATE_DELETE` rows, the `CANDIDATE_DELETE` sheet will not appear.

### Optional Detail Sheets

When using:

```text
--include-detail-sheets
```

the workbook may also include:

| Sheet               | Description                                  |
| ------------------- | -------------------------------------------- |
| `Action Modes`      | Detailed action mode settings                |
| `Scope Analysis`    | Scope and group analysis                     |
| `Policy Conflicts`  | Possible overlapping or conflicting settings |
| `Audit Log Matches` | Audit log matches by policy                  |
| `Snapshot Changes`  | Differences versus previous snapshot         |

---

## Simplified Report Columns

The main policy sheets are intentionally simplified.

Typical columns:

| Column           | Description                             |
| ---------------- | --------------------------------------- |
| `classification` | `REVIEW`, `KEEP`, `INFO`, etc.          |
| `policy`         | Policy display name                     |
| `source`         | `settings_policy` or `placement_policy` |
| `entity`         | Entity type if available                |
| `enabled`        | Whether the policy is enabled           |
| `scope`          | Human-readable scope summary            |
| `settings`       | Human-readable settings summary         |
| `audit`          | Audit evidence summary                  |
| `recommendation` | Reason for the classification           |

Technical counters such as `settings_diff_count`, `settings_compared`, `unresolved_scopes`, and similar fields are hidden from the main sheets to keep the report readable.

Use `--include-detail-sheets` when detailed troubleshooting is needed.

---

## Classification Logic

The script uses conservative classifications.

| Classification     | Meaning                                              |
| ------------------ | ---------------------------------------------------- |
| `CANDIDATE_DELETE` | Strong candidate for cleanup after manual validation |
| `REVIEW`           | Requires manual review                               |
| `KEEP`             | Should likely be retained                            |
| `INFO`             | Informational row                                    |
| `UNKNOWN`          | Fallback classification, created only if needed      |

The script does not classify policies as deletion candidates based only on missing audit logs or missing action history.

Deletion candidates require multiple strong signals, such as:

* non-default policy
* disabled
* empty scope or unresolved scope
* empty groups
* no meaningful settings difference from default
* test/demo/tmp/sandbox-like naming

Manual validation is always required before deleting anything in Turbonomic.

---

## Audit Log Analysis

Audit log analysis is optional and disabled by default.

Enable it with:

```text
--use-admin-auditlogs
```

The script downloads:

```text
/api/v3/admin/auditlogs?days=N
```

The response is expected to be a compressed file, usually gzip/tar.gz.

The script stores it under:

```text
auditlogs/
```

Example:

```text
/opt/turbonomic/turbo-audit/auditlogs/auditlog.tar.gz
```

The audit log parser is best-effort. It searches audit text for policy UUIDs and policy names.

Audit log evidence is used as supporting context, not as the only classification signal.

---

## Snapshot History

The script saves a JSON snapshot for each execution unless disabled.

Default location:

```text
snapshots/
```

Example:

```text
/opt/turbonomic/turbo-audit/snapshots/policy_snapshot_20260626_021243.json
```

Snapshots allow future comparison between executions:

* new policies
* removed policies
* changed scopes
* changed settings
* changed action modes
* changed enabled/disabled state

Disable snapshots with:

```text
--no-snapshot
```

---

## Proxy Handling

By default, the script disables inherited proxy settings for Python `requests`.

This helps avoid issues where internal Turbonomic traffic is incorrectly sent through a corporate proxy.

If proxy environment variables must be used, run with:

```text
--trust-env-proxy
```

Manual proxy cleanup example:

```bash
unset http_proxy
unset https_proxy
unset HTTP_PROXY
unset HTTPS_PROXY
unset all_proxy
unset ALL_PROXY

export NO_PROXY="turbonomic.example.com,.example.com,localhost,127.0.0.1"
export no_proxy="$NO_PROXY"
```

---

## Troubleshooting

### Authentication succeeds but policies return HTTP 403

The user can authenticate but does not have enough permissions to read policy endpoints.

Typical issue:

```text
/api/v3/settingspolicies -> HTTP 403
/api/v3/policies -> HTTP 403
```

Use a user with at least the validated `OPERATIONAL_OBSERVER` role for normal policy audit.

---

### Admin audit logs return HTTP 403

The user does not have enough privileges to use:

```text
/api/v3/admin/auditlogs
```

Use an account with administrative privileges.

The report can still be generated without audit logs by removing:

```text
--use-admin-auditlogs
```

---

### `/api/v3/audit` returns HTTP 404

This is expected in the tested environment.

The script does not use:

```text
/api/v3/audit
```

Use the admin audit log endpoint instead:

```text
/api/v3/admin/auditlogs
```

---

### Inventory count shows exactly 500

This usually means only the first page was counted.

Current versions of the script use numeric `cursor` pagination:

```text
cursor=0
cursor=500
cursor=1000
...
```

If counts are still exactly 500, validate:

```bash
python3 stale_policies.py --help | grep inventory
```

Expected parameters:

```text
--inventory-page-size
--inventory-max-pages
```

---

### Python encoding error

If Python returns:

```text
SyntaxError: Non-UTF-8 code starting with ...
```

convert the file to UTF-8:

```bash
python3 - <<'PY'
from pathlib import Path

p = Path("stale_policies.py")
raw = p.read_bytes()

for enc in ("utf-8", "cp1252", "latin-1"):
    try:
        text = raw.decode(enc)
        break
    except UnicodeDecodeError:
        continue
else:
    raise SystemExit("Unable to detect encoding")

text = text.replace("\r\n", "\n").replace("\r", "\n")
p.write_text(text, encoding="utf-8")
print("Converted stale_policies.py to UTF-8")
PY

python3 -m py_compile stale_policies.py
```

---

## Recommended Execution Example

```bash
cd /opt/turbonomic/turbo-audit

python3 stale_policies.py \
  --host https://turbonomic.example.com/ \
  --user admin_user \
  --password 'PASSWORD' \
  --days 90 \
  --use-admin-auditlogs \
  --admin-audit-days 90 \
  --output /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx \
  --verbose
```

Expected output example:

```text
Authentication successful

-- VMware vCenter Targets ---------------------------
Total VMware targets: 1
All VMware/vCenter targets are in an accepted state

-- VMware Entity Inventory --------------------------
VirtualMachine: 8612
PhysicalMachine: 365
Storage: 922
Cluster: 42
VirtualDataCenter: 46
Total VMware entities: 9987

-- Automation / Settings Policies -------------------
Total settings policies: 78
VMware-specific: 47

-- Admin Audit Logs ---------------------------------
Admin audit logs downloaded
Parsed audit text files: 1

AUDIT SUMMARY
TOTAL rows       : 215
REVIEW           : 19
KEEP             : 55
INFO             : 141
Audit matches    : 40

Report exported: /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx
```

---

## Safety Notes

This script is read-only.

It does not:

* delete policies
* disable policies
* enable policies
* change scopes
* modify settings
* execute actions

Any cleanup decision must be validated manually in the Turbonomic UI before applying changes.

---

## Known Limitations

* Admin audit logs require administrative privileges.
* Audit log parsing is best-effort because logs are exported as compressed text, not structured policy objects.
* Placement policies may not expose `entityType` directly. Future versions may classify placement policies by resolving `consumerGroup`, `providerGroup`, and `mergeGroups`.
* Policy conflicts are advisory and may include expected overlap between defaults, scopes, and custom policies.
* VMware inventory counts depend on `/api/v3/search` pagination behavior.

---

## Suggested Review Workflow

Recommended Excel review order:

1. `Summary`
2. `REVIEW`
3. `All Policies`
4. `VMware Environment`
5. Optional: run again with `--include-detail-sheets`
6. Optional detail review:

   * `Policy Conflicts`
   * `Action Modes`
   * `Scope Analysis`
   * `Audit Log Matches`

The `REVIEW` sheet should be the main operational starting point.

---

## License

MIT
