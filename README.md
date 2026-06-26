# Turbonomic Stale Policy Audit

Non-invasive Python script to audit Turbonomic policies and identify candidates for cleanup, review, or retention.

The script connects to the Turbonomic REST API, reads policy metadata, target status, action history, audit information when available, and inventory counts. It generates an Excel workbook with a summary and detailed policy classification.

The script does **not** modify Turbonomic. It does not create, update, delete, or execute any action.

---

## Purpose

This tool helps identify stale, disabled, unused, empty-scope, or potentially obsolete Turbonomic policies.

Typical use cases:

* Review old or unused automation policies.
* Detect disabled policies that may no longer be needed.
* Identify policies with empty scope.
* Identify policies without recent activity.
* Review VMware-focused policy configuration in on-premises environments.
* Export policy findings to Excel for manual validation.

---

## Tested Scenario

This version was tested against an on-premises Turbonomic instance with a VMware vCenter target.

Example successful run:

```text
Authentication successful

-- VMware vCenter Targets ---------------------------
  Total VMware targets: 1
  WARNING: plvcenter02.example.local: Discovered

-- VMware Entity Inventory --------------------------
  Total VMware entities: 4

-- Audit Log (last 90 days) -------------------------
  Audit log unavailable: 404
  Audit entries retrieved: 0
  Policies with activity: 0

-- Automation Policies ------------------------------
  Total policies: 137
  VMware-specific: 0
  Other entity types: 137

-- Settings Policies --------------------------------
  Total settings policies: 78
  VMware-specific: 47

AUDIT SUMMARY
  TOTAL flagged     : 215
    VMware-specific : 47
    Other types     : 168

  DELETE (safe)     : 27
  REVIEW            : 161
  KEEP              : 27
```

---

## Required Turbonomic Role

The minimum tested role is:

```text
Operational Observer
```

A plain `Observer` role may authenticate successfully but can return HTTP `403` when reading policy endpoints.

The script needs read access to policies, settings policies, targets, search, actions, and optionally audit logs.

Recommended access:

* Role: `Operational Observer`
* Scope: environment, group, or target scope that includes the policies and entities to be audited

The script is read-only, but the user must still be allowed to view policies through the Turbonomic UI/API.

---

## API Endpoints Used

The script uses the following Turbonomic API endpoints:

```text
POST /api/v3/login
GET  /api/v3/targets
GET  /api/v3/search
GET  /api/v3/audit
GET  /api/v3/policies
GET  /api/v3/settingspolicies
GET  /api/v3/actions
```

Endpoint purpose:

| Endpoint                   | Purpose                                    |
| -------------------------- | ------------------------------------------ |
| `/api/v3/login`            | Authenticate to Turbonomic                 |
| `/api/v3/targets`          | Read configured targets, including vCenter |
| `/api/v3/search`           | Count VMware-related entities              |
| `/api/v3/audit`            | Read audit activity when available         |
| `/api/v3/policies`         | Read placement policies                    |
| `/api/v3/settingspolicies` | Read automation/settings policies          |
| `/api/v3/actions`          | Check recent actions related to policies   |

Note: `/api/v3/audit` may return `404` depending on the Turbonomic version, deployment, configuration, or user permissions. The script continues running if audit log data is unavailable, but activity-based classification will be less accurate.

---

## VMware Target Detection

Some Turbonomic versions return VMware vCenter targets as:

```text
category = Hypervisor
type     = vCenter
status   = Discovered
```

Because of that, the script does not rely only on the string `vmware` in the target type. It also checks for `vCenter`, `Hypervisor`, and the target address/name.

Target status values such as `Discovered` can appear even when the target is usable. Validate target health in the Turbonomic UI before treating it as disconnected.

---

## Requirements

### Python

Python 3.8 or later is recommended.

Required Python libraries:

```bash
pip3 install requests openpyxl urllib3
```

Depending on the environment, these libraries may already be installed.

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

Make the script executable if desired:

```bash
chmod +x stale_policies.py
```

---

## Usage

Basic execution:

```bash
python3 stale_policies.py \
  --host https://your-turbonomic-instance.example.com \
  --user operational_observer_user \
  --password 'YourPassword'
```

Recommended execution with explicit output:

```bash
python3 stale_policies.py \
  --host https://your-turbonomic-instance.example.com \
  --user operational_observer_user \
  --password 'YourPassword' \
  --days 90 \
  --output /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx \
  --verbose
```

---

## Parameters

| Parameter    | Required | Default                      | Description                         |
| ------------ | -------: | ---------------------------- | ----------------------------------- |
| `--host`     |      Yes | N/A                          | Turbonomic base URL                 |
| `--user`     |      Yes | N/A                          | Turbonomic API username             |
| `--password` |      Yes | N/A                          | Turbonomic password                 |
| `--days`     |       No | `90`                         | Inactivity threshold in days        |
| `--output`   |       No | `vmware_stale_policies.xlsx` | Excel output file                   |
| `--verbose`  |       No | disabled                     | Print additional processing details |

---

## Output

The script generates an Excel workbook.

Default output:

```text
vmware_stale_policies.xlsx
```

If `--output` is an absolute path, the file is written to that location.

If `--output` is only a filename, the script writes the report under:

```text
~/turbo-audit/
```

Depending on the local copy logic, the script may also copy the file one directory above the execution path.

Example:

```text
Report exported: /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx
Copied to: /opt/turbonomic/vmware_stale_policies.xlsx
```

---

## Excel Sheets

The workbook includes these sheets:

| Sheet                | Description                                                                      |
| -------------------- | -------------------------------------------------------------------------------- |
| `Summary`            | General execution summary, target count, entity count, and classification totals |
| `DELETE`             | Policies classified as likely safe to remove                                     |
| `REVIEW`             | Policies that require manual validation                                          |
| `All Policies`       | Full list of flagged policies                                                    |
| `VMware Environment` | VMware/vCenter target and inventory context                                      |

---

## Classification Logic

Each policy is classified into one of three categories:

| Classification | Meaning                                  |
| -------------- | ---------------------------------------- |
| `DELETE`       | Candidate for deletion after validation  |
| `REVIEW`       | Requires manual review before any action |
| `KEEP`         | Should be retained                       |

Common reasons detected:

* Disabled policy
* No recent actions in the selected period
* Not found in audit log
* Empty scope
* Test/demo/sandbox naming pattern
* Auto-generated or orphan-like naming pattern
* Built-in/default policy

Important: classification is advisory only. Deletions must always be reviewed and executed manually by a Turbonomic administrator.

---

## Non-Invasive Behavior

The script performs read-only operations.

| Operation              |           Performed |
| ---------------------- | ------------------: |
| Read targets           |                 Yes |
| Read entity counts     |                 Yes |
| Read policies          |                 Yes |
| Read settings policies |                 Yes |
| Read action history    |                 Yes |
| Read audit log         | Yes, when available |
| Create policies        |                  No |
| Modify policies        |                  No |
| Delete policies        |                  No |
| Execute actions        |                  No |

---

## Troubleshooting

### Authentication works but policies are not exported

If authentication succeeds but policies are not listed, check the API status codes.

A plain `Observer` role may return:

```text
HTTP 403
```

for:

```text
/api/v3/policies
/api/v3/settingspolicies
```

Use a user with the `Operational Observer` role or higher.

---

### Audit log unavailable: 404

This message means the script could not read:

```text
/api/v3/audit
```

Example:

```text
Audit log unavailable: 404
```

The script continues running, but the fields related to audit activity may be empty or less accurate.

---

### Proxy issues

If the environment has proxy variables configured and the Turbonomic URL should be accessed directly, unset proxy variables or configure `NO_PROXY`.

Example:

```bash
unset http_proxy
unset https_proxy
unset HTTP_PROXY
unset HTTPS_PROXY
unset all_proxy
unset ALL_PROXY

export NO_PROXY="your-turbonomic-instance.example.com,.example.com,localhost,127.0.0.1"
export no_proxy="$NO_PROXY"
```

The script also disables environment proxy usage through `requests` when `session.trust_env = False` is configured.

---

### Non-UTF-8 Python source error

If Python returns an error like:

```text
SyntaxError: Non-UTF-8 code starting with ...
```

convert the script to UTF-8:

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

## Example Run

```bash
python3 stale_policies.py \
  --host https://turbonomic.example.com \
  --user operational_observer \
  --password 'password' \
  --days 90 \
  --output /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx \
  --verbose
```

Expected result:

```text
Authentication successful
Total VMware targets: 1
Total policies: 137
Total settings policies: 78
TOTAL flagged: 215
Report exported: /opt/turbonomic/turbo-audit/vmware_stale_policies.xlsx
```

---

## Security Notes

Avoid passing passwords directly in shell history on shared systems.

For safer execution, use a temporary environment variable:

```bash
export TURBO_PASSWORD='YourPassword'

python3 stale_policies.py \
  --host https://your-turbonomic-instance.example.com \
  --user operational_observer_user \
  --password "$TURBO_PASSWORD"

unset TURBO_PASSWORD
```

---

## License

MIT
