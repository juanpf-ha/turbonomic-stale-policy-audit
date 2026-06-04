#!/usr/bin/env python3
"""
Turbonomic Stale Policy Audit - VMware On-Premises Focus
Optimized for environments monitoring VMware workloads in on-premises datacenters
"""

import requests
import urllib3
import argparse
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

parser = argparse.ArgumentParser(
    description="Turbonomic Stale Policy Audit - VMware On-Premises Focus"
)
parser.add_argument("--host", required=True, help="Turbonomic host URL")
parser.add_argument("--user", required=True, help="API username")
parser.add_argument("--password", required=True, help="API password")
parser.add_argument("--days", default=90, type=int, help="Inactivity threshold in days")
parser.add_argument("--output", default="vmware_stale_policies.xlsx", help="Output Excel file")
parser.add_argument("--verbose", action="store_true", help="Enable verbose output")
args = parser.parse_args()

TURBO_HOST = args.host.rstrip("/")
USERNAME = args.user
PASSWORD = args.password
INACTIVITY_DAYS = args.days
OUTPUT_FILE = args.output
VERBOSE = args.verbose

VMWARE_ENTITY_TYPES = [
    "VirtualMachine", "PhysicalMachine", "Storage", "Datacenter",
    "Cluster", "Host", "Datastore", "VirtualDatacenter", "ResourcePool"
]

TEST_NAME_PATTERNS = [
    "test", "demo", "temp", "tmp", "sandbox", "prueba", "ejemplo",
    "poc", "pilot", "trial", "dev-", "qa-", "staging-"
]

VMWARE_AUTOGEN_PATTERNS = [
    "vcenter", "esxi", "vsan", "nsx", "vrops", "vrealize",
    "default placement", "drs", "ha cluster", "storage drs"
]

ORPHAN_PATTERNS = [
    "::deleted", "::removed", "::migrated", "::decommissioned",
    "old-vcenter", "legacy-", "retired-"
]

def classify_vmware_policy(name, enabled, entity_type, reasons_str, scope_count, last_modified):
    nl = name.lower()
    if not enabled:
        for pattern in TEST_NAME_PATTERNS:
            if pattern in nl:
                return "DELETE", "Disabled test/sandbox policy - safe to remove"
        return "REVIEW", "Manually disabled - verify if intentional before removing"
    for pattern in ORPHAN_PATTERNS:
        if pattern in nl:
            return "DELETE", "Orphaned VMware resource - entity no longer exists"
    if "empty scope" in reasons_str or scope_count == 0:
        for pattern in VMWARE_AUTOGEN_PATTERNS:
            if pattern in nl:
                return "DELETE", "Auto-generated VMware policy with empty scope - target likely disconnected"
        if name.endswith("Defaults") or "default" in nl:
            return "KEEP", "Turbonomic default policy - remove only if VMware entity type is not monitored"
        return "REVIEW", "Empty scope - verify if associated VMware cluster/datacenter is still active"
    if entity_type not in VMWARE_ENTITY_TYPES and entity_type != "UNKNOWN":
        if "no actions" in reasons_str and "not found in audit" in reasons_str:
            return "REVIEW", f"Non-VMware entity type ({entity_type}) with no activity - may be obsolete"
    if last_modified:
        try:
            mod_date = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
            age_days = (datetime.now(mod_date.tzinfo) - mod_date).days
            if age_days > 365 and "no actions" in reasons_str:
                return "REVIEW", f"Policy not modified in {age_days} days and no recent actions - likely obsolete"
        except:
            pass
    return "REVIEW", "Review manually - may still be relevant for VMware environment"

def analyze_vmware_targets(session):
    print("\n── VMware vCenter Targets ───────────────────────────")
    try:
        r = session.get(f"{TURBO_HOST}/api/v3/targets")
        if r.status_code == 200:
            targets = r.json()
            vmware_targets = [t for t in targets if t.get("category") == "Hypervisor"
                            and "vmware" in t.get("type", "").lower()]
            print(f"  Total VMware targets: {len(vmware_targets)}")
            disconnected = []
            for target in vmware_targets:
                status = target.get("status", "UNKNOWN")
                name = target.get("displayName", "N/A")
                if status != "VALIDATED":
                    disconnected.append({"name": name, "status": status, "uuid": target.get("uuid")})
                    print(f"  WARNING: {name}: {status}")
            if not disconnected:
                print("  All VMware targets are connected")
            return vmware_targets, disconnected
        else:
            print(f"  Unable to retrieve targets: {r.status_code}")
            return [], []
    except Exception as e:
        print(f"  Error analyzing VMware targets: {e}")
        return [], []

def get_vmware_entity_counts(session):
    print("\n── VMware Entity Inventory ──────────────────────────")
    entity_counts = {}
    for entity_type in VMWARE_ENTITY_TYPES:
        try:
            r = session.get(f"{TURBO_HOST}/api/v3/search",
                          params={"types": entity_type, "limit": 1})
            if r.status_code == 200:
                data = r.json()
                count = data.get("count", 0) if isinstance(data, dict) else len(data)
                entity_counts[entity_type] = count
                if VERBOSE:
                    print(f"  {entity_type}: {count}")
        except:
            entity_counts[entity_type] = 0
    total = sum(entity_counts.values())
    print(f"  Total VMware entities: {total}")
    return entity_counts

def generate_excel_report(results, vmware_targets, disconnected_targets, entity_counts):
    COLORS = {"DELETE": "FFCCCC", "REVIEW": "FFF2CC", "KEEP": "CCFFCC"}
    HDR_FILL = PatternFill("solid", fgColor="1F4E79")
    HDR_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    DATA_FONT = Font(name="Arial", size=9)
    headers = [
        "Action", "Justification", "Source", "Policy Name", "Entity Type",
        "Enabled", "Scope Count", "In Audit Log", "Last Modified", "Reasons Detected", "ID"
    ]
    col_widths = [10, 60, 18, 50, 20, 10, 12, 13, 20, 50, 38]
    order = ["DELETE", "REVIEW", "KEEP"]

    wb = Workbook()

    ws_summary = wb.active
    ws_summary.title = "Summary"
    summary_data = [
        ["Turbonomic Stale Policy Audit - VMware On-Premises"], [],
        ["Host", TURBO_HOST],
        ["Analysis Date", datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")],
        ["Inactivity Threshold", f"{INACTIVITY_DAYS} days"], [],
        ["VMware Environment"], [],
        ["Total vCenter Targets", len(vmware_targets)],
        ["Disconnected Targets", len(disconnected_targets)],
        ["Total VMware Entities", sum(entity_counts.values())], [],
    ]
    for etype, count in entity_counts.items():
        summary_data.append([f"  {etype}", count])
    summary_data.extend([
        [], ["Policy Audit Results"], [],
        ["Classification", "Count"],
        ["TOTAL Flagged", len(results)],
        ["DELETE (safe)", len([x for x in results if x["action"] == "DELETE"])],
        ["REVIEW (manual)", len([x for x in results if x["action"] == "REVIEW"])],
        ["KEEP (retain)", len([x for x in results if x["action"] == "KEEP"])],
    ])
    for row in summary_data:
        ws_summary.append(row)
    ws_summary["A1"].font = Font(bold=True, size=14, name="Arial", color="1F4E79")
    ws_summary["A6"].font = Font(bold=True, size=11, name="Arial", color="1F4E79")
    ws_summary.column_dimensions["A"].width = 35
    ws_summary.column_dimensions["B"].width = 50

    ws_delete = wb.create_sheet("DELETE")
    ws_delete.append(headers)
    for col, _ in enumerate(headers, 1):
        c = ws_delete.cell(row=1, column=col)
        c.fill = PatternFill("solid", fgColor="C00000")
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
    for rd in [x for x in results if x["action"] == "DELETE"]:
        ws_delete.append([rd["action"], rd["justification"], rd["source"], rd["name"],
            rd["entityType"], rd["enabled"], rd["scope_count"],
            rd["in_audit"], rd["last_modified"], rd["reasons"], rd["id"]])
        for col in range(1, len(headers)+1):
            ws_delete.cell(row=ws_delete.max_row, column=col).fill = PatternFill("solid", fgColor="FFCCCC")
            ws_delete.cell(row=ws_delete.max_row, column=col).font = DATA_FONT
            ws_delete.cell(row=ws_delete.max_row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(col_widths, 1):
        ws_delete.column_dimensions[get_column_letter(i)].width = w
    ws_delete.freeze_panes = "A2"

    ws_review = wb.create_sheet("REVIEW")
    ws_review.append(headers)
    for col, _ in enumerate(headers, 1):
        c = ws_review.cell(row=1, column=col)
        c.fill = PatternFill("solid", fgColor="F4B084")
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    for rd in [x for x in results if x["action"] == "REVIEW"]:
        ws_review.append([rd["action"], rd["justification"], rd["source"], rd["name"],
            rd["entityType"], rd["enabled"], rd["scope_count"],
            rd["in_audit"], rd["last_modified"], rd["reasons"], rd["id"]])
        for col in range(1, len(headers)+1):
            ws_review.cell(row=ws_review.max_row, column=col).fill = PatternFill("solid", fgColor="FFF2CC")
            ws_review.cell(row=ws_review.max_row, column=col).font = DATA_FONT
            ws_review.cell(row=ws_review.max_row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(col_widths, 1):
        ws_review.column_dimensions[get_column_letter(i)].width = w
    ws_review.freeze_panes = "A2"

    ws_all = wb.create_sheet("All Policies")
    ws_all.append(headers)
    for col, _ in enumerate(headers, 1):
        c = ws_all.cell(row=1, column=col)
        c.fill = HDR_FILL; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for rd in sorted(results, key=lambda x: order.index(x["action"])):
        ws_all.append([rd["action"], rd["justification"], rd["source"], rd["name"],
            rd["entityType"], rd["enabled"], rd["scope_count"],
            rd["in_audit"], rd["last_modified"], rd["reasons"], rd["id"]])
        fill = PatternFill("solid", fgColor=COLORS[rd["action"]])
        for col in range(1, len(headers)+1):
            ws_all.cell(row=ws_all.max_row, column=col).fill = fill
            ws_all.cell(row=ws_all.max_row, column=col).font = DATA_FONT
            ws_all.cell(row=ws_all.max_row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(col_widths, 1):
        ws_all.column_dimensions[get_column_letter(i)].width = w
    ws_all.freeze_panes = "A2"

    ws_vmware = wb.create_sheet("VMware Environment")
    ws_vmware.append(["VMware vCenter Targets"])
    ws_vmware.append(["Target Name", "Status", "UUID"])
    for target in vmware_targets:
        status = target.get("status", "UNKNOWN")
        ws_vmware.append([target.get("displayName", "N/A"), status, target.get("uuid", "N/A")])
        if status != "VALIDATED":
            for col in range(1, 4):
                ws_vmware.cell(row=ws_vmware.max_row, column=col).fill = PatternFill("solid", fgColor="FFCCCC")
    ws_vmware.column_dimensions["A"].width = 40
    ws_vmware.column_dimensions["B"].width = 20
    ws_vmware.column_dimensions["C"].width = 38

    wb.save(OUTPUT_FILE)

def main():
    print("=" * 70)
    print("  Turbonomic Stale Policy Audit - VMware On-Premises Focus")
    print("=" * 70)

    session = requests.Session()
    session.verify = False

    print(f"\nConnecting to {TURBO_HOST} ...")
    r_auth = session.post(f"{TURBO_HOST}/api/v3/login",
                         data={"username": USERNAME, "password": PASSWORD})
    if r_auth.status_code != 200:
        print(f"Authentication failed: {r_auth.status_code}")
        exit(1)
    print("Authentication successful")

    vmware_targets, disconnected_targets = analyze_vmware_targets(session)
    entity_counts = get_vmware_entity_counts(session)

    cutoff_date = datetime.utcnow() - timedelta(days=INACTIVITY_DAYS)
    cutoff_ms = int(cutoff_date.timestamp() * 1000)

    print(f"\n── Audit Log (last {INACTIVITY_DAYS} days) ──────────────────────────")
    policy_ids_in_audit = set()
    audit_count = 0
    try:
        offset = 0
        limit = 500
        while True:
            r_audit = session.get(f"{TURBO_HOST}/api/v3/audit",
                                params={"starttime": cutoff_ms, "limit": limit, "offset": offset})
            if r_audit.status_code == 200:
                audit_entries = r_audit.json()
                if not audit_entries:
                    break
                audit_count += len(audit_entries)
                for entry in audit_entries:
                    pid = entry.get("targetId") or entry.get("policyId") or entry.get("uuid")
                    if pid:
                        policy_ids_in_audit.add(str(pid))
                if len(audit_entries) < limit:
                    break
                offset += limit
            else:
                print(f"  Audit log unavailable: {r_audit.status_code}")
                break
        print(f"  Audit entries retrieved: {audit_count}")
        print(f"  Policies with activity: {len(policy_ids_in_audit)}")
    except Exception as e:
        print(f"  Audit log error: {e}")

    results = []

    print(f"\n── Automation Policies ──────────────────────────────")
    r = session.get(f"{TURBO_HOST}/api/v3/policies")
    if r.status_code == 200:
        policies = r.json()
        vmware_policies = [p for p in policies if p.get("entityType") in VMWARE_ENTITY_TYPES]
        print(f"  Total policies: {len(policies)}")
        print(f"  VMware-specific: {len(vmware_policies)}")
        print(f"  Other entity types: {len(policies) - len(vmware_policies)}")
        for i, p in enumerate(policies):
            pid = p.get("uuid")
            name = p.get("displayName", "N/A")
            enabled = p.get("enabled", True)
            etype = p.get("entityType", "UNKNOWN")
            last_mod = p.get("lastModified")
            scope = p.get("scope", [])
            scope_count = len(scope) if isinstance(scope, list) else 0
            if VERBOSE:
                print(f"  [{i+1}/{len(policies)}] {name}")
            ar = session.get(f"{TURBO_HOST}/api/v3/actions",
                           params={"policyId": pid, "starttime": cutoff_ms, "limit": 1})
            recent_actions = ar.json() if ar.ok and isinstance(ar.json(), list) else []
            in_audit = str(pid) in policy_ids_in_audit
            reasons = []
            if not enabled:
                reasons.append("disabled")
            if not recent_actions:
                reasons.append(f"no actions in {INACTIVITY_DAYS} days")
            if not in_audit:
                reasons.append("not found in audit log")
            if scope_count == 0:
                reasons.append("empty scope")
            if reasons:
                rs = "; ".join(reasons)
                action, just = classify_vmware_policy(name, enabled, etype, rs, scope_count, last_mod)
                results.append({"action": action, "justification": just,
                    "source": "automation_policy", "id": pid, "name": name,
                    "entityType": etype, "enabled": str(enabled), "scope_count": scope_count,
                    "in_audit": str(in_audit), "last_modified": last_mod or "N/A", "reasons": rs})

    print(f"\n── Settings Policies ────────────────────────────────")
    r2 = session.get(f"{TURBO_HOST}/api/v3/settingspolicies")
    if r2.status_code == 200:
        spolicies = r2.json()
        vmware_sp = [p for p in spolicies if p.get("entityType") in VMWARE_ENTITY_TYPES]
        print(f"  Total settings policies: {len(spolicies)}")
        print(f"  VMware-specific: {len(vmware_sp)}")
        for i, p in enumerate(spolicies):
            pid = p.get("uuid")
            name = p.get("displayName", "N/A")
            enabled = p.get("enabled", True)
            etype = p.get("entityType", "UNKNOWN")
            last_mod = p.get("lastModified")
            scope = p.get("scope", [])
            scope_count = len(scope) if isinstance(scope, list) else 0
            if VERBOSE:
                print(f"  [{i+1}/{len(spolicies)}] {name}")
            in_audit = str(pid) in policy_ids_in_audit
            reasons = []
            if not enabled:
                reasons.append("disabled")
            if scope_count == 0:
                reasons.append("empty scope")
            if not in_audit:
                reasons.append("not found in audit log")
            if reasons:
                rs = "; ".join(reasons)
                action, just = classify_vmware_policy(name, enabled, etype, rs, scope_count, last_mod)
                results.append({"action": action, "justification": just,
                    "source": "settings_policy", "id": pid, "name": name,
                    "entityType": etype, "enabled": str(enabled), "scope_count": scope_count,
                    "in_audit": str(in_audit), "last_modified": last_mod or "N/A", "reasons": rs})

    to_delete = [x for x in results if x["action"] == "DELETE"]
    to_review  = [x for x in results if x["action"] == "REVIEW"]
    to_keep    = [x for x in results if x["action"] == "KEEP"]
    vmware_flagged     = [x for x in results if x["entityType"] in VMWARE_ENTITY_TYPES]
    non_vmware_flagged = [x for x in results if x["entityType"] not in VMWARE_ENTITY_TYPES]

    print(f"\n{'=' * 70}")
    print(f"  AUDIT SUMMARY")
    print(f"{'=' * 70}")
    print(f"  Host: {TURBO_HOST}")
    print(f"  Inactivity threshold: {INACTIVITY_DAYS} days")
    print(f"  Analysis date: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    print(f"{'=' * 70}")
    print(f"  VMware targets    : {len(vmware_targets)} ({len(disconnected_targets)} disconnected)")
    print(f"  VMware entities   : {sum(entity_counts.values())}")
    print(f"{'=' * 70}")
    print(f"  TOTAL flagged     : {len(results)}")
    print(f"    VMware-specific : {len(vmware_flagged)}")
    print(f"    Other types     : {len(non_vmware_flagged)}")
    print(f"{'=' * 70}")
    print(f"  DELETE (safe)     : {len(to_delete)}")
    print(f"  REVIEW            : {len(to_review)}")
    print(f"  KEEP              : {len(to_keep)}")
    print(f"{'=' * 70}\n")

    generate_excel_report(results, vmware_targets, disconnected_targets, entity_counts)
    print(f"Report exported: ~/turbo-audit/{OUTPUT_FILE}")
    print("Sheets: 'Summary' | 'DELETE' | 'REVIEW' | 'All Policies' | 'VMware Environment'")

if __name__ == "__main__":
    main()

import shutil, os
_dest = os.path.expanduser(f"~/Downloads/{OUTPUT_FILE}")
shutil.copy2(OUTPUT_FILE, _dest)
print(f"Copied to Downloads: {_dest}")
