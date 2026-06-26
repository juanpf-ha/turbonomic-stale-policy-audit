#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Turbonomic Stale Policy Audit - VMware On-Premises Focus

Read-only audit tool for Turbonomic policies.

Main goals:
  * Inventory VMware/vCenter targets and VMware-related entities.
  * Read Automation / Settings Policies from /api/v3/settingspolicies.
  * Read Placement / Policy endpoint data from /api/v3/policies.
  * Compare custom settings policies against defaults when available.
  * Analyze scopes, action modes, duplicated/conflicting settings, and snapshots.
  * Optionally download and parse admin audit logs from /api/v3/admin/auditlogs.

Important:
  * This script is read-only. It does not modify or delete policies.
  * /api/v3/admin/auditlogs normally requires Administrator or Site Administrator
    privileges and returns application/gzip, not JSON.
  * If audit logs are unavailable, the script does not use that absence as a stale
    signal.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import io
import json
import os
import re
import shutil
import tarfile
import tempfile
import zipfile
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

VMWARE_ENTITY_TYPES = [
    "VirtualMachine",
    "PhysicalMachine",
    "Storage",
    "Datacenter",
    "Cluster",
    "VirtualDataCenter",
    "ResourcePool",
]

ACCEPTED_TARGET_STATES = {"validated", "discovered"}

ACTION_MODE_VALUES = {
    "AUTOMATIC",
    "MANUAL",
    "RECOMMEND",
    "DISABLED",
    "EXTERNAL_APPROVAL",
}

TEST_NAME_PATTERNS = [
    "test",
    "demo",
    "temp",
    "tmp",
    "poc",
    "sandbox",
    "prueba",
    "orphan",
    "old",
    "backup",
]

GENERIC_NAMES_TO_IGNORE_FOR_AUDIT = {
    "default",
    "global",
    "policy",
    "policies",
    "automation",
    "settings",
    "placement",
}

CLASS_ORDER = ["CANDIDATE_DELETE", "REVIEW", "KEEP", "INFO", "UNKNOWN"]
CLASS_COLORS = {
    "CANDIDATE_DELETE": "FFCCCC",
    "REVIEW": "FFF2CC",
    "KEEP": "CCFFCC",
    "INFO": "D9EAF7",
    "UNKNOWN": "E7E6E6",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read-only Turbonomic stale policy audit with VMware focus."
    )
    parser.add_argument("--host", required=True, help="Turbonomic base URL")
    parser.add_argument("--user", required=True, help="Turbonomic username")
    parser.add_argument("--password", required=True, help="Turbonomic password")
    parser.add_argument(
        "--days",
        type=int,
        default=90,
        help="Inactivity threshold in days for best-effort action/audit analysis",
    )
    parser.add_argument(
        "--output",
        default="vmware_stale_policies.xlsx",
        help="Excel output path. Relative paths are written under ~/turbo-audit.",
    )
    parser.add_argument("--verbose", action="store_true", help="Print verbose details")
    parser.add_argument(
        "--snapshot-dir",
        default=None,
        help="Directory for policy snapshots. Default: <output_dir>/snapshots",
    )
    parser.add_argument(
        "--no-snapshot", action="store_true", help="Disable snapshot save/compare"
    )
    parser.add_argument(
        "--skip-group-check",
        action="store_true",
        help="Skip best-effort group/scope lookup",
    )
    parser.add_argument(
        "--skip-action-check",
        action="store_true",
        help="Skip best-effort /api/v3/actions policy checks",
    )
    parser.add_argument(
        "--trust-env-proxy",
        action="store_true",
        help="Allow requests to honor http_proxy/https_proxy/NO_PROXY variables",
    )
    parser.add_argument(
        "--use-admin-auditlogs",
        action="store_true",
        help="Download and parse /api/v3/admin/auditlogs. Usually requires Administrator or Site Administrator.",
    )
    parser.add_argument(
        "--admin-audit-days",
        type=int,
        default=None,
        help="Number of days for /api/v3/admin/auditlogs. Default: same as --days.",
    )
    parser.add_argument(
        "--audit-artifact-dir",
        default=None,
        help="Directory to store downloaded audit log artifacts. Default: <output_dir>/auditlogs",
    )
    parser.add_argument(
        "--max-audit-matches-per-policy",
        type=int,
        default=10,
        help="Maximum audit log match lines kept per policy",
    )
    return parser.parse_args()


def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def now_stamp() -> str:
    return now_utc().strftime("%Y%m%d_%H%M%S")


def short_text(text: str, limit: int = 1000) -> str:
    text = text or ""
    return text if len(text) <= limit else text[:limit] + "..."


def normalize_url(host: str) -> str:
    return host.rstrip("/")


def resolve_output_path(output: str) -> Path:
    p = Path(output).expanduser()
    if not p.is_absolute():
        out_dir = Path.home() / "turbo-audit"
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir / p.name
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def vprint(args: argparse.Namespace, msg: str) -> None:
    if args.verbose:
        print(msg)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def create_session(args: argparse.Namespace) -> requests.Session:
    s = requests.Session()
    s.verify = False
    s.trust_env = bool(args.trust_env_proxy)
    return s


def login(session: requests.Session, host: str, args: argparse.Namespace) -> dict[str, Any]:
    print(f"\nConnecting to {host} ...")
    r = session.post(
        f"{host}/api/v3/login",
        data={"username": args.user, "password": args.password},
        timeout=60,
    )
    if r.status_code != 200:
        print(f"Authentication failed: HTTP {r.status_code}")
        print(short_text(r.text))
        raise SystemExit(1)
    print("Authentication successful")
    try:
        return r.json()
    except Exception:
        return {}


def unwrap_collection(data: Any) -> list[Any]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in (
            "results",
            "content",
            "items",
            "data",
            "entities",
            "policies",
            "settingsPolicies",
            "settingspolicies",
        ):
            value = data.get(key)
            if isinstance(value, list):
                return value
    return []


def get_first(obj: dict[str, Any], keys: list[str], default: Any = None) -> Any:
    for k in keys:
        if k in obj and obj.get(k) not in (None, ""):
            return obj.get(k)
    return default


def get_policy_id(policy: dict[str, Any]) -> str:
    return str(get_first(policy, ["uuid", "id", "policyId", "oid"], ""))


def get_policy_name(policy: dict[str, Any]) -> str:
    return str(get_first(policy, ["displayName", "name", "policyName", "uuid", "id"], "N/A"))


def get_entity_type(policy: dict[str, Any]) -> str:
    et = policy.get("entityType")
    if isinstance(et, dict):
        return str(get_first(et, ["value", "name", "displayName"], "UNKNOWN"))
    return str(et or "UNKNOWN")


def is_vmware_entity_type(entity_type: str) -> bool:
    return entity_type in VMWARE_ENTITY_TYPES


def get_enabled(policy: dict[str, Any]) -> bool:
    if "enabled" in policy:
        return bool(policy.get("enabled"))
    if "disabled" in policy:
        return not bool(policy.get("disabled"))
    return True


def get_scopes(policy: dict[str, Any]) -> list[Any]:
    scopes = policy.get("scopes", policy.get("scope", []))
    if isinstance(scopes, list):
        return scopes
    if scopes in (None, ""):
        return []
    return [scopes]


def get_scope_uuid(scope: Any) -> str:
    if isinstance(scope, str):
        return scope
    if isinstance(scope, dict):
        return str(get_first(scope, ["uuid", "id", "groupUuid", "oid"], ""))
    return ""


def get_scope_name(scope: Any) -> str:
    if isinstance(scope, dict):
        return str(get_first(scope, ["displayName", "name", "uuid", "id"], ""))
    return str(scope or "")


def is_default_policy(policy: dict[str, Any], default_ids: set[str]) -> bool:
    pid = get_policy_id(policy)
    if pid and pid in default_ids:
        return True
    for k in ("default", "isDefault", "defaultPolicy", "readOnly", "systemPolicy"):
        if bool(policy.get(k)):
            return True
    return False


def name_has_test_pattern(name: str) -> bool:
    lname = name.lower()
    return any(re.search(rf"(^|[^a-z0-9]){re.escape(p)}([^a-z0-9]|$)", lname) for p in TEST_NAME_PATTERNS)


def get_target_field(target: dict[str, Any], field_name: str) -> Any:
    for field in target.get("inputFields", []) or []:
        if isinstance(field, dict) and field.get("name") == field_name:
            return field.get("value")
    return None


def is_vmware_target(target: dict[str, Any]) -> bool:
    category = str(target.get("category", "")).lower()
    target_type = str(target.get("type", "")).lower()
    display = str(target.get("displayName", target.get("name", ""))).lower()
    address = str(
        get_target_field(target, "address")
        or get_target_field(target, "nameOrAddress")
        or get_target_field(target, "host")
        or ""
    ).lower()

    if category == "hypervisor" and target_type in {"vcenter", "vmware", "vsphere", "vmware vcenter"}:
        return True
    if "vcenter" in target_type or "vmware" in target_type or "vsphere" in target_type:
        return True
    if any(x in display for x in ("vcenter", "vcsa", "vsphere", "vmware")):
        return True
    if any(x in address for x in ("vcenter", "vcsa", "vsphere", "vmware")):
        return True
    return False


def status_is_accepted(status: str) -> bool:
    return str(status).strip().lower() in ACCEPTED_TARGET_STATES


def api_get_json(
    session: requests.Session,
    url: str,
    params: dict[str, Any] | None = None,
    timeout: int = 120,
) -> tuple[int, Any, str]:
    try:
        r = session.get(url, params=params, timeout=timeout)
        try:
            return r.status_code, r.json(), r.text
        except Exception:
            return r.status_code, None, r.text
    except Exception as e:
        return 0, None, str(e)


def analyze_vmware_targets(session: requests.Session, host: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    print("\n-- VMware vCenter Targets ---------------------------")
    status, data, text = api_get_json(session, f"{host}/api/v3/targets")
    if status != 200:
        print(f"  Unable to retrieve targets: HTTP {status}")
        print(f"  Response: {short_text(text)}")
        return [], []

    targets = unwrap_collection(data)
    vmware_targets = [t for t in targets if isinstance(t, dict) and is_vmware_target(t)]
    print(f"  Total VMware targets: {len(vmware_targets)}")

    not_accepted: list[dict[str, Any]] = []
    for target in vmware_targets:
        status_value = str(target.get("status", "UNKNOWN"))
        name = str(
            target.get("displayName")
            or get_target_field(target, "address")
            or get_target_field(target, "nameOrAddress")
            or target.get("uuid")
            or "N/A"
        )
        if not status_is_accepted(status_value):
            not_accepted.append(
                {"name": name, "status": status_value, "uuid": target.get("uuid", "N/A")}
            )
            print(f"  WARNING: {name}: {status_value}")

    if vmware_targets and not not_accepted:
        print("  All VMware/vCenter targets are in an accepted state")
    return vmware_targets, not_accepted

def count_entities_by_search(
    session: requests.Session,
    host: str,
    entity_type: str,
    args: argparse.Namespace,
    page_limit: int = 10000,
) -> tuple[int, str]:
    """
    Count entities returned by /api/v3/search for a given entity type.

    Important:
    /api/v3/search?types=X&limit=1 does not return the total inventory count.
    It only returns one matching element. Therefore, counting len(response)
    with limit=1 is incorrect.

    This function first tries a high limit and uses any explicit total/count
    field if the API returns one. If the response is a list, it counts the
    returned items. For environments larger than page_limit, increase
    page_limit or add cursor-based pagination after validating the local
    Swagger response shape.
    """
    status, data, text = api_get_json(
        session,
        f"{host}/api/v3/search",
        params={"types": entity_type, "limit": page_limit},
        timeout=180,
    )

    if status != 200:
        return 0, f"HTTP {status}: {short_text(text, 200)}"

    if isinstance(data, dict):
        for key in ("count", "total", "totalCount", "total_count"):
            if isinstance(data.get(key), int):
                return int(data[key]), f"HTTP 200 total field={key}"

        items = unwrap_collection(data)
        count = len(items)

        # Best-effort cursor visibility for troubleshooting
        cursor = (
            data.get("cursor")
            or data.get("nextCursor")
            or data.get("next_cursor")
            or data.get("next")
        )
        if cursor:
            return count, f"HTTP 200 partial page count={count}, cursor present"

        return count, f"HTTP 200 dict collection count={count}"

    if isinstance(data, list):
        count = len(data)
        if count >= page_limit:
            return count, f"HTTP 200 list count={count}, may be truncated by limit={page_limit}"
        return count, "HTTP 200 list"

    return 0, f"HTTP 200 unexpected response type={type(data).__name__}"


def get_vmware_entity_counts(session: requests.Session, host: str, args: argparse.Namespace) -> dict[str, int]:
    print("\n-- VMware Entity Inventory --------------------------")
    counts: dict[str, int] = {}

    for entity_type in VMWARE_ENTITY_TYPES:
        count, detail = count_entities_by_search(
            session=session,
            host=host,
            entity_type=entity_type,
            args=args,
            page_limit=20000,
        )
        counts[entity_type] = count
        vprint(args, f"  {entity_type}: {count} ({detail})")

    print(f"  Total VMware entities: {sum(counts.values())}")
    return counts


    print("\n-- VMware Entity Inventory --------------------------")
    counts: dict[str, int] = {}
    for entity_type in VMWARE_ENTITY_TYPES:
        status, data, _ = api_get_json(
            session,
            f"{host}/api/v3/search",
            params={"types": entity_type, "limit": 1},
            timeout=120,
        )
        count = 0
        if status == 200:
            if isinstance(data, dict):
                for k in ("count", "total", "totalCount"):
                    if isinstance(data.get(k), int):
                        count = int(data[k])
                        break
                else:
                    count = len(unwrap_collection(data))
            elif isinstance(data, list):
                count = len(data)
        counts[entity_type] = count
        vprint(args, f"  {entity_type}: {count} (HTTP {status})")
    print(f"  Total VMware entities: {sum(counts.values())}")
    return counts


def fetch_settings_policies(session: requests.Session, host: str, args: argparse.Namespace) -> tuple[list[dict[str, Any]], int, str]:
    print("\n-- Automation / Settings Policies -------------------")
    status, data, text = api_get_json(session, f"{host}/api/v3/settingspolicies")
    if status != 200:
        print(f"  Unable to retrieve settings policies: HTTP {status}")
        print(f"  Response: {short_text(text)}")
        return [], status, text
    policies = [p for p in unwrap_collection(data) if isinstance(p, dict)]
    vmware_count = sum(1 for p in policies if is_vmware_entity_type(get_entity_type(p)))
    print(f"  Total settings policies: {len(policies)}")
    print(f"  VMware-specific: {vmware_count}")
    return policies, status, text


def fetch_default_settings_policies(session: requests.Session, host: str) -> tuple[list[dict[str, Any]], int, str]:
    print("\n-- Default Automation Policies ----------------------")
    status, data, text = api_get_json(
        session,
        f"{host}/api/v3/settingspolicies",
        params={"only_defaults": "true"},
    )
    if status != 200:
        print(f"  Unable to retrieve default settings policies: HTTP {status}")
        print(f"  Response: {short_text(text)}")
        return [], status, text
    defaults = [p for p in unwrap_collection(data) if isinstance(p, dict)]
    print(f"  Default settings policies: {len(defaults)}")
    return defaults, status, text


def fetch_policy_endpoint(session: requests.Session, host: str) -> tuple[list[dict[str, Any]], int, str]:
    print("\n-- Placement / Policy Endpoint ----------------------")
    status, data, text = api_get_json(session, f"{host}/api/v3/policies")
    if status != 200:
        print(f"  Unable to retrieve /policies endpoint: HTTP {status}")
        print(f"  Response: {short_text(text)}")
        return [], status, text
    policies = [p for p in unwrap_collection(data) if isinstance(p, dict)]
    vmware_count = sum(1 for p in policies if is_vmware_entity_type(get_entity_type(p)))
    print(f"  Total policies: {len(policies)}")
    print(f"  VMware-specific by entityType: {vmware_count}")
    print(f"  Other entity types: {len(policies) - vmware_count}")
    return policies, status, text


def flatten_settings(obj: Any, path: str = "") -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if isinstance(obj, dict):
        has_value = "value" in obj
        has_identity = any(k in obj for k in ("uuid", "name", "displayName", "settingType", "type"))
        if has_value and has_identity:
            key = str(
                get_first(
                    obj,
                    ["uuid", "name", "displayName", "settingType", "type"],
                    path or "setting",
                )
            )
            rows.append(
                {
                    "key": key,
                    "path": path,
                    "name": str(get_first(obj, ["displayName", "name", "settingType", "type"], key)),
                    "value": obj.get("value"),
                    "defaultValue": obj.get("defaultValue"),
                    "raw": obj,
                }
            )
        for k, v in obj.items():
            if k in {"links", "href"}:
                continue
            next_path = f"{path}.{k}" if path else str(k)
            rows.extend(flatten_settings(v, next_path))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            next_path = f"{path}[{i}]" if path else f"[{i}]"
            rows.extend(flatten_settings(v, next_path))
    return rows


def build_default_setting_index(defaults: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    index: dict[str, dict[str, Any]] = defaultdict(dict)
    for p in defaults:
        etype = get_entity_type(p)
        for s in flatten_settings(p):
            index[etype][s["key"]] = s.get("value")
    return index


def count_setting_differences(policy: dict[str, Any], default_index: dict[str, dict[str, Any]]) -> tuple[int, int]:
    etype = get_entity_type(policy)
    defaults = default_index.get(etype, {})
    if not defaults:
        return 0, 0
    compared = 0
    diffs = 0
    for s in flatten_settings(policy):
        key = s["key"]
        if key in defaults:
            compared += 1
            if str(s.get("value")) != str(defaults.get(key)):
                diffs += 1
    return diffs, compared


def extract_action_modes(policy: dict[str, Any], default_index: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    etype = get_entity_type(policy)
    defaults = default_index.get(etype, {})
    modes: list[dict[str, Any]] = []
    for s in flatten_settings(policy):
        value = s.get("value")
        value_str = str(value).upper() if value is not None else ""
        if value_str in ACTION_MODE_VALUES:
            default_value = defaults.get(s["key"], s.get("defaultValue"))
            modes.append(
                {
                    "policy_id": get_policy_id(policy),
                    "policy_name": get_policy_name(policy),
                    "entityType": etype,
                    "setting_key": s["key"],
                    "setting_name": s["name"],
                    "path": s["path"],
                    "value": value_str,
                    "default_value": default_value,
                    "risk": action_mode_risk(value_str),
                }
            )
    return modes


def action_mode_risk(value: str) -> str:
    v = str(value).upper()
    if v == "AUTOMATIC":
        return "High impact - automatic execution"
    if v == "EXTERNAL_APPROVAL":
        return "Workflow/integration dependency"
    if v == "MANUAL":
        return "Controlled manual execution"
    if v == "RECOMMEND":
        return "Recommendation only"
    if v == "DISABLED":
        return "Disabled action"
    return "Unknown"


def get_group_member_count(group: dict[str, Any]) -> Any:
    for key in (
        "memberCount",
        "membersCount",
        "entityCount",
        "entitiesCount",
        "count",
        "totalCount",
    ):
        value = group.get(key)
        if isinstance(value, int):
            return value
    members = group.get("members") or group.get("entities")
    if isinstance(members, list):
        return len(members)
    return None


def get_group_info(
    session: requests.Session,
    host: str,
    group_uuid: str,
    cache: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    if not group_uuid:
        return {"uuid": "", "status": "no uuid", "exists": False, "member_count": None}
    if group_uuid in cache:
        return cache[group_uuid]

    status, data, text = api_get_json(session, f"{host}/api/v3/groups/{group_uuid}")
    if status == 200 and isinstance(data, dict):
        info = {
            "uuid": group_uuid,
            "status": "ok",
            "exists": True,
            "name": get_first(data, ["displayName", "name"], group_uuid),
            "groupType": get_first(data, ["groupType", "type", "entityType"], ""),
            "member_count": get_group_member_count(data),
        }
    else:
        info = {
            "uuid": group_uuid,
            "status": f"HTTP {status}",
            "exists": False,
            "name": "",
            "groupType": "",
            "member_count": None,
            "response": short_text(text, 300),
        }
    cache[group_uuid] = info
    return info


def analyze_scopes(
    session: requests.Session,
    host: str,
    policy: dict[str, Any],
    source: str,
    skip_group_check: bool,
    group_cache: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], int, int, int]:
    scopes = get_scopes(policy)
    rows: list[dict[str, Any]] = []
    unresolved = 0
    empty_groups = 0
    checked = 0
    for scope in scopes:
        uuid = get_scope_uuid(scope)
        name = get_scope_name(scope)
        info = {
            "uuid": uuid,
            "name": name,
            "exists": None,
            "member_count": None,
            "groupType": "",
            "status": "not checked",
        }
        if not skip_group_check and uuid:
            checked += 1
            info = get_group_info(session, host, uuid, group_cache)
            if not info.get("exists"):
                unresolved += 1
            elif info.get("member_count") == 0:
                empty_groups += 1
        rows.append(
            {
                "policy_id": get_policy_id(policy),
                "policy_name": get_policy_name(policy),
                "source": source,
                "entityType": get_entity_type(policy),
                "scope_uuid": uuid,
                "scope_name": name or info.get("name", ""),
                "exists": info.get("exists"),
                "member_count": info.get("member_count"),
                "groupType": info.get("groupType", ""),
                "status": info.get("status", ""),
            }
        )
    return rows, checked, unresolved, empty_groups


def check_recent_actions(
    session: requests.Session,
    host: str,
    policy_id: str,
    days: int,
    cache: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    if not policy_id:
        return {"available": False, "count": None, "status": "no policy id"}
    if policy_id in cache:
        return cache[policy_id]

    cutoff_ms = int((now_utc().timestamp() - days * 86400) * 1000)
    status, data, text = api_get_json(
        session,
        f"{host}/api/v3/actions",
        params={"policyId": policy_id, "starttime": cutoff_ms, "limit": 1},
        timeout=60,
    )
    if status == 200:
        items = unwrap_collection(data)
        result = {"available": True, "count": len(items), "status": "ok"}
    else:
        result = {
            "available": False,
            "count": None,
            "status": f"HTTP {status}",
            "response": short_text(text, 300),
        }
    cache[policy_id] = result
    return result


def safe_decode(raw: bytes) -> str:
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def content_disposition_filename(headers: dict[str, str]) -> str | None:
    cd = headers.get("content-disposition") or headers.get("Content-Disposition")
    if not cd:
        return None
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', cd)
    if m:
        return Path(m.group(1)).name
    return None


def fetch_admin_auditlogs(
    session: requests.Session,
    host: str,
    days: int,
    artifact_dir: Path,
) -> dict[str, Any]:
    print("\n-- Admin Audit Logs ---------------------------------")
    artifact_dir.mkdir(parents=True, exist_ok=True)
    result: dict[str, Any] = {
        "requested": True,
        "available": False,
        "status_code": None,
        "message": "not attempted",
        "path": "",
        "sha256": "",
        "files": [],
    }
    try:
        r = session.get(
            f"{host}/api/v3/admin/auditlogs",
            params={"days": days},
            headers={"Accept": "application/gzip"},
            timeout=300,
        )
    except Exception as e:
        result["message"] = str(e)
        print(f"  Unable to download admin audit logs: {e}")
        return result

    result["status_code"] = r.status_code
    if r.status_code != 200:
        result["message"] = f"HTTP {r.status_code}: {short_text(r.text, 500)}"
        if r.status_code == 403:
            print("  Admin audit logs unavailable: HTTP 403 Access denied")
            print("  A user with Administrator or Site Administrator privileges is normally required.")
        else:
            print(f"  Admin audit logs unavailable: HTTP {r.status_code}")
            print(f"  Response: {short_text(r.text, 500)}")
        return result

    filename = content_disposition_filename(r.headers) or f"turbo_auditlogs_{days}d_{now_stamp()}.tar.gz"
    path = artifact_dir / filename
    path.write_bytes(r.content)
    result["available"] = True
    result["message"] = "downloaded"
    result["path"] = str(path)
    result["sha256"] = sha256_file(path)
    print(f"  Admin audit logs downloaded: {path}")
    print(f"  SHA256: {result['sha256']}")

    result["files"] = extract_audit_text_files(path)
    print(f"  Parsed audit text files: {len(result['files'])}")
    return result


def extract_audit_text_files(path: Path, max_file_bytes: int = 20 * 1024 * 1024, max_total_bytes: int = 250 * 1024 * 1024) -> list[dict[str, str]]:
    files: list[dict[str, str]] = []
    total = 0

    def add_file(name: str, raw: bytes) -> None:
        nonlocal total
        if total >= max_total_bytes:
            return
        if len(raw) > max_file_bytes:
            raw = raw[:max_file_bytes]
        total += len(raw)
        text = safe_decode(raw)
        files.append({"name": name, "text": text})

    # tar.gz or tar
    try:
        with tarfile.open(path, "r:*") as tf:
            for member in tf.getmembers():
                if not member.isfile():
                    continue
                fh = tf.extractfile(member)
                if not fh:
                    continue
                add_file(member.name, fh.read(max_file_bytes + 1))
                if total >= max_total_bytes:
                    break
            return files
    except tarfile.TarError:
        pass

    # zip
    try:
        if zipfile.is_zipfile(path):
            with zipfile.ZipFile(path) as zf:
                for name in zf.namelist():
                    if name.endswith("/"):
                        continue
                    with zf.open(name) as fh:
                        add_file(name, fh.read(max_file_bytes + 1))
                    if total >= max_total_bytes:
                        break
            return files
    except Exception:
        pass

    # single gzip
    try:
        with gzip.open(path, "rb") as gf:
            add_file(path.name.replace(".gz", ""), gf.read(max_file_bytes + 1))
            return files
    except Exception:
        pass

    # raw text/binary fallback
    try:
        add_file(path.name, path.read_bytes()[: max_file_bytes + 1])
    except Exception:
        pass
    return files


def build_audit_matches(
    policies: list[dict[str, Any]],
    audit_result: dict[str, Any],
    max_matches_per_policy: int,
) -> dict[str, list[dict[str, str]]]:
    matches: dict[str, list[dict[str, str]]] = defaultdict(list)
    if not audit_result.get("available"):
        return matches

    token_map: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for p in policies:
        pid = get_policy_id(p)
        name = get_policy_name(p)
        if pid:
            token_map[pid.lower()].append((pid, "uuid"))
        clean_name = name.strip()
        if len(clean_name) >= 6 and clean_name.lower() not in GENERIC_NAMES_TO_IGNORE_FOR_AUDIT:
            token_map[clean_name.lower()].append((pid, "name"))

    if not token_map:
        return matches

    for file_item in audit_result.get("files", []):
        fname = file_item.get("name", "")
        text = file_item.get("text", "")
        for line_no, line in enumerate(text.splitlines(), 1):
            lower = line.lower()
            for token, pid_kind_list in token_map.items():
                if token and token in lower:
                    for pid, kind in pid_kind_list:
                        if len(matches[pid]) < max_matches_per_policy:
                            matches[pid].append(
                                {
                                    "file": fname,
                                    "line": str(line_no),
                                    "match_type": kind,
                                    "text": short_text(line.strip(), 500),
                                }
                            )
                    break
    return matches


def classify_settings_policy(
    policy: dict[str, Any],
    default_ids: set[str],
    scope_count: int,
    unresolved_scopes: int,
    empty_groups: int,
    setting_diff_count: int,
    setting_compare_count: int,
    action_modes: list[dict[str, Any]],
    audit_match_count: int,
    actions_info: dict[str, Any] | None,
) -> tuple[str, str, list[str]]:
    name = get_policy_name(policy)
    etype = get_entity_type(policy)
    enabled = get_enabled(policy)
    default_policy = is_default_policy(policy, default_ids)
    vmware = is_vmware_entity_type(etype)
    reasons: list[str] = []

    if default_policy:
        reasons.append("default/system policy")
    if not enabled:
        reasons.append("disabled")
    if scope_count == 0 and not default_policy:
        reasons.append("empty scope")
    if unresolved_scopes:
        reasons.append(f"unresolved scopes: {unresolved_scopes}")
    if empty_groups:
        reasons.append(f"empty groups: {empty_groups}")
    if setting_compare_count > 0 and setting_diff_count == 0 and not default_policy:
        reasons.append("no differences from default")
    if name_has_test_pattern(name):
        reasons.append("test/demo/temp naming pattern")
    if action_modes:
        automatic = sum(1 for a in action_modes if a.get("value") == "AUTOMATIC")
        disabled_modes = sum(1 for a in action_modes if a.get("value") == "DISABLED")
        if automatic:
            reasons.append(f"automatic action modes: {automatic}")
        if disabled_modes:
            reasons.append(f"disabled action modes: {disabled_modes}")
    if audit_match_count:
        reasons.append(f"audit log matches: {audit_match_count}")
    if actions_info and actions_info.get("available") and actions_info.get("count"):
        reasons.append("recent action evidence")

    if default_policy:
        return "KEEP", "Default/system policy; do not delete from stale cleanup.", reasons

    # Strong delete candidate only when several safe/stale signals are combined.
    if (
        not enabled
        and (scope_count == 0 or unresolved_scopes > 0 or empty_groups > 0)
        and (name_has_test_pattern(name) or (setting_compare_count > 0 and setting_diff_count == 0))
        and audit_match_count == 0
    ):
        return (
            "CANDIDATE_DELETE",
            "Disabled non-default policy with empty/unresolved/empty scope and additional stale signal. Validate manually before deletion.",
            reasons,
        )

    if not vmware:
        return "INFO", "Non-VMware entity type included for visibility; not part of VMware cleanup decision.", reasons

    if not enabled or scope_count == 0 or unresolved_scopes or empty_groups:
        return "REVIEW", "Policy has scope/enabled signals that require manual validation.", reasons

    if setting_compare_count > 0 and setting_diff_count == 0:
        return "REVIEW", "Custom policy appears equivalent to default; validate whether it is redundant.", reasons

    if action_modes:
        return "REVIEW", "Policy contains action mode settings; review operational impact.", reasons

    return "KEEP", "No strong stale indicators detected.", reasons


def classify_placement_policy(policy: dict[str, Any], audit_match_count: int) -> tuple[str, str, list[str]]:
    name = get_policy_name(policy)
    etype = get_entity_type(policy)
    enabled = get_enabled(policy)
    reasons: list[str] = []
    if not enabled:
        reasons.append("disabled")
    if name_has_test_pattern(name):
        reasons.append("test/demo/temp naming pattern")
    if audit_match_count:
        reasons.append(f"audit log matches: {audit_match_count}")

    if is_vmware_entity_type(etype):
        if not enabled or name_has_test_pattern(name):
            return "REVIEW", "VMware-related placement/policy endpoint object requires manual review.", reasons
        return "INFO", "VMware-related placement/policy endpoint object; analyze separately from automation settings.", reasons

    return "INFO", "Placement/policy endpoint object with non-VMware or unknown entity type; included for inventory only.", reasons


def make_policy_row(
    policy: dict[str, Any],
    source: str,
    classification: str,
    justification: str,
    reasons: list[str],
    default_ids: set[str],
    setting_diff_count: int = 0,
    setting_compare_count: int = 0,
    action_mode_count: int = 0,
    scope_count: int | None = None,
    unresolved_scopes: int = 0,
    empty_groups: int = 0,
    audit_match_count: int = 0,
    actions_info: dict[str, Any] | None = None,
) -> dict[str, Any]:
    etype = get_entity_type(policy)
    scopes = get_scopes(policy)
    if scope_count is None:
        scope_count = len(scopes)
    return {
        "classification": classification,
        "source": source,
        "vmware_specific": str(is_vmware_entity_type(etype)),
        "name": get_policy_name(policy),
        "entityType": etype,
        "enabled": str(get_enabled(policy)),
        "default_policy": str(is_default_policy(policy, default_ids)),
        "scope_count": scope_count,
        "unresolved_scopes": unresolved_scopes,
        "empty_groups": empty_groups,
        "settings_diff_count": setting_diff_count,
        "settings_compared": setting_compare_count,
        "action_mode_count": action_mode_count,
        "audit_matches": audit_match_count,
        "action_check": actions_info.get("status") if actions_info else "not checked",
        "reasons": "; ".join(reasons),
        "justification": justification,
        "last_modified": get_first(policy, ["lastModified", "modifiedTime", "updatedAt"], "N/A"),
        "id": get_policy_id(policy),
    }


def analyze_conflicts(settings_policies: list[dict[str, Any]]) -> list[dict[str, Any]]:
    index: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for p in settings_policies:
        etype = get_entity_type(p)
        scopes = get_scopes(p)
        scope_ids = sorted([get_scope_uuid(s) or get_scope_name(s) for s in scopes])
        scope_key = ",".join(scope_ids) if scope_ids else "<empty-scope>"
        for setting in flatten_settings(p):
            key = (etype, scope_key, setting["key"])
            index[key].append(
                {
                    "policy_id": get_policy_id(p),
                    "policy_name": get_policy_name(p),
                    "entityType": etype,
                    "scope_key": scope_key,
                    "setting_key": setting["key"],
                    "setting_name": setting["name"],
                    "value": setting.get("value"),
                }
            )

    conflicts: list[dict[str, Any]] = []
    for (_etype, _scope, _setting), rows in index.items():
        values = {json.dumps(r.get("value"), sort_keys=True, default=str) for r in rows}
        policy_ids = {r.get("policy_id") for r in rows}
        if len(policy_ids) > 1 and len(values) > 1:
            for r in rows:
                r2 = dict(r)
                r2["conflict_count"] = len(rows)
                r2["distinct_values"] = len(values)
                conflicts.append(r2)
    return conflicts


def save_snapshot(
    snapshot_dir: Path,
    metadata: dict[str, Any],
    policy_rows: list[dict[str, Any]],
) -> tuple[Path, list[dict[str, Any]]]:
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    current = {
        "metadata": metadata,
        "policies": {r.get("id") or f"row-{i}": r for i, r in enumerate(policy_rows)},
    }

    previous_files = sorted(snapshot_dir.glob("policy_snapshot_*.json"))
    previous_file = previous_files[-1] if previous_files else None
    changes: list[dict[str, Any]] = []
    if previous_file:
        try:
            previous = json.loads(previous_file.read_text(encoding="utf-8"))
            prev_policies = previous.get("policies", {})
            curr_policies = current.get("policies", {})
            prev_ids = set(prev_policies)
            curr_ids = set(curr_policies)
            for pid in sorted(curr_ids - prev_ids):
                changes.append({"change": "ADDED", "id": pid, "name": curr_policies[pid].get("name", ""), "details": "New policy row in current snapshot"})
            for pid in sorted(prev_ids - curr_ids):
                changes.append({"change": "REMOVED", "id": pid, "name": prev_policies[pid].get("name", ""), "details": "Policy row no longer present"})
            watch_fields = ["classification", "enabled", "scope_count", "reasons", "settings_diff_count", "action_mode_count"]
            for pid in sorted(curr_ids & prev_ids):
                diffs = []
                for f in watch_fields:
                    if str(prev_policies[pid].get(f)) != str(curr_policies[pid].get(f)):
                        diffs.append(f"{f}: {prev_policies[pid].get(f)} -> {curr_policies[pid].get(f)}")
                if diffs:
                    changes.append({"change": "CHANGED", "id": pid, "name": curr_policies[pid].get("name", ""), "details": "; ".join(diffs)})
        except Exception as e:
            changes.append({"change": "SNAPSHOT_ERROR", "id": "", "name": str(previous_file), "details": str(e)})

    out = snapshot_dir / f"policy_snapshot_{now_stamp()}.json"
    out.write_text(json.dumps(current, indent=2, sort_keys=True, default=str), encoding="utf-8")
    return out, changes


def append_rows(ws, rows: list[dict[str, Any]], headers: list[str]) -> None:
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])
        classification = row.get("classification")
        fill_color = CLASS_COLORS.get(str(classification), "FFFFFF")
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=ws.max_row, column=c)
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    for i, h in enumerate(headers, 1):
        width = min(max(len(h) + 2, 12), 60)
        if h in {"justification", "reasons", "details", "text"}:
            width = 70
        if h in {"id", "policy_id", "scope_uuid"}:
            width = 38
        ws.column_dimensions[get_column_letter(i)].width = width


def generate_excel_report(
    output_path: Path,
    metadata: dict[str, Any],
    policy_rows: list[dict[str, Any]],
    action_mode_rows: list[dict[str, Any]],
    scope_rows: list[dict[str, Any]],
    conflict_rows: list[dict[str, Any]],
    audit_match_rows: list[dict[str, Any]],
    snapshot_changes: list[dict[str, Any]],
    vmware_targets: list[dict[str, Any]],
    not_accepted_targets: list[dict[str, Any]],
    entity_counts: dict[str, int],
) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    summary_rows = [
        ["Turbonomic Stale Policy Audit - VMware On-Premises"],
        [],
        ["Host", metadata.get("host")],
        ["Analysis Date", metadata.get("analysis_date")],
        ["Inactivity Threshold", f"{metadata.get('days')} days"],
        ["Audit Requested", metadata.get("audit_requested")],
        ["Audit Available", metadata.get("audit_available")],
        ["Audit Message", metadata.get("audit_message")],
        ["Audit Artifact", metadata.get("audit_artifact")],
        ["Snapshot", metadata.get("snapshot_path")],
        [],
        ["VMware Environment"],
        ["VMware Targets", len(vmware_targets)],
        ["Targets Not Accepted", len(not_accepted_targets)],
        ["Total VMware Entities", sum(entity_counts.values())],
    ]
    for etype, count in entity_counts.items():
        summary_rows.append([f"  {etype}", count])
    summary_rows.extend([
        [],
        ["Policy Audit Results"],
        ["TOTAL rows", len(policy_rows)],
        ["VMware-specific", sum(1 for r in policy_rows if r.get("vmware_specific") == "True")],
        ["Other types", sum(1 for r in policy_rows if r.get("vmware_specific") != "True")],
    ])
    for cls in CLASS_ORDER:
        summary_rows.append([cls, sum(1 for r in policy_rows if r.get("classification") == cls)])
    summary_rows.extend([
        [],
        ["Additional Analysis"],
        ["Action mode rows", len(action_mode_rows)],
        ["Scope analysis rows", len(scope_rows)],
        ["Conflict rows", len(conflict_rows)],
        ["Audit log match rows", len(audit_match_rows)],
        ["Snapshot change rows", len(snapshot_changes)],
    ])
    for row in summary_rows:
        ws.append(row)
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 100

    policy_headers = [
        "classification", "source", "vmware_specific", "name", "entityType", "enabled",
        "default_policy", "scope_count", "unresolved_scopes", "empty_groups",
        "settings_diff_count", "settings_compared", "action_mode_count", "audit_matches",
        "action_check", "reasons", "justification", "last_modified", "id",
    ]
    for cls in CLASS_ORDER:
        ws_cls = wb.create_sheet(cls[:31])
        append_rows(ws_cls, [r for r in policy_rows if r.get("classification") == cls], policy_headers)

    ws_all = wb.create_sheet("All Policies")
    sorted_rows = sorted(policy_rows, key=lambda r: (CLASS_ORDER.index(r.get("classification", "UNKNOWN")) if r.get("classification") in CLASS_ORDER else 99, r.get("source", ""), r.get("name", "")))
    append_rows(ws_all, sorted_rows, policy_headers)

    ws_modes = wb.create_sheet("Action Modes")
    append_rows(ws_modes, action_mode_rows, ["policy_name", "entityType", "setting_name", "setting_key", "value", "default_value", "risk", "path", "policy_id"])

    ws_scope = wb.create_sheet("Scope Analysis")
    append_rows(ws_scope, scope_rows, ["policy_name", "source", "entityType", "scope_name", "scope_uuid", "exists", "member_count", "groupType", "status", "policy_id"])

    ws_conflicts = wb.create_sheet("Policy Conflicts")
    append_rows(ws_conflicts, conflict_rows, ["policy_name", "entityType", "scope_key", "setting_name", "setting_key", "value", "conflict_count", "distinct_values", "policy_id"])

    ws_audit = wb.create_sheet("Audit Log Matches")
    append_rows(ws_audit, audit_match_rows, ["policy_name", "source", "match_type", "file", "line", "text", "policy_id"])

    ws_changes = wb.create_sheet("Snapshot Changes")
    append_rows(ws_changes, snapshot_changes, ["change", "name", "details", "id"])

    ws_vmware = wb.create_sheet("VMware Environment")
    ws_vmware.append(["VMware vCenter Targets"])
    ws_vmware.append(["Target Name", "Type", "Category", "Status", "UUID"])
    for target in vmware_targets:
        ws_vmware.append([
            target.get("displayName") or get_target_field(target, "address") or "N/A",
            target.get("type", ""),
            target.get("category", ""),
            target.get("status", ""),
            target.get("uuid", ""),
        ])
    ws_vmware.append([])
    ws_vmware.append(["Entity Type", "Count"])
    for etype, count in entity_counts.items():
        ws_vmware.append([etype, count])
    for col in range(1, 6):
        ws_vmware.column_dimensions[get_column_letter(col)].width = 32

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def copy_to_parent_if_possible(output_path: Path) -> Path | None:
    try:
        parent_dest = output_path.parent.parent / output_path.name
        if parent_dest != output_path:
            shutil.copy2(output_path, parent_dest)
            return parent_dest
    except Exception:
        return None
    return None


def main() -> None:
    args = parse_args()
    host = normalize_url(args.host)
    output_path = resolve_output_path(args.output)
    output_dir = output_path.parent
    audit_days = args.admin_audit_days if args.admin_audit_days is not None else args.days
    audit_artifact_dir = Path(args.audit_artifact_dir).expanduser() if args.audit_artifact_dir else output_dir / "auditlogs"
    snapshot_dir = Path(args.snapshot_dir).expanduser() if args.snapshot_dir else output_dir / "snapshots"

    print("=" * 84)
    print("  Turbonomic Stale Policy Audit - VMware On-Premises Focus")
    print("=" * 84)

    session = create_session(args)
    login_info = login(session, host, args)

    vmware_targets, not_accepted_targets = analyze_vmware_targets(session, host)
    entity_counts = get_vmware_entity_counts(session, host, args)

    settings_policies, settings_status, _ = fetch_settings_policies(session, host, args)
    default_policies, default_status, _ = fetch_default_settings_policies(session, host)
    placement_policies, placement_status, _ = fetch_policy_endpoint(session, host)

    default_ids = {get_policy_id(p) for p in default_policies if get_policy_id(p)}
    default_index = build_default_setting_index(default_policies)

    all_raw_policies = settings_policies + placement_policies

    audit_result = {
        "requested": bool(args.use_admin_auditlogs),
        "available": False,
        "status_code": None,
        "message": "not requested",
        "path": "",
        "sha256": "",
        "files": [],
    }
    if args.use_admin_auditlogs:
        audit_result = fetch_admin_auditlogs(session, host, audit_days, audit_artifact_dir)

    audit_matches = build_audit_matches(all_raw_policies, audit_result, args.max_audit_matches_per_policy)

    group_cache: dict[str, dict[str, Any]] = {}
    action_cache: dict[str, dict[str, Any]] = {}
    policy_rows: list[dict[str, Any]] = []
    action_mode_rows: list[dict[str, Any]] = []
    scope_rows: list[dict[str, Any]] = []
    audit_match_rows: list[dict[str, Any]] = []

    # Settings / automation policies
    for p in settings_policies:
        pid = get_policy_id(p)
        scopes = get_scopes(p)
        policy_scope_rows, _scope_checked, unresolved, empty_groups = analyze_scopes(
            session,
            host,
            p,
            "settings_policy",
            args.skip_group_check,
            group_cache,
        )
        scope_rows.extend(policy_scope_rows)
        setting_diff_count, setting_compare_count = count_setting_differences(p, default_index)
        modes = extract_action_modes(p, default_index)
        action_mode_rows.extend(modes)
        actions_info = None
        if not args.skip_action_check:
            actions_info = check_recent_actions(session, host, pid, args.days, action_cache)
        audit_count = len(audit_matches.get(pid, []))
        classification, justification, reasons = classify_settings_policy(
            p,
            default_ids,
            len(scopes),
            unresolved,
            empty_groups,
            setting_diff_count,
            setting_compare_count,
            modes,
            audit_count,
            actions_info,
        )
        policy_rows.append(
            make_policy_row(
                p,
                "settings_policy",
                classification,
                justification,
                reasons,
                default_ids,
                setting_diff_count=setting_diff_count,
                setting_compare_count=setting_compare_count,
                action_mode_count=len(modes),
                scope_count=len(scopes),
                unresolved_scopes=unresolved,
                empty_groups=empty_groups,
                audit_match_count=audit_count,
                actions_info=actions_info,
            )
        )

    # Placement / policy endpoint rows are inventory-oriented and intentionally analyzed separately.
    for p in placement_policies:
        pid = get_policy_id(p)
        audit_count = len(audit_matches.get(pid, []))
        classification, justification, reasons = classify_placement_policy(p, audit_count)
        policy_rows.append(
            make_policy_row(
                p,
                "placement_policy_endpoint",
                classification,
                justification,
                reasons,
                default_ids,
                audit_match_count=audit_count,
                actions_info=None,
            )
        )

    # Audit matches worksheet rows
    id_to_name_source = {r["id"]: (r["name"], r["source"]) for r in policy_rows if r.get("id")}
    for pid, matches in audit_matches.items():
        name, source = id_to_name_source.get(pid, ("", ""))
        for m in matches:
            audit_match_rows.append(
                {
                    "policy_id": pid,
                    "policy_name": name,
                    "source": source,
                    "match_type": m.get("match_type", ""),
                    "file": m.get("file", ""),
                    "line": m.get("line", ""),
                    "text": m.get("text", ""),
                }
            )

    conflict_rows = analyze_conflicts(settings_policies)

    metadata = {
        "host": host,
        "analysis_date": now_utc().strftime("%Y-%m-%d %H:%M UTC"),
        "days": args.days,
        "user": login_info.get("username", args.user),
        "roles": ",".join([r.get("name", "") for r in login_info.get("roles", [])]) if isinstance(login_info.get("roles"), list) else "",
        "settings_status": settings_status,
        "default_settings_status": default_status,
        "placement_status": placement_status,
        "audit_requested": str(audit_result.get("requested")),
        "audit_available": str(audit_result.get("available")),
        "audit_status_code": audit_result.get("status_code"),
        "audit_message": audit_result.get("message"),
        "audit_artifact": audit_result.get("path"),
        "snapshot_path": "",
    }

    snapshot_changes: list[dict[str, Any]] = []
    if not args.no_snapshot:
        snapshot_path, snapshot_changes = save_snapshot(snapshot_dir, metadata, policy_rows)
        metadata["snapshot_path"] = str(snapshot_path)
        print(f"\nSnapshot saved: {snapshot_path}")

    print("\n" + "=" * 84)
    print("  AUDIT SUMMARY")
    print("=" * 84)
    print(f"  Host: {host}")
    print(f"  Inactivity threshold: {args.days} days")
    print(f"  Analysis date: {metadata['analysis_date']}")
    print(f"  User/roles: {metadata.get('user')} / {metadata.get('roles')}")
    if args.use_admin_auditlogs:
        print(f"  Admin auditlogs: {metadata['audit_available']} ({metadata['audit_message']})")
    else:
        print("  Admin auditlogs: not requested")
    print("=" * 84)
    print(f"  VMware targets    : {len(vmware_targets)} ({len(not_accepted_targets)} not accepted)")
    print(f"  VMware entities   : {sum(entity_counts.values())}")
    print("=" * 84)
    print(f"  TOTAL rows        : {len(policy_rows)}")
    print(f"    VMware-specific : {sum(1 for r in policy_rows if r.get('vmware_specific') == 'True')}")
    print(f"    Other types     : {sum(1 for r in policy_rows if r.get('vmware_specific') != 'True')}")
    print("=" * 84)
    for cls in CLASS_ORDER:
        print(f"  {cls:<16}: {sum(1 for r in policy_rows if r.get('classification') == cls)}")
    print("=" * 84)
    print(f"  Action mode rows  : {len(action_mode_rows)}")
    print(f"  Scope rows        : {len(scope_rows)}")
    print(f"  Conflict rows     : {len(conflict_rows)}")
    print(f"  Audit matches     : {len(audit_match_rows)}")
    print(f"  Snapshot changes  : {len(snapshot_changes)}")
    print("=" * 84)

    final_path = generate_excel_report(
        output_path,
        metadata,
        policy_rows,
        action_mode_rows,
        scope_rows,
        conflict_rows,
        audit_match_rows,
        snapshot_changes,
        vmware_targets,
        not_accepted_targets,
        entity_counts,
    )
    print(f"\nReport exported: {final_path}")
    copied = copy_to_parent_if_possible(final_path)
    if copied:
        print(f"Copied to: {copied}")
    print(
        "Sheets: 'Summary' | 'CANDIDATE_DELETE' | 'REVIEW' | 'KEEP' | 'INFO' | 'UNKNOWN' | "
        "'All Policies' | 'Action Modes' | 'Scope Analysis' | 'Policy Conflicts' | "
        "'Audit Log Matches' | 'Snapshot Changes' | 'VMware Environment'"
    )


if __name__ == "__main__":
    main()
