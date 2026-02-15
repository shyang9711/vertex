"""
Data IO functions for loading, saving, and importing/exporting client data.
"""

from __future__ import annotations

import json
import os
import sys
import copy
from pathlib import Path
from typing import List, Dict, Any
from datetime import datetime

# Import path handling
try:
    from vertex.utils.app_logging import get_logger
    from vertex.utils.helpers import (
        ensure_relation_dict,
        ensure_relation_link,
        normalize_logs,
        normalize_ein_digits,
        normalize_ssn_digits,
        normalize_phone_digits,
        tokenize,
        _account_manager_key,
        _account_manager_id_from_key,
        remove_stale_back_links,
    )
except ModuleNotFoundError:
    from utils.app_logging import get_logger
    from utils.helpers import (
        ensure_relation_dict,
        ensure_relation_link,
        normalize_logs,
        normalize_ein_digits,
        normalize_ssn_digits,
        normalize_phone_digits,
        tokenize,
        _account_manager_key,
        _account_manager_id_from_key,
        remove_stale_back_links,
    )

try:
    from tkinter import messagebox
except ImportError:
    messagebox = None

LOG = get_logger("io")


# Data directory setup (same as client_manager.py)
def portable_root() -> Path:
    """For EXE: folder that contains vertex.exe (NOT _MEI). For source run: script directory."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    try:
        return Path(__file__).resolve().parent.parent
    except NameError:
        return Path(os.getcwd()).resolve()


def appdata_root() -> Path:
    """Get AppData directory for Vertex."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "Vertex"


# Prefer a sibling "data" folder next to the EXE (portable install),
# but if it doesn't exist and can't be created, fall back to AppData.
_PORTABLE = portable_root() / "data"
try:
    _PORTABLE.mkdir(parents=True, exist_ok=True)
    DATA_ROOT = _PORTABLE
except Exception:
    DATA_ROOT = appdata_root() / "data"
    DATA_ROOT.mkdir(parents=True, exist_ok=True)

CLIENTS_DIR      = DATA_ROOT / "clients"
TASKS_DIR        = DATA_ROOT / "tasks"
MATCH_RULES_DIR  = DATA_ROOT / "match_rules"
MONTHLY_DATA_DIR = DATA_ROOT / "monthly_data"
VENDOR_LISTS_DIR = DATA_ROOT / "vendor_lists"

# Log folder at same level as data folder (e.g. .../log next to .../data)
LOG_DIR = DATA_ROOT.parent / "log"
LOG_DIR.mkdir(parents=True, exist_ok=True)

for _p in (CLIENTS_DIR, TASKS_DIR, MATCH_RULES_DIR, MONTHLY_DATA_DIR, VENDOR_LISTS_DIR):
    _p.mkdir(parents=True, exist_ok=True)

DATA_DIR  = CLIENTS_DIR
DATA_FILE = CLIENTS_DIR / "clients.json"

ACCOUNT_MANAGERS_FILE = CLIENTS_DIR / "account_managers.json"
TASKS_FILE            = TASKS_DIR / "tasks.json"
MONTHLY_STATE_FILE    = MONTHLY_DATA_DIR / "monthly_state.json"
MIGRATIONS_FILE       = DATA_ROOT / "migrations.json"

# Export _read_json_file and _write_json_file for use by other modules if needed
# (They're used internally but may be needed elsewhere)


def _read_json_file(path: Path, default):
    """Read JSON file, return default on error."""
    try:
        if not path.exists():
            return default
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def _write_json_file(path: Path, obj) -> None:
    """Write JSON file, creating parent directories if needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8")


def load_clients() -> List[Dict[str, Any]]:
    """Load clients from clients.json file."""
    if not DATA_FILE.exists():
        DATA_FILE.write_text("[]", encoding="utf-8")
        return []
    try:
        data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        out = []
        for c in data:
            c = dict(c) if isinstance(c, dict) else {}

            # Core
            c.setdefault("name","")
            c.setdefault("dba","")
            c.setdefault("ein","")
            c.setdefault("ssn","")
            c.setdefault("file_location","")
            c.setdefault("memo","")

            # client extended fields
            c.setdefault("addr1","")
            c.setdefault("addr2","")
            c.setdefault("city","")
            c.setdefault("state","")
            c.setdefault("zip","")
            c.setdefault("acct_mgr","")
            c.setdefault("edd_number","")
            c.setdefault("sales_tax_account","")
            c.setdefault("entity_type","")

            # Tax rates
            c.setdefault("ui_rate","")
            c.setdefault("sales_tax_rate","")
            c.setdefault("other_tax_rates","")
            c.setdefault("tax_rates_last_checked","")

            # Personnel - migrate officers to relations if present
            relations = c.get("relations")
            if relations is None:
                legacy = c.get("owner","")
                relations = [legacy] if legacy else []
            norm_offs: List[Dict[str,str]] = []
            if isinstance(relations, list):
                for x in relations:
                    norm_offs.append(ensure_relation_dict(x))
            elif isinstance(relations, dict):
                norm_offs.append(ensure_relation_dict(relations))
            else:
                norm_offs = [ensure_relation_dict(relations)] if relations else []
            
            # Migrate any remaining officers to relations
            officers = c.get("officers", []) or []
            if officers:
                for o in officers:
                    if isinstance(o, dict):
                        norm_offs.append(ensure_relation_dict(o))
                # Remove officers after migration
                c.pop("officers", None)
            
            # Normalize relations to ensure they use id format
            normalized_relations = []
            for rel in norm_offs:
                # Try to parse as relation link (for entity links)
                rel_link = ensure_relation_link(rel)
                if rel_link.get("id"):
                    # Entity link - use link format
                    normalized_relations.append(rel_link)
                else:
                    # Person relation - use dict format
                    normalized_relations.append(ensure_relation_dict(rel))
            
            c["relations"] = normalized_relations
            c.pop("owner", None)
            # Remove post_save_links if present (should not be in saved data)
            c.pop("post_save_links", None)

            # Optional employees list (if present, normalize same way)
            emps = c.get("employees", [])
            if isinstance(emps, list):
                c["employees"] = [ensure_relation_dict(x) for x in emps]
            elif isinstance(emps, dict):
                c["employees"] = [ensure_relation_dict(emps)]
            else:
                c["employees"] = []

            # Logs list (optional)
            c.setdefault("logs", [])
            c["logs"] = normalize_logs(c.get("logs", []))

            out.append(c)
        return out
    except Exception as e:
        LOG.exception("Failed to load clients.json: %s", e)
        if messagebox:
            messagebox.showerror("Load Error", f"Couldn't read clients.json:\n{e}")
        return []


def _normalize_clients_for_io(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Create a clean list of client dicts ready to be written to JSON.
    Used by both save_clients() and Export All.
    """
    out: List[Dict[str, Any]] = []
    for c in items:
        cc = dict(c)

        # Normalize relations/employees
        # For relations, preserve id format for entity links, but also keep all data fields
        normalized_relations = []
        for o in cc.get("relations", []):
            # Check if it's an entity link (has id)
            rel_link = ensure_relation_link(o)  # This now preserves all fields
            if rel_link.get("id"):
                # Entity link - ensure_relation_link already preserves all fields
                normalized_relations.append(rel_link)
            else:
                # Person relation (no link) - just normalize
                normalized_relations.append(ensure_relation_dict(o))
        cc["relations"] = normalized_relations
        cc["employees"] = [ensure_relation_dict(o) for o in cc.get("employees", [])]
        
        # Ensure officers field is removed (migrated to relations)
        cc.pop("officers", None)
        
        # Remove post_save_links - it should not be saved to JSON
        cc.pop("post_save_links", None)

        # Ensure keys exist
        for k, v in {
            "name":"", "dba":"", "ein":"", "ssn":"", "file_location":"", "memo":"",
            "addr1":"", "addr2":"", "city":"", "state":"", "zip":"",
            "acct_mgr":"", "edd_number":"", "sales_tax_account":"", "entity_type":"",
            "ui_rate":"", "sales_tax_rate":"", "other_tax_rates":"",
            "tax_rates_last_checked":"", "logs":[]
        }.items():

            if k not in cc:
                cc[k] = v

        # Drop derived fields
        cc.pop("emails", None)
        cc.pop("phones", None)

        out.append(cc)
    return out


def save_clients(items: List[Dict[str, Any]]) -> None:
    """
    Save current in-memory clients to the program's internal clients.json.
    Runs remove_stale_back_links so relation changes are reflected in data every time.
    """
    if not items:
        return
    remove_stale_back_links(items, log=LOG)
    to_save = _normalize_clients_for_io(items)
    try:
        path = DATA_FILE.resolve()
        data = json.dumps(to_save, indent=2, ensure_ascii=False)
        with path.open("w", encoding="utf-8") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        rel_counts = [len(c.get("relations") or []) for c in to_save]
        LOG.info("save_clients: wrote %s to %s (relation counts: %s)", len(to_save), path, rel_counts)
    except Exception as e:
        LOG.exception("Error writing clients.json: %s", e)
        if messagebox:
            messagebox.showerror("Save Error", f"Couldn't save clients.json:\n{e}")


def migrate_tasks_client_to_client(tasks_path: Path, remove_old_keys: bool = True) -> dict:
    """
    Convert tasks.json:
      company_idx  -> client_idx
      company_name -> client_name

    remove_old_keys=True will remove company_* fields after copying.
    """
    tasks = _read_json_file(tasks_path, default=[])
    if not isinstance(tasks, list):
        return {"updated": 0, "skipped": 0, "error": "tasks.json is not a list"}

    updated = 0
    skipped = 0

    for t in tasks:
        if not isinstance(t, dict):
            skipped += 1
            continue

        changed = False

        # idx
        if "client_idx" not in t and "company_idx" in t:
            t["client_idx"] = t.get("company_idx")
            changed = True
        if remove_old_keys and "company_idx" in t:
            t.pop("company_idx")
            changed = True

        # name
        if "client_name" not in t and "company_name" in t:
            t["client_name"] = t.get("company_name")
            changed = True
        if remove_old_keys and "company_name" in t:
            t.pop("company_name")
            changed = True

        if changed:
            updated += 1

    if updated > 0:
        _write_json_file(tasks_path, tasks)

    return {"updated": updated, "skipped": skipped}


def migrate_tasks_client_id_to_ein_ssn(tasks_path: Path, clients: List[Dict[str, Any]]) -> dict:
    """
    Migrate tasks: update client_id to match the EIN or SSN of the entity it belongs to.
    
    For each task, finds the entity by:
    1. client_idx (if available)
    2. client_name (if available) 
    3. existing client_id (if available)
    
    Then updates the task's client_id to the entity's EIN (preferred) or SSN.
    """
    tasks = _read_json_file(tasks_path, default=[])
    if not isinstance(tasks, list):
        return {"updated": 0, "skipped": 0, "error": "tasks.json is not a list"}
    
    if not isinstance(clients, list):
        return {"updated": 0, "skipped": 0, "error": "clients is not a list"}

    updated = 0
    skipped = 0
    not_found = 0
    
    # Build lookup maps for efficient matching
    clients_by_idx = {i: c for i, c in enumerate(clients) if isinstance(c, dict)}
    clients_by_name = {}
    clients_by_id = {}
    clients_by_ein = {}
    clients_by_ssn = {}
    
    for i, c in enumerate(clients):
        if not isinstance(c, dict):
            continue
        
        # Index map (already done above, but for clarity)
        clients_by_idx[i] = c
        
        # Name map (case-insensitive)
        name = (c.get("name") or "").strip()
        if name:
            clients_by_name[name.casefold()] = c
        
        # ID map
        cid = str(c.get("id", "") or "").strip()
        if cid:
            clients_by_id[cid] = c
        
        # EIN map (normalized)
        ein = normalize_ein_digits(c.get("ein", "") or "")
        if ein:
            clients_by_ein[ein] = c
        
        # SSN map (normalized)
        ssn = normalize_ssn_digits(c.get("ssn", "") or "")
        if ssn:
            clients_by_ssn[ssn] = c

    for t in tasks:
        if not isinstance(t, dict):
            skipped += 1
            continue
        
        # Find the entity this task belongs to
        entity = None
        
        # 1) Try by client_idx
        t_idx = t.get("client_idx")
        if isinstance(t_idx, int) and t_idx in clients_by_idx:
            entity = clients_by_idx[t_idx]
        
        # 2) Try by client_name
        if entity is None:
            t_name = (t.get("client_name") or "").strip()
            if t_name:
                entity = clients_by_name.get(t_name.casefold())
        
        # 3) Try by existing client_id
        if entity is None:
            t_cid = str(t.get("client_id", "") or "").strip()
            if t_cid:
                # Try direct match
                entity = clients_by_id.get(t_cid)
                # Try EIN match
                if entity is None:
                    ein_normalized = normalize_ein_digits(t_cid)
                    if ein_normalized:
                        entity = clients_by_ein.get(ein_normalized)
                # Try SSN match
                if entity is None:
                    ssn_normalized = normalize_ssn_digits(t_cid)
                    if ssn_normalized:
                        entity = clients_by_ssn.get(ssn_normalized)
        
        if entity is None:
            not_found += 1
            skipped += 1
            continue
        
        # Get EIN or SSN from entity (prefer EIN)
        ein = normalize_ein_digits(entity.get("ein", "") or "")
        ssn = normalize_ssn_digits(entity.get("ssn", "") or "")
        
        # Determine new client_id value
        new_client_id = None
        if ein:
            new_client_id = ein
        elif ssn:
            new_client_id = ssn
        
        if not new_client_id:
            # Entity has no EIN or SSN, skip
            skipped += 1
            continue
        
        # Update client_id if it's different (normalize current for comparison)
        current_client_id = str(t.get("client_id", "") or "").strip()
        current_normalized = normalize_ein_digits(current_client_id) or normalize_ssn_digits(current_client_id)
        
        # Update if different or if current is not normalized
        if current_normalized != new_client_id or current_client_id != new_client_id:
            t["client_id"] = new_client_id
            updated += 1
        else:
            skipped += 1

    if updated > 0:
        _write_json_file(tasks_path, tasks)

    return {
        "updated": updated,
        "skipped": skipped,
        "not_found": not_found
    }


def migrate_officers_to_relations(clients: List[Dict[str, Any]], remove_old_key: bool = True) -> dict:
    """
    Migrate per-client 'officers' -> 'relations', then optionally remove 'officers'.

    Rules:
      - Officers are normalized with ensure_relation_dict
      - Officers are appended into relations if not a duplicate
      - Deduping prefers id if present; else uses (name/email/phone) case-insensitive
      - Returns stats for UI reporting
    """
    updated_clients = 0
    moved = 0
    skipped_dupes = 0
    removed_keys = 0

    def _norm(s: str) -> str:
        return (s or "").strip().casefold()

    def _dedupe_key(o: dict) -> str:
        o = ensure_relation_dict(o)
        lid = str(o.get("id") or o.get("linked_client_id") or "").strip()
        if lid:
            return f"link:{lid}"
        email = _norm(str(o.get("email") or ""))
        phone = normalize_phone_digits(str(o.get("phone") or ""))
        name = _norm(str(o.get("name") or ""))
        return f"p:{name}|{email}|{phone}"

    for c in clients or []:
        if not isinstance(c, dict):
            continue

        if "officers" not in c:
            continue

        officers = c.get("officers")
        if officers is None:
            officers_list = []
        elif isinstance(officers, list):
            officers_list = officers
        elif isinstance(officers, dict):
            officers_list = [officers]
        else:
            # string/other legacy
            officers_list = [officers]

        # Ensure relations exists + normalized
        rels = c.get("relations")
        if rels is None:
            rels_list: List[Dict[str, Any]] = []
        elif isinstance(rels, list):
            rels_list = rels
        elif isinstance(rels, dict):
            rels_list = [rels]
        else:
            rels_list = [rels] if rels else []

        rels_list = [ensure_relation_dict(x) for x in rels_list]

        existing = set()
        for r in rels_list:
            existing.add(_dedupe_key(r))

        did_change = False
        for off in officers_list:
            od = ensure_relation_dict(off)
            k = _dedupe_key(od)
            if k in existing:
                skipped_dupes += 1
                continue
            rels_list.append(od)
            existing.add(k)
            moved += 1
            did_change = True

        # Write back normalized relations
        c["relations"] = rels_list

        # Remove officers key
        if remove_old_key and "officers" in c:
            c.pop("officers", None)
            removed_keys += 1
            did_change = True

        if did_change:
            updated_clients += 1

    return {
        "clients_touched": updated_clients,
        "officers_moved": moved,
        "officer_dupes_skipped": skipped_dupes,
        "officers_keys_removed": removed_keys,
    }


def _deep_merge_no_overwrite(dst, src):
    """
    Merge src into dst WITHOUT overwriting existing values.
    - dict: recurse; only fill missing keys
    - list: if both lists, append items that aren't already present (simple equality)
    - other: keep dst as-is
    """
    if isinstance(dst, dict) and isinstance(src, dict):
        for k, v in src.items():
            if k not in dst:
                dst[k] = v
            else:
                _deep_merge_no_overwrite(dst[k], v)
    elif isinstance(dst, list) and isinstance(src, list):
        for item in src:
            if item not in dst:
                dst.append(item)
    # else: keep dst unchanged


def export_all_to_json(out_path: Path, clients: list[dict]):
    """
    Export ALL program data into a single JSON file.
    Includes:
      - clients + account_managers
      - tasks + monthly_state
      - match_rules/*.json (including vendor_rules_*.json, accounts_*.json, client_list.json)
      - vendor_lists/*.csv
    """
    out_path = Path(out_path)

    def _read_text(p: Path) -> str:
        return p.read_text(encoding="utf-8", errors="replace")

    payload = {
        "version": 1,
        "clients": clients,
        "account_managers": [],
        "tasks": [],
        "monthly_state": {},
        "match_rules": {},    # filename -> parsed json
        "vendor_lists": {},   # filename -> csv text
    }

    # account managers
    if ACCOUNT_MANAGERS_FILE.exists():
        payload["account_managers"] = _read_json_file(ACCOUNT_MANAGERS_FILE, default=[])

    # tasks
    if TASKS_FILE.exists():
        payload["tasks"] = _read_json_file(TASKS_FILE, default=[])

    # monthly state
    if MONTHLY_STATE_FILE.exists():
        payload["monthly_state"] = _read_json_file(MONTHLY_STATE_FILE, default={})

    # match_rules/*.json
    if MATCH_RULES_DIR.exists():
        for p in MATCH_RULES_DIR.glob("*.json"):
            try:
                payload["match_rules"][p.name] = _read_json_file(p, default={})
            except Exception:
                # If any file is malformed, still export something rather than failing the whole export
                payload["match_rules"][p.name] = {"__raw__": _read_text(p)}

    # vendor_lists/*.csv
    if VENDOR_LISTS_DIR.exists():
        for p in VENDOR_LISTS_DIR.glob("*.csv"):
            payload["vendor_lists"][p.name] = _read_text(p)

        # some of your files are .CSV uppercase
        for p in VENDOR_LISTS_DIR.glob("*.CSV"):
            payload["vendor_lists"][p.name] = _read_text(p)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def export_selected_to_json(out_path: Path, clients: list[dict], selections: dict):
    """
    Export selected program data into a single JSON file.

    selections keys (bool):
      - clients
      - match_rules
      - monthly_data
      - tasks
      - vendor_lists
    Notes:
      - "clients" includes account_managers.json too.
      - vendor_lists are stored as filename -> csv text.
      - match_rules are stored as filename -> json content.
    """
    out_path = Path(out_path)

    payload = {
        "version": 1,
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "includes": {k: bool(v) for k, v in selections.items()},
    }

    # Clients (+ account managers)
    if selections.get("clients"):
        include_logs = bool(selections.get("include_logs", True))

        if include_logs:
            payload["clients"] = clients
        else:
            cleaned = copy.deepcopy(clients)
            for c in cleaned:
                if isinstance(c, dict):
                    c["logs"] = []
            payload["clients"] = cleaned


        if ACCOUNT_MANAGERS_FILE.exists():
            payload["account_managers"] = _read_json_file(ACCOUNT_MANAGERS_FILE, default=[])
        else:
            payload["account_managers"] = []

    # Tasks
    if selections.get("tasks"):
        payload["tasks"] = _read_json_file(TASKS_FILE, default=[]) if TASKS_FILE.exists() else []

    # Monthly data
    if selections.get("monthly_data"):
        payload["monthly_state"] = _read_json_file(MONTHLY_STATE_FILE, default={}) if MONTHLY_STATE_FILE.exists() else {}

    # Match rules (all *.json under match_rules/)
    if selections.get("match_rules"):
        rules = {}
        if MATCH_RULES_DIR.exists():
            for p in MATCH_RULES_DIR.glob("*.json"):
                rules[p.name] = _read_json_file(p, default={})
        payload["match_rules"] = rules

    # Vendor lists (all *.csv or *.CSV under vendor_lists/)
    if selections.get("vendor_lists"):
        vendor_lists = {}
        if VENDOR_LISTS_DIR.exists():
            for p in list(VENDOR_LISTS_DIR.glob("*.csv")) + list(VENDOR_LISTS_DIR.glob("*.CSV")):
                vendor_lists[p.name] = p.read_text(encoding="utf-8", errors="replace")
        payload["vendor_lists"] = vendor_lists

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def import_all_from_json(in_path: Path, clients: list[dict]) -> dict:
    """
    Import ALL data from a single export JSON.

    Rules:
      - clients: do NOT duplicate if EIN or id already exists
      - account_managers: do NOT duplicate if id already exists
      - tasks: do NOT duplicate if id already exists
              AND remap client_idx using client_name against current clients list
      - match_rules: do NOT overwrite existing json files
      - vendor_lists: overwrite
    """
    in_path = Path(in_path)
    data = _read_json_file(in_path, default={})
    if not isinstance(data, dict):
        return {"clients_added": 0, "account_managers_added": 0, "tasks_added": 0,
                "match_rules_added": 0, "vendor_lists_written": 0}

    # --- helpers ---
    def _norm_ein(x: str) -> str:
        return normalize_ein_digits(str(x or ""))

    def _norm_name(x: str) -> str:
        return " ".join(tokenize(str(x or "")))

    # --- existing client keys (id/ein/ssn) ---
    existing_client_ids = set()
    existing_client_eins = set()
    existing_client_ssns = set()

    def _norm_ssn(x: str) -> str:
        # your project stores some individuals' SSN in ssn or (legacy) ein; normalize to last 9 digits
        return normalize_ssn_digits(str(x or ""))

    for c in clients:
        if not isinstance(c, dict):
            continue

        cid = str(c.get("id", "")).strip()
        if cid:
            existing_client_ids.add(cid)

        ein = _norm_ein(c.get("ein", ""))
        if ein:
            existing_client_eins.add(ein)

        ssn = _norm_ssn(c.get("ssn", "") or c.get("ein", ""))
        if ssn:
            existing_client_ssns.add(ssn)

    # --- import clients ---
    clients_added = 0
    clients_skipped_duplicates = 0

    incoming_clients = data.get("clients", [])
    if isinstance(incoming_clients, list):
        for c in incoming_clients:
            if not isinstance(c, dict):
                continue

            cid = str(c.get("id", "")).strip()
            ein = _norm_ein(c.get("ein", ""))
            ssn = _norm_ssn(c.get("ssn", "") or c.get("ein", ""))

            # Skip if ANY identifier matches an existing client
            if cid and cid in existing_client_ids:
                clients_skipped_duplicates += 1
                continue
            if ein and ein in existing_client_eins:
                clients_skipped_duplicates += 1
                continue
            if ssn and ssn in existing_client_ssns:
                clients_skipped_duplicates += 1
                continue

            clients.append(c)
            clients_added += 1

            if cid:
                existing_client_ids.add(cid)
            if ein:
                existing_client_eins.add(ein)
            if ssn:
                existing_client_ssns.add(ssn)


    # Build name->idx map AFTER possibly adding clients
    name_to_idx = {}
    for idx, c in enumerate(clients):
        nm = _norm_name(c.get("name", ""))
        if nm and nm not in name_to_idx:
            name_to_idx[nm] = idx

    # --- import account managers (dedupe by (name/email/phone) + normalize to deterministic ids) ---
    am_added = 0
    if isinstance(data.get("account_managers"), list):
        existing_ams = _read_json_file(ACCOUNT_MANAGERS_FILE, default=[])
        if not isinstance(existing_ams, list):
            existing_ams = []

        # Normalize existing first (adds ids + dedupes)
        norm_existing = []
        seen_keys = set()
        for x in existing_ams:
            if not isinstance(x, dict):
                x = {"name": str(x)}
            key = _account_manager_key(x)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            amid = str(x.get("id", "") or "").strip()
            if not amid:
                amid = _account_manager_id_from_key(key)
            norm_existing.append({
                "id": amid,
                "name": str(x.get("name", "") or "").strip(),
                "email": str(x.get("email", "") or "").strip(),
                "phone": str(x.get("phone", "") or "").strip(),
            })

        existing_ams = norm_existing
        existing_keys = { _account_manager_key(x) for x in existing_ams if isinstance(x, dict) }

        for am in data["account_managers"]:
            if not isinstance(am, dict):
                continue

            key = _account_manager_key(am)
            if key in existing_keys:
                continue

            amid = str(am.get("id", "") or "").strip()
            if not amid:
                amid = _account_manager_id_from_key(key)

            existing_ams.append({
                "id": amid,
                "name": str(am.get("name", "") or "").strip(),
                "email": str(am.get("email", "") or "").strip(),
                "phone": str(am.get("phone", "") or "").strip(),
            })
            existing_keys.add(key)
            am_added += 1

        ACCOUNT_MANAGERS_FILE.parent.mkdir(parents=True, exist_ok=True)
        ACCOUNT_MANAGERS_FILE.write_text(json.dumps(existing_ams, indent=2, ensure_ascii=False), encoding="utf-8")


    # --- import monthly_state (safe: write if file missing; else keep existing) ---
    if isinstance(data.get("monthly_state"), dict):
        if not MONTHLY_STATE_FILE.exists():
            MONTHLY_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
            MONTHLY_STATE_FILE.write_text(json.dumps(data["monthly_state"], indent=2, ensure_ascii=False), encoding="utf-8")

    # --- import match_rules (do NOT overwrite existing files) ---
    mr_added = 0
    match_rules = data.get("match_rules", {})
    if isinstance(match_rules, dict):
        MATCH_RULES_DIR.mkdir(parents=True, exist_ok=True)
        for fname, contents in match_rules.items():
            try:
                fname = str(fname)
                if not fname.lower().endswith(".json"):
                    continue
                dest = MATCH_RULES_DIR / fname
                if dest.exists():
                    continue  # do not overwrite
                # contents can be dict/list/etc.
                dest.write_text(json.dumps(contents, indent=2, ensure_ascii=False), encoding="utf-8")
                mr_added += 1
            except Exception:
                pass

    # --- import vendor_lists (overwrite) ---
    vl_written = 0
    vendor_lists = data.get("vendor_lists", {})
    if isinstance(vendor_lists, dict):
        VENDOR_LISTS_DIR.mkdir(parents=True, exist_ok=True)
        for fname, csv_text in vendor_lists.items():
            try:
                fname = str(fname)
                if not fname.lower().endswith((".csv",)):
                    continue
                dest = VENDOR_LISTS_DIR / fname
                dest.write_text(csv_text, encoding="utf-8")
                vl_written += 1
            except Exception:
                pass

    # --- import tasks (remap client_idx by name) ---
    tasks_added = 0
    incoming_tasks = data.get("tasks", [])
    if isinstance(incoming_tasks, list):
        existing_tasks = _read_json_file(TASKS_FILE, default=[])
        if not isinstance(existing_tasks, list):
            existing_tasks = []

        existing_task_ids = {str(t.get("id", "")).strip() for t in existing_tasks if isinstance(t, dict)}

        for t in incoming_tasks:
            if not isinstance(t, dict):
                continue

            tid = str(t.get("id", "")).strip()
            if tid and tid in existing_task_ids:
                continue

            # Remap client_idx by name if present
            client_name = str(t.get("client_name", "")).strip()
            if client_name:
                mapped_idx = name_to_idx.get(_norm_name(client_name))
                if mapped_idx is not None:
                    t["client_idx"] = mapped_idx

            existing_tasks.append(t)
            if tid:
                existing_task_ids.add(tid)
            tasks_added += 1

        if tasks_added > 0:
            _write_json_file(TASKS_FILE, existing_tasks)

    return {
        "clients_added": clients_added,
        "clients_skipped_duplicates": clients_skipped_duplicates,
        "account_managers_added": am_added,
        "tasks_added": tasks_added,
        "match_rules_added": mr_added,
        "vendor_lists_written": vl_written,
    }
