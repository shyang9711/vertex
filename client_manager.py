from __future__ import annotations

import sys, pathlib, types
import ctypes
from ctypes import wintypes

# Make imports work in BOTH layouts:
_BASE = pathlib.Path(__file__).resolve().parent

# Ensure "functions" is importable even when the repo root IS the functions folder
if "functions" not in sys.modules:
    _pkg = types.ModuleType("functions")
    _pkg.__path__ = [str(_BASE)]   # so "functions.pages" resolves to "<base>/pages"
    sys.modules["functions"] = _pkg

# Make sure Python can import sibling folders like pages/, models/, utils/ if needed
if str(_BASE) not in sys.path:
    sys.path.insert(0, str(_BASE))

# Keep legacy behavior (in case you still run from <Scripts>/)
_PARENT = _BASE.parent
if str(_PARENT) not in sys.path:
    sys.path.insert(0, str(_PARENT))

import os, sys, json, re, hashlib, webbrowser, subprocess, datetime as dt, urllib.request, urllib.error, ssl, urllib.parse
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import tkinter.font as tkfont
import csv

# pages
try:
    from vertex.pages.dashboard_page import DashboardPage
    from vertex.pages.profile_tab import init_profile_tab
    from vertex.pages.documents_tab import init_documents_tab
    from vertex.pages.logs_tab import init_logs_tab
    from vertex.pages.checklist_page import ChecklistPage
    from vertex.pages.action_page import ActionRunnerPage
    from vertex.pages.reports_page import ReportsPage
    from vertex.pages.note_page import NotePage

    from vertex.models.taskbar_model import TaskbarModel
    from vertex.utils.app_logging import get_logger
    from vertex.utils.app_update import check_for_updates, enforce_major_update_on_startup
    from vertex.utils.helpers import (
        ensure_relation_dict, display_relation_name,
        ensure_relation_link, merge_relations,
        migrate_officer_business_links_to_relations,
        is_migration_done, mark_migration_done,
    )

    
    from vertex.config import APP_NAME, UPDATE_POLICY_ASSET_NAME, APP_VERSION, GITHUB_REPO, GITHUB_RELEASES_URL, GITHUB_API_LATEST, UPDATE_ASSET_NAME, ENTITY_TYPES, US_STATES, ROLES

    from vertex.ui.dialogs.clientdialog import ClientDialog
    from vertex.ui.dialogs.linkdialog import LinkDialog
    from vertex.ui.components.autocomplete import AutocompletePopup
    from vertex.ui.components.scrollframe import ScrollFrame

except ModuleNotFoundError:
    from pages.dashboard_page import DashboardPage
    from pages.profile_tab import init_profile_tab
    from pages.documents_tab import init_documents_tab
    from pages.logs_tab import init_logs_tab
    from pages.checklist_page import ChecklistPage
    from pages.action_page import ActionRunnerPage
    from pages.reports_page import ReportsPage
    from pages.note_page import NotePage

    from models.taskbar_model import TaskbarModel
    from utils.app_logging import get_logger
    from utils.app_update import check_for_updates, enforce_major_update_on_startup
    from utils.helpers import (
    ensure_relation_dict, display_relation_name,
    ensure_relation_link, merge_relations,
    migrate_officer_business_links_to_relations,
    is_migration_done, mark_migration_done,
    )
    
    from config import APP_NAME, UPDATE_POLICY_ASSET_NAME, APP_VERSION, GITHUB_REPO, GITHUB_RELEASES_URL, GITHUB_API_LATEST, UPDATE_ASSET_NAME, ENTITY_TYPES, US_STATES, ROLES

    from ui.dialogs.clientdialog import ClientDialog
    from ui.dialogs.linkdialog import LinkDialog
    from ui.components.autocomplete import AutocompletePopup
    from ui.components.scrollframe import ScrollFrame

# NewUI preference from styles/, fallback to functions/
try:
    from styles.new_ui import NewUI
except Exception:
    NewUI = None

LOG = get_logger("client_manager")

# -----------------------------
# Single-instance (Windows mutex)
# -----------------------------
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_user32 = ctypes.WinDLL("user32", use_last_error=True)

ERROR_ALREADY_EXISTS = 183
MB_OK = 0x00000000
MB_ICONINFORMATION = 0x00000040
MB_TOPMOST = 0x00040000

_CreateMutexW = _kernel32.CreateMutexW
_CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
_CreateMutexW.restype = wintypes.HANDLE

_GetLastError = _kernel32.GetLastError
_GetLastError.argtypes = []
_GetLastError.restype = wintypes.DWORD

_MessageBoxW = _user32.MessageBoxW
_MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]
_MessageBoxW.restype = ctypes.c_int

_single_instance_mutex_handle = None

def enforce_single_instance(app_id: str = "Vertex") -> bool:
    """
    Returns True if this is the first instance.
    Returns False if another instance is already running (and exits caller typically).
    """
    global _single_instance_mutex_handle

    # Use a stable global mutex name
    mutex_name = f"Global\\{app_id}_SingleInstance_Mutex"

    h = _CreateMutexW(None, False, mutex_name)
    if not h:
        # If mutex creation fails, don't block startup (but you can choose to)
        return True

    _single_instance_mutex_handle = h
    last_err = _GetLastError()

    if last_err == ERROR_ALREADY_EXISTS:
        # Another instance already created the mutex
        _MessageBoxW(
            None,
            "Vertex is already running.\n\nPlease close the existing Vertex window before opening a new one.",
            "Vertex",
            MB_OK | MB_ICONINFORMATION | MB_TOPMOST,
        )
        return False

    return True

# -------------------- Storage paths --------------------
def portable_root() -> Path:
    # For EXE: folder that contains vertex.exe (NOT _MEI)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # For source run
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path(os.getcwd()).resolve()

def appdata_root() -> Path:
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

for _p in (CLIENTS_DIR, TASKS_DIR, MATCH_RULES_DIR, MONTHLY_DATA_DIR, VENDOR_LISTS_DIR):
    _p.mkdir(parents=True, exist_ok=True)

DATA_DIR  = CLIENTS_DIR
DATA_FILE = CLIENTS_DIR / "clients.json"

ACCOUNT_MANAGERS_FILE = CLIENTS_DIR / "account_managers.json"
TASKS_FILE            = TASKS_DIR / "tasks.json"
MONTHLY_STATE_FILE    = MONTHLY_DATA_DIR / "monthly_state.json"
client_LIST_FILE     = MATCH_RULES_DIR / "client_list.json"


# -------------------- System Fault Handler -----------
import faulthandler, sys, os, tempfile
_log = os.path.join(tempfile.gettempdir(), "lineupcpa_fatal.log")
try:
    faulthandler.enable(file=open(_log, "w"))
except Exception:
    faulthandler.enable()  # fallback to stderr


# -------------------- Helpers --------------------
PHONE_DIGITS_RE = re.compile(r"\d")

def normalize_phone_digits(s: str) -> str:
    digits = "".join(PHONE_DIGITS_RE.findall(s or ""))
    return digits[-10:] if len(digits) >= 10 else digits

def normalize_ein_digits(s: str) -> str:
    return "".join(PHONE_DIGITS_RE.findall(s or ""))[-9:]

def normalize_ssn_digits(s: str) -> str:
    return "".join(PHONE_DIGITS_RE.findall(s or ""))[-9:]

def normalize_logs(logs):
    out = []
    for x in logs or []:
        if isinstance(x, dict):
            out.append({
                "ts":   str(x.get("ts","")).strip(),
                "user": str(x.get("user","")).strip(),
                "text": str(x.get("text","")).strip(),
                "done": bool(x.get("done", False)),
            })
        else:
            out.append({"ts":"", "user":"", "text":str(x), "done": False})
    return out

_AM_WS_RE = re.compile(r"\s+")

def _account_manager_key(am: dict) -> str:
    """Stable, case-insensitive key for deduping account managers."""
    if not isinstance(am, dict):
        am = {"name": str(am)}
    name  = _AM_WS_RE.sub(" ", str(am.get("name", "") or "").strip()).casefold()
    email = str(am.get("email", "") or "").strip().casefold()
    phone = normalize_phone_digits(str(am.get("phone", "") or "").strip())
    return f"{name}|{email}|{phone}"

def _account_manager_id_from_key(key: str) -> str:
    """Deterministic short id from key (so imports don't create duplicates)."""
    h = hashlib.sha1((key or "").encode("utf-8", errors="ignore")).hexdigest()
    return f"am_{h[:12]}"

def tokenize(s: str) -> List[str]:
    if s is None: return []
    s = str(s).lower().strip()
    parts = re.split(r"[^a-z0-9@._\-&]+", s)
    return [p for p in parts if p]

def norm_text(s: str) -> str:
    return " ".join(tokenize(s))

def display_relation_name(o: Dict[str, str]) -> str:
    o = ensure_relation_dict(o)
    return o.get("name","").strip()

def relations_to_display_lines(relations: List[Dict[str,str]]) -> List[str]:
    return [display_relation_name(o) for o in (relations or []) if display_relation_name(o)]

def relations_to_flat_emails(relations: List[Dict[str,str]]) -> List[str]:
    seen, out = set(), []
    for o in relations or []:
        e = str(ensure_relation_dict(o).get("email","")).strip()
        if e and e not in seen:
            seen.add(e); out.append(e)
    return out

def relations_to_flat_phones(relations: List[Dict[str,str]]) -> List[str]:
    seen, out = set(), []
    for o in relations or []:
        p = str(ensure_relation_dict(o).get("phone","")).strip()
        if p and p not in seen:
            seen.add(p); out.append(p)
    return out

def is_valid_person_payload(data):
    return isinstance(data, (tuple, list)) and len(data) == 3 and data[0] is not None

# ---- Quarter / Tax helpers ---------------------------------------------------
def today_date() -> dt.date:
    return dt.date.today()

def quarter_start(d: dt.date) -> dt.date:
    q = (d.month - 1) // 3
    first_month = q * 3 + 1
    return dt.date(d.year, first_month, 1)

def new_quarter_started(last_checked_iso: str | None) -> bool:
    try:
        if not last_checked_iso:
            return True
        last = dt.date.fromisoformat(last_checked_iso)
    except Exception:
        return True
    start_now = quarter_start(today_date())
    return last < start_now

def safe_fetch_sales_tax_rate(state: str, city: str) -> Optional[float]:
    """
    Optional API call to fetch a sales tax rate by state/city.
    Configure env var TAX_API_URL like:
      https://example.com/tax?state={state}&city={city}
    Returns None on failure/offline.
    """
    state = (state or "").strip()
    city  = (city or "").strip()
    if not state:
        return None
    tmpl = os.environ.get("TAX_API_URL", "")
    if not tmpl:
        return None
    url = tmpl.format(state=state, city=urllib.parse.quote(city))
    try:
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(url, timeout=5, context=ctx) as resp:
            if resp.status != 200:
                return None
            data = resp.read().decode("utf-8", errors="ignore")
            m = re.search(r"(\d+(?:\.\d+)?)", data)
            if not m: return None
            val = float(m.group(1))
            if val <= 0.2:  # treat as fraction
                val *= 100.0
            return round(val, 4)
    except Exception:
        return None

# -------------------- Data IO --------------------
def load_clients() -> List[Dict[str, Any]]:
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

            # Personnel
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
            c["relations"] = norm_offs
            c.pop("owner", None)

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
        cc["relations"]  = [ensure_relation_dict(o) for o in cc.get("relations", [])]
        cc["employees"] = [ensure_relation_dict(o) for o in cc.get("employees", [])]

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
    This is what 'Save Data' uses.
    """
    to_save = _normalize_clients_for_io(items)
    try:
        DATA_FILE.write_text(json.dumps(to_save, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        LOG.exception("Error writing clients.json: %s", e)
        messagebox.showerror("Save Error", f"Couldn't save clients.json:\n{e}")


def _read_json_file(path: Path, default):
    try:
        if not path.exists():
            return default
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default

def _write_json_file(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8")

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

        # name
        if "client_name" not in t and "company_name" in t:
            t["client_name"] = t.get("company_name")
            changed = True

        if remove_old_keys:
            if "company_idx" in t:
                t.pop("company_idx", None)
                changed = True
            if "company_name" in t:
                t.pop("company_name", None)
                changed = True

        if changed:
            updated += 1
        else:
            skipped += 1

    _write_json_file(tasks_path, tasks)
    return {"updated": updated, "skipped": skipped, "path": str(tasks_path)}

def migrate_officers_to_relations(clients: List[Dict[str, Any]], remove_old_key: bool = True) -> dict:
    """
    Migrate per-client 'officers' -> 'relations', then optionally remove 'officers'.

    Rules:
      - Officers are normalized with ensure_relation_dict
      - Officers are appended into relations if not a duplicate
      - Deduping prefers linked_client_id if present; else uses (name/email/phone) case-insensitive
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
        lid = str(o.get("linked_client_id") or "").strip()
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
                dst[k] = _deep_merge_no_overwrite(dst[k], v)
        return dst
    if isinstance(dst, list) and isinstance(src, list):
        for item in src:
            if item not in dst:
                dst.append(item)
        return dst
    return dst

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
            import copy
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

    # --- import vendor_lists (OVERWRITE) ---
    vl_written = 0
    vendor_lists = data.get("vendor_lists", {})
    if isinstance(vendor_lists, dict):
        VENDOR_LISTS_DIR.mkdir(parents=True, exist_ok=True)
        for fname, text in vendor_lists.items():
            try:
                fname = str(fname)
                if not (fname.lower().endswith(".csv")):
                    continue
                dest = VENDOR_LISTS_DIR / fname
                dest.write_text(str(text), encoding="utf-8", errors="replace")
                vl_written += 1
            except Exception:
                pass

    # --- import tasks (no dup by id, AND remap client_idx) ---
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

            # remap client_idx based on client_name (so tasks show under correct client)
            cname = _norm_name(t.get("client_name", ""))
            if cname and cname in name_to_idx:
                t["client_idx"] = name_to_idx[cname]

            existing_tasks.append(t)
            tasks_added += 1
            if tid:
                existing_task_ids.add(tid)

        TASKS_FILE.parent.mkdir(parents=True, exist_ok=True)
        TASKS_FILE.write_text(json.dumps(existing_tasks, indent=2, ensure_ascii=False), encoding="utf-8")

    return {
        "clients_added": clients_added,
        "clients_skipped_duplicates": clients_skipped_duplicates,
        "account_managers_added": am_added,
        "tasks_added": tasks_added,
        "match_rules_added": mr_added,
        "vendor_lists_written": vl_written,
    }


def show_about_dialog(parent: tk.Misc | None = None):
    msg = (
        f"{APP_NAME}\n"
        f"Version {APP_VERSION}\n\n"
        f"Data folder:\n{DATA_ROOT}"
    )
    messagebox.showinfo("About", msg, parent=parent)

# -------------------- Log Dialog --------------------
class LogDialog(tk.Toplevel):
    def __init__(self, master, title="Log Entry", initial=None):
        super().__init__(master)
        self.title(title); self.resizable(False, False); self.result=None
        init = initial or {}
        self.v_ts   = tk.StringVar(value=init.get("ts",""))
        self.v_user = tk.StringVar(value=init.get("user",""))
        frm = ttk.Frame(self, padding=12); frm.grid(row=0, column=0, sticky="nsew")

        def row(lbl, var, r, width=40):
            ttk.Label(frm, text=lbl).grid(row=r, column=0, sticky="w", pady=(0,2))
            ent = ttk.Entry(frm, textvariable=var, width=width)
            ent.grid(row=r+1, column=0, sticky="we", pady=(0,6)); return ent

        e1 = row("Timestamp (ISO, optional)", self.v_ts, 0, 40)
        row("User/Initials", self.v_user, 2, 20)

        ttk.Label(frm, text="Text").grid(row=4, column=0, sticky="w", pady=(0,2))
        self.txt = ScrolledText(frm, width=56, height=8, wrap="word")
        self.txt.grid(row=5, column=0, sticky="nsew")
        if "text" in init: self.txt.insert("1.0", init["text"])

        btns = ttk.Frame(frm); btns.grid(row=6, column=0, sticky="e", pady=(8,0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda _e: self._save()); self.bind("<Escape>", lambda _e: self.destroy())
        self.after(50, e1.focus_set); self.grab_set(); self.transient(master)

    def _save(self):
        import datetime
        ts = self.v_ts.get().strip()
        if not ts:
            # Fill now in ISO if user left blank
            ts = datetime.datetime.now().isoformat(timespec="seconds")
        self.result = {
            "ts": ts,
            "user": self.v_user.get().strip(),
            "text": self.txt.get("1.0", "end").strip(),
        }
        self.destroy()


# -------------------- Main App with Navigation --------------------
class App(ttk.Frame):
    COLS = ("name",
            "dba",
            "relations",
            "ein",
            "emails",
            "phones",
            # "memo"
           )

    def __init__(self, master):
        super().__init__(master, padding=(10,8))
        self.pack(fill="both", expand=True)
        self.log = get_logger("client_manager")
        self.log.info("App init")
        self.root = master

        self.items: List[Dict[str, Any]] = load_clients()

        self.style = ttk.Style()
        try: self.style.theme_use("clam")
        except Exception: pass
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.base_row_px = int(self.default_font.metrics("linespace") * 1.35)

        style = ttk.Style()
        style.configure("Treeview", rowheight=40)
        self._last_viewed_idx = None

        # Data Import
        self.DATA_ROOT = DATA_ROOT
        self.CLIENTS_DIR = CLIENTS_DIR
        self.TASKS_DIR = TASKS_DIR
        self.MATCH_RULES_DIR = MATCH_RULES_DIR
        self.MONTHLY_DATA_DIR = MONTHLY_DATA_DIR
        self.VENDOR_LISTS_DIR = VENDOR_LISTS_DIR

        # --- Manager filter state ---
        self._mgr_filter_active = set()   # empty set = All managers
        self._mgr_menu = None


        # --- App menu / taskbar (very top, above the Back/Forward + Search row)
        
        self.account_managers = self._load_account_managers()

        def _get_acct_mgrs():
            # Return a fresh copy (dialog can mutate its own copy)
            return list(self.account_managers)

        def _set_acct_mgrs(lst):
            # Save to disk and keep in memory
            self.account_managers = self._normalize_acct_mgr_list(lst)
            self._save_account_managers(self.account_managers)
            
        self.get_account_managers = _get_acct_mgrs
        self.set_account_managers = _set_acct_mgrs

        def _open_prefs():
            # If you already have a Preferences dialog, call it here.
            # Otherwise, the TaskbarModel can show its default prefs dialog.
            try:
                TaskbarModel(self.winfo_toplevel())._default_prefs_dialog()
            except Exception:
                pass

        self.taskbar = TaskbarModel(
            self.winfo_toplevel(),
            # Close the main window / app
            on_exit=lambda: self.winfo_toplevel().destroy(),

            # New client menu item
            on_new_client=self.on_new,

            # Account manager list dialog
            get_account_managers=_get_acct_mgrs,
            set_account_managers=_set_acct_mgrs,

            # Preferences
            on_open_preferences=_open_prefs,

            # Action → Client Actions…
            on_open_actions=lambda: self.open_actions_page(),

            # Edit → Batch Task Settings…
            on_open_batch_tasks=lambda: self.dashboard.open_batch_task_settings(),

            # File → Save / Import / Export
            on_save_data=self._save_all_data,
            on_import_data=self._import_data_dialog,
            on_export_data=self._export_selected_dialog,
            on_update_data=self._update_data_dialog,

            # Update
            on_check_updates=lambda: check_for_updates(
                    parent=self.winfo_toplevel(),
                    app_name=APP_NAME,
                    app_version=APP_VERSION,
                    github_api_latest=GITHUB_API_LATEST,
                    github_releases_url=GITHUB_RELEASES_URL,
                    update_asset_name=UPDATE_ASSET_NAME,
                ),
            on_about=self._show_about,
        )

        # Top bar
        topbar = ttk.Frame(self, style="Card.TFrame", padding=(10, 8)); topbar.pack(side=tk.TOP, fill=tk.X)
        self.btn_home = ttk.Button(topbar, text="🏠 Home", command=self.go_home, width=10, style="NewUI.TButton")
        self.btn_back = ttk.Button(topbar, text="◀ Back", command=self.nav_back, state=tk.DISABLED, width=10, style="NewUI.TButton")
        self.btn_fwd  = ttk.Button(topbar, text="Forward ▶", command=self.nav_forward, state=tk.DISABLED, width=10, style="NewUI.TButton")
        
        self.btn_home.pack(side=tk.LEFT)
        self.btn_back.pack(side=tk.LEFT); self.btn_fwd.pack(side=tk.LEFT, padx=(6,12))
        self.btn_back.bind("<Button-3>", self._show_back_menu)
        self.btn_fwd.bind("<Button-3>", self._show_forward_menu)
        ttk.Frame(topbar).pack(side=tk.LEFT, expand=True)

        self.q = tk.StringVar()
        search_box = ttk.Frame(topbar)
        search_box.pack(side=tk.LEFT)

        ttk.Label(search_box, text="Search", style="OnPanel.TLabel").pack(side=tk.LEFT, padx=(0, 6))
        self.search_entry = ttk.Entry(search_box, textvariable=self.q, width=46)
        self.search_entry.pack(side=tk.LEFT)
        
        ttk.Button(search_box, text="✕",   command=self.clear_search).pack(side=tk.LEFT, padx=(6,0))
        
        # Status
        self.status = tk.StringVar(value=f"Loaded {len(self.items)} clients. Data: {DATA_FILE}")
        ttk.Label(self, textvariable=self.status).pack(side=tk.BOTTOM, fill=tk.X, pady=(6,0), anchor="w")

        # Page host
        self.page_host = ttk.Frame(self)
        self.page_host.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(6,0))

        self.dashboard = DashboardPage(self)
        self.notes = NotePage(self)

        # Main page
        self._build_main_page()

        # Autocomplete
        self._ac = AutocompletePopup(self, self.search_entry, on_choose=self._open_from_suggestion)

        # Hide suggestions when clicking anywhere outside the popup or the entry
        def _dismiss_ac_on_click(e):
            if not self._ac or not self._ac.winfo_exists() or not self._ac.winfo_viewable():
                return
            w = e.widget
            in_popup = str(w).startswith(str(self._ac))
            on_entry = (w is self.search_entry) or str(w).startswith(str(self.search_entry))
            if not (in_popup or on_entry):
                self._ac.hide()

        def _on_q_change(*_):
            self._update_suggestions()
            if getattr(self, "_current_page", (None,))[0] == "search":
                self.populate()

        self.q.trace_add("write", _on_q_change)

        def _submit_top_search(_e=None):
            if not self.q.get().strip():
                return "break"
            self.navigate("search", None, push=True)
            self._ac.hide()
            return "break"

        def _on_search_up(e):
            if not self._ac.winfo_viewable():
                self._update_suggestions()
                if not self._ac.winfo_viewable():
                    return "break"
            self._ac.focus_listbox()
            self._ac.listbox.event_generate("<Up>")
            return "break"

        def _on_search_down(e):
            if not self._ac.winfo_viewable():
                self._update_suggestions()
                if not self._ac.winfo_viewable():
                    return "break"
            self._ac.focus_listbox()
            self._ac.listbox.event_generate("<Down>")
            return "break"

        def _on_search_return(e):
            txt = self._ac.current_text() or self.search_entry.get()
            self._open_from_suggestion(txt)
            return "break"

        def _entry_focus_out(e):
            def _check():
                w = self.focus_get()
                if w and str(w).startswith(str(self._ac)):
                    return
                self._ac.hide()
            self.after(1, _check)

        def _search_click_clear(_e):
            self.search_entry.selection_clear()

        def _focus_search_accel(e):
            if not (e.state & 0x4):
                return
            if self.focus_get() is not self.search_entry:
                self.search_entry.focus_set()
                self.search_entry.select_range(0, tk.END)
            else:
                self.search_entry.focus_set()
            return "break"

        def _ensure_bidirectional_link(
            self,
            a_id: str,
            b_id: str,
            a_label: str,
            b_label: str,
            a_is_business: bool,
            b_is_business: bool,
        ):
            def find(cid):
                for c in self.items:
                    if str(c.get("id", "")) == cid:
                        return c
                return None

            a = find(a_id)
            b = find(b_id)
            if not a or not b:
                return

            def ensure_link(src, dst_id, dst_label, role):
                offs = src.setdefault("relations", [])
                for o in offs:
                    if o.get("linked_client_id") == dst_id:
                        return
                offs.append({
                    "name": dst_label,
                    "first_name": dst_label,
                    "role": role,
                    "linked_client_id": dst_id,
                    "linked_client_label": dst_label,
                })

            # Role rules
            role_a_to_b = "business" if b_is_business else ""
            role_b_to_a = "business" if a_is_business else ""

            ensure_link(a, b_id, b_label, role_a_to_b)
            ensure_link(b, a_id, a_label, role_b_to_a)

        self.search_entry.bind("<KeyPress-Down>", _on_search_down)
        self.search_entry.bind("<KeyPress-Up>", _on_search_up)
        self.search_entry.bind("<Return>", _submit_top_search)
        self.search_entry.bind("<Escape>", lambda e: (self._ac.hide(), "break"))
        self.search_entry.bind("<FocusOut>", _entry_focus_out)
        self.search_entry.bind("<Control-a>",  lambda e: (e.widget.select_range(0, 'end'), e.widget.icursor('end'), "break"))

        # Quick search: Ctrl/Cmd+F focuses search
        self.winfo_toplevel().bind("<Control-f>", _focus_search_accel)
        self.winfo_toplevel().bind("<Command-f>", _focus_search_accel)

        # Primary actions
        self.bind_all("<Control-n>", lambda e: self.on_new())
        self.bind_class("Treeview", "<Delete>", self._on_delete_from_tree)
        self.bind_all("<Button-1>", _search_click_clear, add="+")
        self.search_entry.bind("<FocusIn>",  _search_click_clear, add="+")

        self.bind_all("<Alt-Left>",  lambda e: self.nav_back())
        self.bind_all("<Alt-Right>", lambda e: self.nav_forward())

        # History stacks
        self._history: List[Tuple[str, Any]] = []
        self._future:  List[Tuple[str, Any]] = []
        self._current_page: Tuple[str, Any] = ("main", None)

        # Start
        self.log.info("Navigate start -> main")
        self.navigate("main", None, push=False)
        self.populate()
        self.update_sales_tax_rates_if_due()

        # --- Manager filter state
        self._mgr_filter_active = set()
        self._mgr_menu = None

    def _show_about(self):
        show_about_dialog(self.winfo_toplevel())

    # Taskbar helper
    def _clients_folder(self) -> str:
        """
        Resolve the clients/ folder.
        Priority:
        1) A 'clients' folder next to this file.
        2) A 'clients' folder in current working directory.
        3) Create a 'clients' next to this file if missing.
        """
        here = os.path.dirname(os.path.abspath(__file__))
        p1 = os.path.join(here, "clients")
        if os.path.isdir(p1):
            return p1
        p2 = os.path.join(os.getcwd(), "clients")
        if os.path.isdir(p2):
            return p2
        os.makedirs(p1, exist_ok=True)
        return p1

    def _account_managers_path(self) -> Path:
        # Prefer instance attribute when available (EXE-safe)
        base = getattr(self, "CLIENTS_DIR", None)
        if base:
            p = Path(base) / "account_managers.json"
        else:
            # Fallback to module-level constants (still correct)
            p = ACCOUNT_MANAGERS_FILE

        p.parent.mkdir(parents=True, exist_ok=True)
        return p

    
    def _account_manager_names(self):
        """Return list[str] of manager names for dropdown."""
        lst = getattr(self, "account_managers", []) or []
        names = []
        for x in lst:
            try:
                if isinstance(x, dict):
                    nm = (x.get("name") or "").strip()
                else:
                    nm = str(x).strip()
            except Exception:
                nm = ""
            if nm:
                names.append(nm)
        # unique + stable order
        seen = set(); out = []
        for nm in names:
            k = nm.casefold()
            if k in seen: continue
            seen.add(k)
            out.append(nm)
        return out

    def _make_mgr_combobox(self, parent, *, initial_name=""):
        """Create a readonly Combobox of account managers; returns (combobox, StringVar)."""
        v = tk.StringVar(value=initial_name or "")
        cb = ttk.Combobox(parent, textvariable=v, state="readonly", values=self._account_manager_names(), width=30)
        cb.configure(state="normal")
        cb["values"] = self._account_manager_names()
        def _validate_and_lock(event=None):
            names = self._account_manager_names()
            cur = (v.get() or "").strip()
            if cur:
                exists_ci = any(n.casefold() == cur.casefold() for n in names)
                if not exists_ci:
                    am = getattr(self, "account_managers", []) or []
                    am.append({"name": cur})
                    self.account_managers = self._normalize_acct_mgr_list(am)
                    self._save_account_managers(self.account_managers)
                    cb["values"] = self._account_manager_names()
                else:
                    for n in names:
                        if n.casefold() == cur.casefold():
                            v.set(n)
                            break
            cb.configure(state="readonly")

        cb.bind("<<ComboboxSelected>>", _validate_and_lock)
        cb.bind("<FocusOut>", _validate_and_lock)
        return cb, v

    def _normalize_acct_mgr_list(self, items):
        """
        Ensure list of dicts with keys: id, name, email, phone.
        Accepts list[str] or list[dict].

        - id is deterministic (derived from name/email/phone) so imports/exports won't duplicate.
        - Dedupes case-insensitively by (name,email,phone).
        """
        out = []
        seen_keys = set()

        for x in (items or []):
            if isinstance(x, dict):
                am = {
                    "id":    str(x.get("id", "") or "").strip(),
                    "name":  str(x.get("name", "") or "").strip(),
                    "email": str(x.get("email", "") or "").strip(),
                    "phone": str(x.get("phone", "") or "").strip(),
                }
            else:
                s = str(x).strip()
                if not s:
                    continue
                am = {"id": "", "name": s, "email": "", "phone": ""}

            key = _account_manager_key(am)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            if not am["id"]:
                am["id"] = _account_manager_id_from_key(key)

            out.append(am)

        return out


    def _load_account_managers(self):
        """Load from clients/account_managers.json; return [] if missing/error."""
        path = self._account_managers_path()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return self._normalize_acct_mgr_list(data)
        except FileNotFoundError:
            return []
        except Exception as e:
            try:
                messagebox.showwarning("Account Managers", f"Failed to load:\n{e}")
            except Exception:
                pass
            return []

    def _save_account_managers(self, lst):
        """Save to clients/account_managers.json (creates folder if needed)."""
        lst = self._normalize_acct_mgr_list(lst)
        path = self._account_managers_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(lst, f, ensure_ascii=False, indent=2)
            try:
                if hasattr(self, "status"):
                    self.status.set(f"Saved {len(lst)} account manager(s) → {path}")
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Account Managers", f"Failed to save:\n{e}")

    def _update_data_dialog(self):
        try:
            if not messagebox.askyesno(
                "Update Data",
                "This will update data files.\n\nProceed?"
            ):
                return

            stats_tasks = {"updated": 0, "skipped": 0, "path": str(TASKS_FILE)}
            stats_off   = {"clients_touched": 0, "officers_moved": 0, "officer_dupes_skipped": 0, "officers_keys_removed": 0}

            # 1) tasks.json migration (company_* -> client_*)
            if not is_migration_done(MIG_TASKS_CLIENT_TO_CLIENT):
                stats_tasks = migrate_tasks_client_to_client(TASKS_FILE, remove_old_keys=True)
                mark_migration_done(MIG_TASKS_CLIENT_TO_CLIENT, {"path": str(TASKS_FILE)})
            else:
                stats_tasks = {"updated": 0, "skipped": -1, "path": str(TASKS_FILE)}  # skipped=-1 => already migrated

            # 2) clients.json migration (officers -> relations)
            if not is_migration_done(MIG_OFFICERS_TO_RELATIONS):
                stats_off = migrate_officers_to_relations(self.items, remove_old_key=True)
                save_clients(self.items)
                mark_migration_done(MIG_OFFICERS_TO_RELATIONS, {"clients_touched": stats_off.get("clients_touched", 0)})
            else:
                # still ensure clients are saved? usually no need
                stats_off = {"clients_touched": 0, "officers_moved": 0, "officer_dupes_skipped": 0, "officers_keys_removed": 0}

            # refresh UI
            try:
                dash = getattr(self, "dashboard", None)
                if dash and hasattr(dash, "reload_from_disk"):
                    dash.reload_from_disk()
            except Exception:
                pass

            try:
                self.populate()
                self._update_suggestions()
            except Exception:
                pass

            # Pretty display for "already migrated"
            tasks_status = (
                "Already migrated" if stats_tasks.get("skipped") == -1
                else f"Updated: {stats_tasks.get('updated', 0)}\n  Skipped: {stats_tasks.get('skipped', 0)}"
            )

            messagebox.showinfo(
                "Update Data",
                "tasks.json:\n"
                f"  {tasks_status}\n\n"
                "clients.json (officers → relations):\n"
                f"  Clients touched: {stats_off.get('clients_touched', 0)}\n"
                f"  Officers moved: {stats_off.get('officers_moved', 0)}\n"
                f"  Dupes skipped: {stats_off.get('officer_dupes_skipped', 0)}\n"
                f"  officers keys removed: {stats_off.get('officers_keys_removed', 0)}\n\n"
                f"migrations flag file: {str(MIGRATIONS_FILE)}"
            )

        except Exception as e:
            messagebox.showerror("Update Data", f"Failed updating data:\n{e}")


  
    def _mgr_filter_names_ci(self):
        """Case-insensitive set for matching."""
        return { (n or "").casefold() for n in self._mgr_filter_active }

    def _toggle_mgr_filter(self, name: str | None):
        """
        name == None  -> toggle 'All'
        else          -> toggle a specific manager name
        """
        if name is None:
            if self._mgr_filter_active:
                self._mgr_filter_active.clear()
            else:
                pass
        else:
            if name in self._mgr_filter_active:
                self._mgr_filter_active.remove(name)
            else:
                self._mgr_filter_active.add(name)
        self._refresh_mgr_menu_checks()
        self.populate()

    def _refresh_mgr_menu_checks(self):
        """Refresh checkmarks in the dropdown."""
        if not self._mgr_menu:
            return

        # index 0 = "All" command
        all_selected = (len(self._mgr_filter_active) == 0)
        self._mgr_menu.entryconfig(0, label=f"{'☑' if all_selected else '☐'}  All")

        # index 1 = separator (skip relabeling!)
        # manager items start at index 2
        names = self._account_manager_names()
        start_idx = 2
        for i, nm in enumerate(names, start=start_idx):
            checked = (nm in self._mgr_filter_active)
            self._mgr_menu.entryconfig(i, label=f"{'☑' if checked else '☐'}  {nm}")


    def _open_mgr_menu(self, button_widget):
        """Build (or rebuild) and show the manager filter dropdown."""
        import tkinter as tk
        if self._mgr_menu:
            self._mgr_menu.destroy()
            self._mgr_menu = None

        m = tk.Menu(button_widget, tearoff=False)
        # All
        m.add_command(
            label=f"{'☑' if not self._mgr_filter_active else '☐'}  All",
            command=lambda: self._toggle_mgr_filter(None)
        )
        m.add_separator()
        for nm in self._account_manager_names():
            m.add_command(
                label=f"{'☑' if nm in self._mgr_filter_active else '☐'}  {nm}",
                command=lambda nm=nm: self._toggle_mgr_filter(nm)
            )
        self._mgr_menu = m
        # place under the button
        x = button_widget.winfo_rootx()
        y = button_widget.winfo_rooty() + button_widget.winfo_height()
        m.tk_popup(x, y)

    def build_link_candidates(self) -> list[dict]:
        """
        Build a list of link candidates for LinkDialog autocomplete.

        IMPORTANT:
        - Clients only (companies + individuals in clients.json). Do NOT include personnels.
        - Uses SSN for Individual, EIN for Business.
        - If missing required ID, candidate is shown but NOT linkable (id="").
        """
        cands: list[dict] = []

        def _clean(s: str) -> str:
            return (s or "").strip()

        def _is_individual_client(c: dict) -> bool:
            if bool(c.get("is_individual")):
                return True
            et = (c.get("entity_type") or "").strip().casefold()
            return et == "individual"

        def _candidate_id_and_label(i: int, c: dict) -> tuple[str, str, bool]:
            name = _clean(c.get("name", "")) or f"Client #{i}"
            is_ind = _is_individual_client(c)

            if is_ind:
                ssn9 = normalize_ssn_digits(c.get("ssn", "") or c.get("ein", ""))
                if ssn9:
                    return (f"ssn:{ssn9}", f"{name} — SSN {ssn9}", False)  # is_company=False
                return ("", f"{name} — (Missing SSN)", False)
            else:
                ein9 = normalize_ein_digits(c.get("ein", ""))
                if ein9:
                    return (f"ein:{ein9}", f"{name} — EIN {ein9}", True)   # is_company=True
                return ("", f"{name} — (Missing EIN)", True)

        for i, c in enumerate(getattr(self, "items", []) or []):
            if not isinstance(c, dict):
                continue
            cname = _clean(c.get("name", ""))
            if not cname:
                continue

            cid, label, is_company = _candidate_id_and_label(i, c)
            cands.append({
                "id": cid,             # "" means not linkable until SSN/EIN exists
                "label": label,
                "is_company": is_company,
            })

        # Dedupe by label (case-insensitive), keep first occurrence
        seen = set()
        out = []
        for x in cands:
            lab = _clean(x.get("label", ""))
            if not lab:
                continue
            k = lab.casefold()
            if k in seen:
                continue
            seen.add(k)
            out.append(x)

        return out


    def get_client_by_id(self, client_key: str):
        """Resolve a client by a key used in link candidates:
        - 'client:<id>'  -> match client['id']
        - 'ein:<ein>'    -> match client['ein'] (digits-only compare)
        - 'idx:<n>'      -> index into self.items
        - raw id         -> match client['id']
        Returns the client dict or None.
        """
        _, c = self._resolve_client_key(client_key)
        return c

    def _resolve_client_key(self, client_key: str):
        """Return (idx, client_dict) for a candidate key; (None, None) if not found.

        Supported keys:
        - client:<id>
        - ein:<9>
        - ssn:<9>
        - raw id fallback -> match client['id']
        """
        key = str(client_key or "").strip()
        if not key:
            return (None, None)

        items = getattr(self, "items", []) or []

        # client:<id>
        if key.startswith("client:"):
            target_id = key.split(":", 1)[1].strip()
            for i, c in enumerate(items):
                if isinstance(c, dict) and str(c.get("id") or "").strip() == target_id:
                    return (i, c)
            return (None, None)

        # ein:<ein>
        if key.startswith("ein:"):
            want = normalize_ein_digits(key.split(":", 1)[1])
            if not want:
                return (None, None)
            for i, c in enumerate(items):
                if not isinstance(c, dict):
                    continue
                have = normalize_ein_digits(c.get("ein", ""))
                if have and have == want:
                    return (i, c)
            return (None, None)


        # ssn:<ssn>
        if key.startswith("ssn:"):
            want = normalize_ssn_digits(key.split(":", 1)[1])
            if not want:
                return (None, None)
            for i, c in enumerate(items):
                if not isinstance(c, dict):
                    continue
                have = normalize_ssn_digits(c.get("ssn", "") or c.get("ein", ""))

                if have and have == want:
                    return (i, c)
            return (None, None)

        # raw id fallback
        for i, c in enumerate(items):
            if isinstance(c, dict) and str(c.get("id") or "").strip() == key:
                return (i, c)

        return (None, None)

    def _canonical_client_key(self, c: dict, idx: int | None):
        """Prefer stable keys for storage in linked_client_id.

        - Individual -> ssn:<9>
        - Business   -> ein:<9>
        - No idx fallback (do not persist random ids)
        """
        if not isinstance(c, dict):
            return ""

        is_individual = bool(c.get("is_individual")) or ((c.get("entity_type") or "").strip().casefold() == "individual")

        if is_individual:
            ssn9 = normalize_ssn_digits(c.get("ssn", ""))
            return f"ssn:{ssn9}" if ssn9 else ""
        else:
            ein9 = normalize_ein_digits(c.get("ein", ""))
            return f"ein:{ein9}" if ein9 else ""


    # ---------- Main page ----------
    def _build_search_page(self):
        self.page_search = ttk.Frame(self.page_host)

        topBtns = ttk.Frame(self.page_search); topBtns.pack(side=tk.TOP, fill=tk.X)
    

        # Left-aligned group
        leftGrp = ttk.Frame(topBtns); leftGrp.pack(side=tk.LEFT)
        ttk.Button(leftGrp, text="New", style="NewUI.Accent.TButton", command=self.on_new).pack(side=tk.LEFT)
        ttk.Button(leftGrp, text="Edit", style="NewUI.TButton", command=self.on_edit).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(leftGrp, text="Delete", style="NewUI.TButton", command=self.on_delete).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(leftGrp, text="Copy Emails", style="NewUI.TButton", command=self.copy_emails).pack(side=tk.LEFT, padx=(12,0))
        ttk.Button(leftGrp, text="Show Memo", style="NewUI.TButton", command=self.show_memo).pack(side=tk.LEFT, padx=(6,0))

        # Spacer expands to push the right group to the edge
        ttk.Button(leftGrp, text="Open Data Folder", style="NewUI.TButton",
                command=lambda: self.open_path(DATA_DIR)).pack(side=tk.LEFT, padx=(0,6))
        
        ttk.Frame(topBtns).pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Right-aligned group: Open Data Folder + Managers ▾
        rightGrp = ttk.Frame(topBtns); rightGrp.pack(side=tk.RIGHT)

        mgr_btn_search = ttk.Button(
            rightGrp, text="Managers ▾",
            command=lambda b=None: self._open_mgr_menu(mgr_btn_search)
        )
        mgr_btn_search.pack(side=tk.LEFT)

        style = ttk.Style()
        style.configure("Search.Treeview", rowheight=80)

        self.tree = ttk.Treeview(self.page_search, style="Search.Treeview", columns=self.COLS, show="headings", selectmode="browse")
        label_map = {"dba":"DBA", "ein":"EIN"}
        for c in self.COLS:
            header = label_map.get(c, c.replace("_"," ").title())
            self.tree.heading(c, text=header, command=lambda c=c: self.sort_by(c, False))
            if c in ("phones","ein"):
                w = 100
            elif c in ("emails","relations"):
                w = 200
            elif c == "name":
                w = 220
            else:
                w = 160
            self.tree.column(c, width=w, anchor="w")
    
        yscr = ttk.Scrollbar(self.page_search, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscr.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscr.pack(side=tk.RIGHT, fill=tk.Y)
    
        self.tree.bind("<Double-1>", self._on_main_double_click)  # reuse handler
    
        self.menu = tk.Menu(self.page_search, tearoff=False)
        self.menu.add_command(label="Show Memo", command=self.show_memo)
        self.menu.add_command(label="Open File Location", command=self.open_selected_file_location)
        self.menu.add_separator()
        self.menu.add_command(label="Delete", command=self.on_delete)
        self.tree.bind("<Button-3>", self._show_context_menu)

    # ---------- Main Dashboard ----------
    def _build_main_page(self):
        self.page_main = self.dashboard.ensure(self.page_host)


    def _ensure_main_page(self):
        if not hasattr(self, "page_main") or not self.page_main.winfo_exists():
            self.page_main = self.dashboard.ensure(self.page_host)

    # ---------- client Detail (Notebook) ----------
    def _build_detail_page(self, idx: int):
        self._clear_page_host()

        c = self.items[idx]
        page = ttk.Frame(self.page_host)
        page.pack(fill=tk.BOTH, expand=True)

        # ------------- HEADER ----------------
        GAP = 28  # equal horizontal spacing between EIN / Entity / City-State
        
        header = ttk.Frame(page, style="Card.TFrame", padding=(8, 8, 8, 4))
        header.pack(side=tk.TOP, fill=tk.X)
        
        # Left stack: Name (big) + DBA (just below)
        name_lbl = ttk.Label(header, text=c.get("name", ""), style="Header.TLabel", font=("Segoe UI", 18, "bold"))     
        dba = c.get("dba", "")
        if dba:
            name_lbl.grid(row=0, column=0, sticky="w")
            dba_lbl = ttk.Label(header, text=f"DBA: {dba}", font=("Segoe UI", 11, "italic"), style="Subtle.TLabel")
            dba_lbl.grid(row=1, column=0, sticky="w")
        else:
            name_lbl.grid(row=0, column=0, rowspan=2, sticky="w")
        
        # Right row: EIN · Entity · City,State  (equal gaps)
        info = ttk.Frame(header)
        info.grid(row=0, column=1, rowspan=2, sticky="w", padx=(GAP, 0))
        
        ein_text   = f"EIN: {c.get('ein','') or '—'}"
        ent_text   = f"Entity: {c.get('entity_type','') or '—'}"
        city_state = ", ".join(p for p in [c.get("city",""), c.get("state","")] if p) or "—"
        mgr_text   = f"Managed by: {c.get('acct_mgr','') or '—'}"
        
        ttk.Label(info, text=ein_text,   font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(info, text=ent_text,   font=("Segoe UI", 12, "bold")).grid(row=0, column=1, sticky="w", padx=(GAP, 0))
        ttk.Label(info, text=city_state, font=("Segoe UI", 12, "bold")).grid(row=0, column=2, sticky="w", padx=(GAP, 0))
        ttk.Label(info, text=mgr_text,   font=("Segoe UI", 12, "bold")).grid(row=0, column=3, sticky="w", padx=(GAP, 0))

        # Actions on the far right of the header
        actions = ttk.Frame(header)
        actions.grid(row=0, column=2, rowspan=2, sticky="e", padx=(16, 0))
        
        ttk.Button(actions, text="Edit Client", style="NewUI.Accent.TButton",
                   command=lambda i=idx: self._detail_edit(i)).pack(side="top", fill="x")
        ttk.Button(actions, text="Delete Client", style="NewUI.Accent.TButton",
                   command=lambda i=idx: self._delete_client(i)).pack(side="top", fill="x", pady=(6,0))
        ttk.Button(actions, text="Open File Client", style="NewUI.Accent.TButton",
                   command=lambda: self.open_path(c.get("file_location",""))).pack(side="top", fill="x", pady=(6,0))
        
        # Column sizing so name (col 0) hugs content, middle info (col 1) expands, actions (col 2) stays tight.
        header.grid_columnconfigure(0, weight=0)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=0)

        page.bind_all("<Control-e>", lambda e=None, i=idx: self._detail_edit(i))


        nb = ttk.Notebook(page); nb.pack(fill="both", expand=True, padx=6, pady=6)
        self._detail_notebook = nb

        # Profile tab
        init_profile_tab(
            nb,
            self,
            c,
            edit_rates_cb=lambda i=idx: self._edit_rates(i),
            refresh_sales_cb=lambda i=idx: self._refresh_sales_tax_for(i),
        )

        # Documents tab (add to the SAME notebook that already has "Profile")
        def _save_clients(items):
            save_clients(items)

        try:
            init_documents_tab(nb, self, c, _save_clients)
        except Exception as ex:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Documents Tab Error", f"{type(ex).__name__}: {ex}")

        init_logs_tab(nb, self, c, _save_clients)

        # Analysis tab
        ana = ttk.Frame(nb, padding=8); nb.add(ana, text="Analysis")
        ttk.Label(ana, text=f"relations: {len(c.get('relations',[]))}").pack(anchor="w")
        ttk.Label(ana, text=f"Employees: {len(c.get('employees',[]))}").pack(anchor="w")
        ttk.Label(ana, text=f"Has EIN: {'Yes' if c.get('ein') else 'No'}").pack(anchor="w")
        ttk.Label(ana, text=f"Entity Type: {c.get('entity_type','') or '—'}").pack(anchor="w")

        footer = ttk.Frame(page, padding=(8,2)); footer.pack(side=tk.BOTTOM, fill=tk.X)

        self._current_detail_idx = idx
        self._current_page = ("detail", idx)

    # ---------- Personnel Detail Page ----------
    def _build_person_page(self, client_idx: int, role_key: str, person_idx: int):
        self._clear_page_host()

        c = self.items[client_idx]
        people = c.get(role_key, [])
        if person_idx < 0 or person_idx >= len(people):
            self.navigate("detail", client_idx, replace=True); return
        p = ensure_relation_dict(people[person_idx])

        page = ttk.Frame(self.page_host)
        page.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(page, padding=(8,8,8,4)); header.pack(side=tk.TOP, fill=tk.X)

        title_box = ttk.Frame(header); title_box.pack(side=tk.LEFT, padx=12, fill=tk.X, expand=True)
        ttk.Label(title_box, text=p.get("name","") or "(Person)", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(title_box, text=f"{'relations' if role_key=='relations' else 'Employee'} — {c.get('name','')}", foreground="#555").pack(anchor="w", pady=(2,0))

        actions = ttk.Frame(header); actions.pack(side=tk.RIGHT)
        ttk.Button(actions, text="Edit", command=lambda: self._edit_person(client_idx, role_key, person_idx)).pack(side=tk.TOP, fill=tk.X)
        ttk.Button(actions, text="Back to client", command=lambda: self.navigate('detail', client_idx, push=True)).pack(side=tk.TOP, fill=tk.X, pady=(6,0))

        nb = ttk.Notebook(page); nb.pack(fill="both", expand=True, padx=6, pady=6)
        self._detail_notebook = nb

        # Profile tab
        prof = ttk.Frame(nb, padding=10); nb.add(prof, text="Profile")
        def row(lbl, val):
            f = ttk.Frame(prof); f.pack(anchor="w", pady=2, fill="x")
            ttk.Label(f, text=f"{lbl}:", width=18).pack(side="left")
            ttk.Label(f, text=val or "—", wraplength=600).pack(side="left", fill="x", expand=True)

        row("First Name", p.get("first_name"))
        row("Middle Name", p.get("middle_name"))
        row("Last Name", p.get("last_name"))
        row("Nickname", p.get("nickname"))
        row("Email", p.get("email"))
        row("Phone", p.get("phone"))
        row("Address 1", p.get("addr1"))
        row("Address 2", p.get("addr2"))
        row("City", p.get("city"))
        row("State", p.get("state"))
        row("Zip", p.get("zip"))
        row("DOB", p.get("dob"))

        # Documents tab (placeholder)
        docs = ttk.Frame(nb, padding=10); nb.add(docs, text="Documents")
        ttk.Label(docs, text="Attach or link personnel documents here (not implemented).").pack(anchor="w")

        # Logs tab (placeholder)
        logs = ttk.Frame(nb, padding=10); nb.add(logs, text="Logs")
        ttk.Label(logs, text="Personnel logs go here (not implemented).").pack(anchor="w")

        footer = ttk.Frame(page, padding=(8,2)); footer.pack(side=tk.BOTTOM, fill=tk.X)

        self._current_page = ("person", (client_idx, role_key, person_idx))

    def _edit_person(self, client_idx: int, role_key: str, person_idx: int):
        c = self.items[client_idx]
        p = ensure_relation_dict(c.get(role_key, [])[person_idx])

        prev_link = (p.get("linked_client_id") or "").strip()

        link_cands = []
        try:
            link_cands = self.build_link_candidates()
        except Exception:
            link_cands = []

        d = LinkDialog(self, f"Edit {'relations' if role_key=='relations' else 'Employee'}", p, link_candidates=link_cands)
        self.wait_window(d)
        if d.result:
            newp = ensure_relation_dict(d.result)
            if not (newp.get("linked_client_id") or "").strip():
                newp["linked_client_label"] = ""
            c[role_key][person_idx] = newp

            # Persist this side first
            save_clients(self.items)

            # Sync other side if link changed
            new_link = (newp.get("linked_client_id") or "").strip()
            if prev_link != new_link:
                self._sync_link_change_for_person(client_idx, prev_link, new_link)

            self.navigate("person", (client_idx, role_key, person_idx), replace=True)



    # ---- Rates editing / refresh ----
    def _edit_rates(self, idx: int):
        c = self.items[idx]
        dlg = tk.Toplevel(self); dlg.title("Edit Tax Rates"); dlg.resizable(False, False)
        v_ui   = tk.StringVar(value=c.get("ui_rate",""))
        v_stx  = tk.StringVar(value=c.get("sales_tax_rate",""))
        v_oth  = tk.StringVar(value=c.get("other_tax_rates",""))
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)
        def row(lbl, var, width=22):
            f = ttk.Frame(frm); f.pack(fill="x", pady=3)
            ttk.Label(f, text=lbl, width=18).pack(side="left")
            ttk.Entry(f, textvariable=var, width=width).pack(side="left")
        row("UI Rate (%)", v_ui)
        row("Sales Tax Rate (%)", v_stx)
        row("Other (text)", v_oth, width=40)
        btns = ttk.Frame(frm); btns.pack(anchor="e", pady=(8,0))
        def save():
            c["ui_rate"] = v_ui.get().strip()
            c["sales_tax_rate"] = v_stx.get().strip()
            c["other_tax_rates"] = v_oth.get().strip()
            c["tax_rates_last_checked"] = today_date().isoformat()
            save_clients(self.items)
            self.navigate("detail", idx, replace=True)
            dlg.destroy()
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=(6,0))
        ttk.Button(btns, text="Save", command=save).pack(side="right")
        dlg.grab_set(); dlg.transient(self)

    def _refresh_sales_tax_for(self, idx: int):
        c = self.items[idx]
        rate = safe_fetch_sales_tax_rate(c.get("state",""), c.get("city",""))
        if rate is not None:
            c["sales_tax_rate"] = f"{rate}"
            c["tax_rates_last_checked"] = today_date().isoformat()
            save_clients(self.items)
            messagebox.showinfo("Sales Tax", f"Sales tax rate updated to {rate}%")
            self.navigate("detail", idx, replace=True)
        else:
            messagebox.showinfo("Sales Tax", "Could not fetch sales tax rate (offline or API not configured).")

    def update_sales_tax_rates_if_due(self):
        changed = False
        for c in self.items:
            if new_quarter_started(c.get("tax_rates_last_checked")):
                rate = safe_fetch_sales_tax_rate(c.get("state",""), c.get("city",""))
                if rate is not None:
                    c["sales_tax_rate"] = f"{rate}"
                    c["tax_rates_last_checked"] = today_date().isoformat()
                    changed = True
        if changed:
            save_clients(self.items)

    # ---- Actions page integration ----
    def open_actions_page(self, tool_key: str | None = None):
        """
        Navigate to the in-app Actions page. Optionally preselect a tool:
        tool_key in {"pos_parse", "validate_tax", "vendor_match"}
        """
        # stash a small payload; navigate() will build the page and apply it
        self.navigate("actions", payload=tool_key, push=True)

    def _ensure_actions_page(self):
        if not hasattr(self, "_actions_page"):
            self._actions_page = ActionRunnerPage(self)
            self.page_actions = self._actions_page.ensure(self.page_host)
        else:
            # Make sure it is attached to the current host
            self.page_actions = self._actions_page.ensure(self.page_host)

    # --- Data operations used by Taskbar / File menu -----------------
    def _save_all_data(self):
        """Flush current in-memory data to internal storage (clients.json, later tasks/rules)."""
        save_clients(self.items)
        if hasattr(self, "status"):
            self.status.set(f"Saved {len(self.items)} client(s).")

    def _import_data_dialog(self):
        path_str = filedialog.askopenfilename(
            title="Import Data (JSON)",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path_str:
            return

        path = Path(path_str)
        self.log.info("Importing data from %s", path)

        stats = import_all_from_json(path, self.items)

        # Reload account managers immediately so dropdowns reflect imports
        try:
            self.account_managers = self._load_account_managers()
        except Exception:
            pass

        # Refresh dashboard tasks immediately (no restart needed)
        try:
            if hasattr(self, "dashboard") and self.dashboard:
                self.dashboard.reload_from_disk()
        except Exception:
            pass

        # always persist clients list if any were added
        if stats.get("clients_added", 0):
            save_clients(self.items)

        # refresh UI (clients + suggestions). Tasks pages read from TASKS_FILE, but remap above makes them consistent.
        self.populate()
        self._update_suggestions()

        if hasattr(self, "status"):
            self.status.set(
                f"Imported from {path.name} | "
                f"clients +{stats.get('clients_added',0)}, "
                f"AM +{stats.get('account_managers_added',0)}, "
                f"tasks +{stats.get('tasks_added',0)}, "
                f"rules +{stats.get('match_rules_added',0)}, "
                f"vendor_lists {stats.get('vendor_lists_written',0)}"
            )

    def _export_selected_dialog(self):
        win = tk.Toplevel(self)
        win.title("Export Data")
        win.transient(self)
        win.grab_set()

        # --- state ---
        var_all = tk.BooleanVar(value=True)
        var_clients = tk.BooleanVar(value=True)
        var_include_logs = tk.BooleanVar(value=True)
        var_rules = tk.BooleanVar(value=True)
        var_monthly = tk.BooleanVar(value=True)
        var_tasks = tk.BooleanVar(value=True)
        var_vendors = tk.BooleanVar(value=True)

        folder_var = tk.StringVar(value="")

        def set_children(state: bool):
            var_clients.set(state)
            var_rules.set(state)
            var_monthly.set(state)
            var_tasks.set(state)
            var_vendors.set(state)

        def on_all_toggle():
            set_children(var_all.get())

        def on_child_toggle():
            # If any child unchecked -> All becomes false. If all checked -> All true.
            all_on = all([
                var_clients.get(),
                var_rules.get(),
                var_monthly.get(),
                var_tasks.get(),
                var_vendors.get(),
            ])
            var_all.set(all_on)

            if var_clients.get():
                cb_logs.config(state="normal")
            else:
                cb_logs.config(state="disabled")
                var_include_logs.set(False)
        def choose_folder():
            folder = filedialog.askdirectory(title="Choose export folder")
            if folder:
                folder_var.set(folder)

        # --- UI ---
        frm = tk.Frame(win, padx=12, pady=12)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Choose what to export:").pack(anchor="w")

        tk.Checkbutton(frm, text="All", variable=var_all, command=on_all_toggle).pack(anchor="w", pady=(6, 0))

        tk.Checkbutton(frm, text="Clients (includes Account Managers)", variable=var_clients, command=on_child_toggle).pack(anchor="w")
        cb_logs = tk.Checkbutton(frm, text="↳ Include Client Logs", variable=var_include_logs, command=on_child_toggle)
        cb_logs.pack(anchor="w", padx=24)
        tk.Checkbutton(frm, text="Match Rules", variable=var_rules, command=on_child_toggle).pack(anchor="w")
        tk.Checkbutton(frm, text="Monthly Data", variable=var_monthly, command=on_child_toggle).pack(anchor="w")
        tk.Checkbutton(frm, text="Tasks", variable=var_tasks, command=on_child_toggle).pack(anchor="w")
        tk.Checkbutton(frm, text="Vendor Lists (CSV)", variable=var_vendors, command=on_child_toggle).pack(anchor="w")

        sep = tk.Frame(frm, height=10)
        sep.pack()

        row = tk.Frame(frm)
        row.pack(fill="x")
        tk.Button(row, text="Choose Folder…", command=choose_folder).pack(side="left")
        tk.Label(row, textvariable=folder_var).pack(side="left", padx=8)

        btns = tk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))

        def do_export():
            folder = folder_var.get().strip()
            if not folder:
                messagebox.showwarning("Export", "Please choose an export folder.")
                return

            selections = {
                "clients": var_clients.get(),
                "include_logs": var_include_logs.get(),
                "match_rules": var_rules.get(),
                "monthly_data": var_monthly.get(),
                "tasks": var_tasks.get(),
                "vendor_lists": var_vendors.get(),
            }
            if not any(selections.values()):
                messagebox.showwarning("Export", "Please select at least one item to export.")
                return

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = Path(folder) / f"Vertex_Export_{ts}.json"

            try:
                export_selected_to_json(out_path, self.items, selections)
                messagebox.showinfo("Export", f"Exported successfully to:\n{out_path}")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Export Failed", f"{e}")

        tk.Button(btns, text="Cancel", command=win.destroy).pack(side="right")
        tk.Button(btns, text="Export", command=do_export).pack(side="right", padx=(0, 8))

        # default: All checked
        on_all_toggle()

    # ---------- Misc ----------
    def _show_context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.menu.tk_popup(event.x_root, event.y_root)

    def _set_row_height_for_items(self, items: List[Dict[str, Any]]):
        max_lines = 1
        for c in items:
            relations_lines = max(1, len(relations_to_display_lines(c.get("relations", []))))
            emails_lines   = max(1, len(relations_to_flat_emails(c.get("relations", []))) or 1)
            phones_lines   = max(1, len(relations_to_flat_phones(c.get("relations",[]))) or 1)
            max_lines = max(max_lines, relations_lines, emails_lines, phones_lines)
        row_px = max(self.base_row_px, int(self.base_row_px * 1.45) * max_lines)
        self.style.configure("Treeview", rowheight=row_px)

    def _memo_preview(self, text: str) -> str:
        text = (text or "").replace("\n", " ").strip()
        if not text: return ""
        memo_w = 220
        avg_px = max(6, self.default_font.measure("M"))
        cap = max(12, int((memo_w - 20) / avg_px))
        return (text[:cap-1] + "…") if len(text) > cap else text

    def _build_logs_panel(self, parent, logs_list, on_change):
        """
        parent: a Frame to build into
        logs_list: list[dict] (ts,user,text)
        on_change: callback(new_logs_list) to persist from page
        """
        frame = ttk.Frame(parent); frame.pack(fill=tk.BOTH, expand=True)
    
        cols = ("ts", "user", "text")
        tv = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        tv.heading("ts",   text="Timestamp")
        tv.heading("user", text="User")
        tv.heading("text", text="Text")
        tv.column("ts", width=160, anchor="w")
        tv.column("user", width=80, anchor="w")
        tv.column("text", width=600, anchor="w")
        y = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=y.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y.pack(side=tk.RIGHT, fill=tk.Y)
    
        def refresh_tv():
            tv.delete(*tv.get_children())
            for entry in logs_list:
                tv.insert("", "end", values=(entry.get("ts",""), entry.get("user",""), entry.get("text","")))
    
        def add_log():
            d = LogDialog(self.winfo_toplevel(), "Add Log")
            self.wait_window(d)
            if d.result:
                logs_list.append(d.result)
                on_change(logs_list)
                refresh_tv()
    
        def edit_log():
            sel = tv.selection()
            if not sel:
                messagebox.showinfo("Edit Log", "Select a log row to edit."); return
            v = tv.item(sel[0], "values")
            init = {"ts": v[0], "user": v[1], "text": v[2]}
            d = LogDialog(self.winfo_toplevel(), "Edit Log", init)
            self.wait_window(d)
            if d.result:
                idx = tv.index(sel[0])
                logs_list[idx] = d.result
                on_change(logs_list)
                refresh_tv()
    
        def delete_log():
            sel = tv.selection()
            if not sel:
                messagebox.showinfo("Delete Log", "Select a log row to delete."); return
            if not messagebox.askyesno("Confirm Delete", "Delete selected log?"):
                return
            idx = tv.index(sel[0])
            del logs_list[idx]
            on_change(logs_list)
            refresh_tv()
    
        # context & double-click
        menu = tk.Menu(frame, tearoff=False)
        menu.add_command(label="Edit", command=edit_log)
        menu.add_command(label="Delete", command=delete_log)
        tv.bind("<Button-3>", lambda e: (tv.identify_row(e.y) and tv.selection_set(tv.identify_row(e.y)), menu.tk_popup(e.x_root, e.y_root)))
        tv.bind("<Double-1>", lambda _e: edit_log())
    
        btns = ttk.Frame(parent); btns.pack(fill="x", pady=(6,0))
        ttk.Button(btns, text="Add", command=add_log).pack(side=tk.LEFT)
        ttk.Button(btns, text="Edit", command=edit_log).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(btns, text="Delete", command=delete_log).pack(side=tk.LEFT, padx=(6,0))
    
        refresh_tv()
        return frame

    def filtered_items(self) -> List[Dict[str, Any]]:
        q = self.q.get().strip()
        q_raw = q.casefold()
        raw_tokens = q_raw.split()

        if not q:
            base = list(self.items)
        else:
            q_norm = norm_text(q)
            q_tokens = q_norm.split()
            is_digits_only = q.isdigit()
            q_digits = "".join(PHONE_DIGITS_RE.findall(q))
            last10 = q_digits[-10:] if len(q_digits) >= 4 else q_digits
            last9  = q_digits[-9:]  if len(q_digits) >= 4 else q_digits

            res = []
            for c in self.items:
                relations_phones = relations_to_flat_phones(c.get("relations",[]))
                phones_digits_full = ["".join(PHONE_DIGITS_RE.findall(p or "")) for p in relations_phones]
                phones_norm_last10 = [normalize_phone_digits(p) for p in relations_phones]
                phone_hit = False
                if last10:
                    if any(last10 in p for p in phones_digits_full) or any(p.endswith(last10) for p in phones_norm_last10 if p):
                        phone_hit = True

                ein_digits = normalize_ein_digits(c.get("ein",""))
                ein_hit = bool(last9 and ein_digits and ein_digits.endswith(last9))

                hay = " ".join([
                    norm_text(c.get("name","")),
                    norm_text(c.get("dba","")),
                    norm_text(c.get("entity_type","")),
                    norm_text(c.get("acct_mgr","")),
                    norm_text(c.get("edd_number","")),
                    norm_text(c.get("sales_tax_account","")),
                    norm_text(c.get("addr1","")), norm_text(c.get("addr2","")),
                    norm_text(c.get("city","")), norm_text(c.get("state","")), norm_text(c.get("zip","")),
                    " ".join(norm_text(ensure_relation_dict(o).get("name","")) for o in c.get("relations",[])),
                    " ".join(norm_text(ensure_relation_dict(o).get("first_name","")) for o in c.get("relations",[])),
                    " ".join(norm_text(ensure_relation_dict(o).get("last_name","")) for o in c.get("relations",[])),
                    " ".join(norm_text(ensure_relation_dict(o).get("nickname","")) for o in c.get("relations",[])),
                    " ".join(norm_text(ensure_relation_dict(o).get("email","")) for o in c.get("relations",[])),
                    norm_text(c.get("file_location","")),
                    norm_text(c.get("memo","")),
                ])


                text_hit = all(tok in hay for tok in q_tokens) if q_norm else False
                if is_digits_only:
                    if phone_hit or ein_hit or text_hit:
                        res.append(c)
                else:
                    if text_hit:
                        res.append(c)
            base = res

        # Manager filter (single/multi). Empty set = All managers.
        if self._mgr_filter_active:
            allowed_ci = {(n or "").casefold() for n in self._mgr_filter_active}
            base = [c for c in base if (c.get("acct_mgr","") or "").casefold() in allowed_ci]
        return base


    def populate(self):
        if not hasattr(self, "tree") or not self.tree.winfo_exists(): return
        self.tree.delete(*self.tree.get_children())
        items = sorted(self.filtered_items(), key=lambda c: (c.get("name","").lower()))
        self._set_row_height_for_items(items)
        for c in items:
            relations_lines = relations_to_display_lines(c.get("relations",[]))
            emails_lines  = relations_to_flat_emails(c.get("relations",[]))
            phones_lines  = relations_to_flat_phones(c.get("relations",[]))
            vals = (
                c.get("name",""),
                c.get("dba",""),
                "\n".join(relations_lines),
                c.get("ein",""),
                "\n".join(emails_lines),
                "\n".join(phones_lines),
                self._memo_preview(c.get("memo","")),
            )
            self.tree.insert("", "end", values=vals)
        filt = ", ".join(sorted(self._mgr_filter_active)) or "All managers"
        self.status.set(f"Showing {len(items)} / {len(self.items)} clients — {filt}. Data: {DATA_FILE}")

        NewUI.stripe_tree(self.tree)
        try:
            self._focus_last_in_search()   # <-- move/add here (END of populate)
        except Exception:
            pass
    
    def _focus_client_in_search(self, key: str):
        if not key or not hasattr(self, "tree") or not self.tree.winfo_exists():
            return
        key = (key or "").strip().lower()

        # search rows
        for iid in self.tree.get_children(""):
            vals = self.tree.item(iid, "values") or ()
            name = (vals[0] if len(vals) > 0 else "") or ""
            ein  = (vals[3] if len(vals) > 3 else "") or ""
            if key in ((name or "").strip().lower(), (ein or "").strip().lower()):
                try:
                    self.tree.selection_set(iid)
                    self.tree.focus(iid)
                    self.tree.see(iid)
                except Exception:
                    pass
                break

    def _update_suggestions(self):
        q = self.q.get().strip()
        if not q:
            self._ac.hide(); return
        
        if self.focus_get() is not self.search_entry:
            self._ac.hide()
            return
    
        matches = self.filtered_items()
        lines = [f"{c.get('name','')} — {c.get('dba','') or 'No DBA'} — {c.get('ein','') or 'No EIN'}" for c in matches[:20]]
        self._ac.show(lines)

    def _open_from_suggestion(self, text: str):
        if not text: 
            self._ac.hide()
            return
        name = text.split("—", 1)[0].strip()
        idx = next((i for i, c in enumerate(self.items) if (c.get("name","").strip() == name)), None)
        if idx is not None:
            try:
                c = self.items[idx]
                self._return_focus_key = (c.get("ein") or c.get("name") or "").strip()
            except Exception:
                self._return_focus_key = None
            self._ac.hide()
            self.navigate("detail", idx, push=True)
        else:
            self.navigate("search", None, push=True)

    def clear_search(self):
        self.q.set("")

    def _clear_page_host(self):
        for w in self.page_host.winfo_children():
            w.destroy()

    def refresh(self):
        self.items = load_clients()
        self.populate()
        self._update_suggestions()

    # ---------- CRUD ----------
    def _find_index_by_row_values(self, row_vals):
        name, dba, ein = row_vals[0], row_vals[1], row_vals[3]
        for i, c in enumerate(self.items):
            if (c.get("name","") == name and c.get("dba","") == dba and str(c.get("ein","")) == str(ein)):
                return i
        memo_snip = row_vals[6] if len(row_vals) > 6 else ""
        for i, c in enumerate(self.items):
            if c.get("name","") == name and (c.get("memo","") or "").startswith(memo_snip.rstrip("…")):
                return i
        return None

    def selected_index(self):
        sel = self.tree.selection()
        if not sel: return None
        row = self.tree.item(sel[0], "values")
        return self._find_index_by_row_values(row)

    def _dup_id_conflict(self, candidate: dict, ignore_idx: int | None = None):
        """
        Returns (kind, digits, other_idx, other_client) if candidate conflicts with an existing client.
        kind is "SSN" or "EIN".
        """
        if not isinstance(candidate, dict):
            return None

        is_individual = bool(candidate.get("is_individual")) or ((candidate.get("entity_type") or "").strip().casefold() == "individual")

        cand_ssn = normalize_ssn_digits(candidate.get("ssn", "") or candidate.get("ein", ""))
        cand_ein = normalize_ein_digits(candidate.get("ein", ""))

        # If individual, enforce SSN uniqueness when present.
        # If business, enforce EIN uniqueness when present.
        for i, c in enumerate(getattr(self, "items", []) or []):
            if ignore_idx is not None and i == ignore_idx:
                continue
            if not isinstance(c, dict):
                continue

            exist_ssn = normalize_ssn_digits(c.get("ssn", "") or c.get("ein", ""))
            exist_ein = normalize_ein_digits(c.get("ein", ""))

            if cand_ssn and exist_ssn and cand_ssn == exist_ssn:
                return ("SSN", cand_ssn, i, c)
            if cand_ein and exist_ein and cand_ein == exist_ein:
                return ("EIN", cand_ein, i, c)

        return None

    def on_new(self):
        dlg = ClientDialog(self, "New Client"); self.wait_window(dlg)
        if dlg.result:
            conflict = self._dup_id_conflict(dlg.result, ignore_idx=None)
            if conflict:
                kind, digits, other_idx, other_client = conflict
                other_name = (other_client.get("name") or "").strip() or f"Client #{other_idx}"
                messagebox.showerror(
                    "Duplicate ID",
                    f"A client with the same {kind} already exists.\n\n"
                    f"{kind}: {digits}\n"
                    f"Existing client: {other_name}\n\n"
                    f"This client was NOT added."
                )
                return

            self.items.append(dlg.result)
            save_clients(self.items)
            self.populate()
            self._update_suggestions()


    def on_edit(self):
        idx = getattr(self, "_current_detail_idx", None)
        if self._current_page[0] == "main" or idx is None:
            idx = self.selected_index()
            if idx is None:
                messagebox.showinfo("Edit", "Select a row to edit."); return
        dlg = ClientDialog(self, "Edit Client", self.items[idx]); self.wait_window(dlg)
        if dlg.result:
            conflict = self._dup_id_conflict(dlg.result, ignore_idx=idx)
            if conflict:
                kind, digits, other_idx, other_client = conflict
                other_name = (other_client.get("name") or "").strip() or f"Client #{other_idx}"
                messagebox.showerror(
                    "Duplicate ID",
                    f"Another client already has this {kind}.\n\n"
                    f"{kind}: {digits}\n"
                    f"Existing client: {other_name}\n\n"
                    f"Your changes were NOT saved."
                )
                return

            self.items[idx] = dlg.result
            save_clients(self.items)
            self.populate()
            self._update_suggestions()
            if self._current_page[0] != "main":
                self.navigate(self._current_page[0], self._current_page[1], replace=True)


    def _detail_edit(self, idx: int):
        self.on_edit()

    def on_delete(self):
        self._delete_client()

    # ---------- Actions ----------
    def copy_emails(self):
        items = self.filtered_items()
        seen, acc = set(), []
        for c in items:
            for e in relations_to_flat_emails(c.get("relations",[])):
                if e and e not in seen:
                    seen.add(e); acc.append(e)
        if not acc:
            self.status.set("No emails found for current filter."); return
        if len(acc) == 1:
            self.clipboard_clear(); self.clipboard_append(acc[0])
            self.status.set("Copied 1 email to clipboard."); return

        win = tk.Toplevel(self.winfo_toplevel()); win.title("Choose email(s) to copy")
        ttk.Label(win, text=f"{len(acc)} emails found. Select and click Copy.").pack(anchor="w", padx=10, pady=(10,4))
        lb = tk.Listbox(win, selectmode=tk.EXTENDED, height=min(12, len(acc)), width=56)
        for e in acc: lb.insert(tk.END, e)
        lb.pack(fill="both", expand=True, padx=10, pady=(0,8))
        btns = ttk.Frame(win); btns.pack(fill="x", padx=10, pady=(0,10))
        def do_copy_selected():
            sel = [lb.get(i) for i in lb.curselection()]
            if not sel:
                messagebox.showinfo("Copy Emails", "Select one or more emails."); return
            txt = "; ".join(sel)
            self.clipboard_clear(); self.clipboard_append(txt)
            self.status.set(f"Copied {len(sel)} email(s) to clipboard."); win.destroy()
        def do_copy_all():
            txt = "; ".join(acc)
            self.clipboard_clear(); self.clipboard_append(txt)
            self.status.set(f"Copied all {len(acc)} emails to clipboard."); win.destroy()
        ttk.Button(btns, text="Copy Selected", command=do_copy_selected).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Copy All", command=do_copy_all).pack(side=tk.RIGHT, padx=(0,6))
        win.grab_set()

    def show_memo(self):
        i = self.selected_index()
        if i is None:
            messagebox.showinfo("Show Memo", "Select a row first."); return
        memo = self.items[i].get("memo","")
        win = tk.Toplevel(self.winfo_toplevel()); win.title("Memo")
        txt = ScrolledText(win, width=72, height=16, wrap="word"); txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", memo); txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0,10)); win.grab_set()

    def open_selected_file_location(self):
        i = self.selected_index()
        if i is None: return
        path = str(self.items[i].get("file_location","") or "").strip()
        if not path:
            messagebox.showinfo("Open File Location", "No file location set for this client."); return
        self.open_path(path)

    def open_path(self, path: str | Path):
        path = str(path)
        try:
            if re.match(r"^https?://", path):
                webbrowser.open(path); return
            if os.path.isdir(path):
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", path])
                else:
                    subprocess.run(["xdg-open", path])
            elif os.path.exists(path):
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", path])
                else:
                    subprocess.run(["xdg-open", path])
            else:
                messagebox.showwarning("Not found", f"Path not found:\n{path}")
        except Exception as e:
            LOG.exception("Open path failed: %s", path)
            messagebox.showerror("Open Error", str(e))

    def _delete_client(self, idx: int | None = None):
        if idx is None:
            if getattr(self, "_current_page", ("main", None))[0] != "main":
                idx = getattr(self, "_current_detail_idx", None)
            if idx is None:
                idx = self.selected_index()

        if idx is None or idx < 0 or idx >= len(self.items):
            messagebox.showinfo("Delete", "Select a client to delete first.")
            return

        c = self.items[idx]
        if not messagebox.askyesno(
            "Confirm Delete",
            f"Delete this client?\n\n{c.get('name','')}  (DBA: {c.get('dba','') or '—'})"
        ):
            return

        # Optional safety: if you want to block deleting clients without stable IDs
        # target_id = self._canonical_client_key(c, idx)
        # if not target_id:
        #     messagebox.showerror("Delete Error", "This entity has no ID (EIN/SSN). Cannot safely remove links.")
        #     return

        # Delete exactly once
        del self.items[idx]

        # Persist
        self.save_clients_data()

        # After deletion, return to search and refresh
        self.navigate("search", None, push=False, replace=True)
        self.populate()
        self._update_suggestions()
        self.status.set("Client deleted.")

    def _on_delete_from_tree(self, event):
        tv = event.widget
        if not isinstance(tv, ttk.Treeview):
            return
        sel = tv.selection()
        if not sel:
            return

        iid = sel[0]
        self._delete_client_at_index(iid, tv)


    def _delete_client_at_index(self, iid, tv):
        """
        Delete the client corresponding to a specific Treeview item id (iid).
        IMPORTANT: Treeview row order != self.items order when sorted/filtered.
        """
        try:
            row_vals = tv.item(iid, "values") or ()
        except Exception:
            row_vals = ()

        if not row_vals:
            return

        real_idx = self._find_index_by_row_values(row_vals)
        if real_idx is None:
            messagebox.showwarning("Delete", "Could not resolve the selected client in data.")
            return

        self._delete_client(real_idx)


    def _client_label(self, idx: int) -> str:
        try:
            c = self.items[idx]
            return (c.get("name") or "").strip() or f"Client #{idx}"
        except Exception:
            return f"Client #{idx}"

    def _client_link_id(self, idx: int) -> str:
        """
        Stable identifier used for linked_client_id.

        Rules:
        - Individual -> ssn:<9>
        - Business   -> ein:<9>
        - If missing required ID, returns "" (caller should block linking)
        """
        c = self.items[idx]

        is_individual = bool(c.get("is_individual")) or ((c.get("entity_type") or "").strip().casefold() == "individual")

        if is_individual:
            ssn9 = normalize_ssn_digits(c.get("ssn", ""))
            return f"ssn:{ssn9}" if ssn9 else ""
        else:
            ein9 = normalize_ein_digits(c.get("ein", ""))
            return f"ein:{ein9}" if ein9 else ""

    def _ensure_people_list(self, client_idx: int, role_key: str) -> list:
        c = self.items[client_idx]
        c.setdefault(role_key, [])
        if not isinstance(c[role_key], list):
            c[role_key] = []
        return c[role_key]

    def _is_link_record_to(self, person_dict: dict, target_link_id: str) -> bool:
        try:
            return str(person_dict.get("linked_client_id") or "").strip() == str(target_link_id or "").strip()
        except Exception:
            return False

    def _remove_link_record(self, client_idx: int, role_key: str, target_link_id: str) -> bool:
        """
        Remove any personnel entry whose linked_client_id == target_link_id.
        Returns True if anything removed.
        """
        people = self._ensure_people_list(client_idx, role_key)
        before = len(people)
        people[:] = [ensure_relation_dict(x) for x in people if not self._is_link_record_to(ensure_relation_dict(x), target_link_id)]
        return len(people) != before

    def _upsert_link_record(self, client_idx: int, role_key: str, record: dict) -> bool:
        """
        Insert or replace a link-record in role_key by matching linked_client_id.
        Returns True if changed.
        """
        people = self._ensure_people_list(client_idx, role_key)
        record = ensure_relation_dict(record)
        tgt = str(record.get("linked_client_id") or "").strip()
        if not tgt:
            return False

        for i, x in enumerate(people):
            x = ensure_relation_dict(x)
            if str(x.get("linked_client_id") or "").strip() == tgt:
                # Replace existing
                people[i] = record
                return True

        people.append(record)
        return True

    def _build_link_record_from_client(self, source_client_idx: int, target_client_idx: int) -> dict:
        """
        Create the personnel dict that will be stored on source, pointing to target.
        Pulls fields from target client (address + contact), and best-effort individual name parts.
        """
        target = self.items[target_client_idx]
        target_label = self._client_label(target_client_idx)
        target_link_id = self._client_link_id(target_client_idx)

        is_individual = bool(target.get("is_individual")) or ((target.get("entity_type") or "").strip().casefold() == "individual")

        # Prefer target's own email/phone if present
        email = (target.get("email") or "").strip() if isinstance(target.get("email"), str) else ""
        phone = (target.get("phone") or "").strip() if isinstance(target.get("phone"), str) else ""

        # Fallback: use the target's first relations contact info
        if (not email) or (not phone):
            try:
                offs = target.get("relations", []) or []
                if offs:
                    o0 = ensure_relation_dict(offs[0])
                    if not email:
                        email = (o0.get("email") or "").strip()
                    if not phone:
                        phone = (o0.get("phone") or "").strip()
            except Exception:
                pass

        first_name = ""
        middle_name = ""
        last_name = ""
        nickname = ""

        if is_individual:
            # Your data pattern: individual nickname often stored in dba
            nickname = (target.get("dba") or "").strip()

            # Best-effort parse of target["name"] into first/middle/last
            full = (target.get("name") or "").strip()
            parts = [p for p in re.split(r"\s+", full) if p]
            if len(parts) == 1:
                first_name = parts[0]
            elif len(parts) == 2:
                first_name, last_name = parts
            elif len(parts) >= 3:
                first_name = parts[0]
                last_name = parts[-1]
                middle_name = " ".join(parts[1:-1])

        return ensure_relation_dict({
            "name": target_label,
            "nickname": nickname or target_label,

            "first_name": first_name,
            "middle_name": middle_name,
            "last_name": last_name,

            "email": email,
            "phone": phone,

            "addr1": (target.get("addr1") or "").strip(),
            "addr2": (target.get("addr2") or "").strip(),
            "city":  (target.get("city") or "").strip(),
            "state": (target.get("state") or "").strip(),
            "zip":   (target.get("zip") or "").strip(),
            "dob": (target.get("dob") or "").strip() if isinstance(target.get("dob"), str) else "",

            "role": "linked",

            "linked_client_id": target_link_id,
            "linked_client_label": target_label,
        })


    def _sync_bidirectional_link(self, a_idx: int, b_idx: int, *, add: bool) -> bool:
        """
        Ensure both sides reflect the same link state.
        If add=True: add/update link record on both sides.
        If add=False: remove link record on both sides.
        Returns True if anything changed.
        """
        if a_idx is None or b_idx is None:
            return False
        if a_idx == b_idx:
            return False

        changed = False
        a_to_b_id = self._client_link_id(b_idx)
        b_to_a_id = self._client_link_id(a_idx)

        if add:
            rec_a = self._build_link_record_from_client(a_idx, b_idx)
            rec_b = self._build_link_record_from_client(b_idx, a_idx)

            # Choose where these link-records live:
            # - for businesses, you probably want them under "relations"
            # - if you want "employees" for some cases, pass role_key into caller
            changed |= self._upsert_link_record(a_idx, "relations", rec_a)
            changed |= self._upsert_link_record(b_idx, "relations", rec_b)
        else:
            changed |= self._remove_link_record(a_idx, "relations", a_to_b_id)
            changed |= self._remove_link_record(b_idx, "relations", b_to_a_id)

        return changed

    def _sync_link_change_for_person(self, this_client_idx: int, prev_link_id: str, new_link_id: str) -> None:
        """
        Called when an relations/Employee's linked_client_id changes.
        Mirrors unlink/link on BOTH clients.
        """
        prev_link_id = (prev_link_id or "").strip()
        new_link_id  = (new_link_id or "").strip()

        prev_idx = self._find_client_idx_by_id_or_ein(prev_link_id) if prev_link_id else None
        new_idx  = self._find_client_idx_by_id_or_ein(new_link_id)  if new_link_id else None

        changed = False

        # unlink old
        if prev_idx is not None:
            changed |= self._sync_bidirectional_link(this_client_idx, prev_idx, add=False)

        # link new
        if new_idx is not None:
            changed |= self._sync_bidirectional_link(this_client_idx, new_idx, add=True)

        if changed:
            save_clients(self.items)
            self.populate()
            self._update_suggestions()

    def link_personnel_from_client(self, client_idx: int):
        """
        UI action for your new 'Link' button in Personnel/Business:
        - user searches for a client to link
        - relations dialog opens prefilled
        - on save: record added to this client + bidirectional link ensured
        """
        # candidates: clients only (avoid showing personnels)
        cands = []
        try:
            cands = self.build_link_candidates()
        except Exception:
            cands = []

        if not cands:
            messagebox.showinfo("Link", "No clients available to link.")
            return

        # --- small chooser dialog (Entry + autocomplete) ---
        win = tk.Toplevel(self)
        win.title("Link to Client")
        win.transient(self.winfo_toplevel())
        win.grab_set()

        v = tk.StringVar(value="")
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Search client to link:").pack(anchor="w")
        ent = ttk.Entry(frm, textvariable=v, width=60)
        ent.pack(fill="x", pady=(4, 8))

        label_to_id = {x.get("label",""): x.get("id","") for x in cands if x.get("label")}
        labels = [x.get("label","") for x in cands if x.get("label")]

        popup = AutocompletePopup(win, ent, on_choose=lambda txt: v.set(txt))

        def matches(prefix: str) -> list[str]:
            tokens = (prefix or "").strip().lower().split()
            if not tokens:
                return labels
            out = []
            for lab in labels:
                low = lab.lower()
                if all(t in low for t in tokens):
                    out.append(lab)
            return out

        def refresh_popup(*_):
            popup.show(matches(v.get()))
            # preselect first
            if popup.listbox.size() > 0:
                popup.listbox.selection_clear(0, "end")
                popup.listbox.selection_set(0)
                popup.listbox.activate(0)

        def choose_current():
            # pick selected from popup or exact match
            if popup.winfo_viewable():
                cur = popup.listbox.curselection()
                if cur:
                    txt = popup.listbox.get(int(cur[0]))
                    v.set(txt)
            txt = v.get().strip()
            if txt not in label_to_id:
                messagebox.showwarning("Link", "Please choose a client from the list.")
                return None
            return txt, label_to_id[txt]

        def on_ok():
            res = choose_current()
            if not res:
                return
            lab, link_id = res

            target_idx = self._find_client_idx_by_id_or_ein(link_id)
            if target_idx is None:
                messagebox.showwarning("Link", "Could not resolve selected client.")
                return

            # ---- enforce required ID before linking (NO random/idx IDs) ----
            target = self.items[target_idx]
            is_individual = bool(target.get("is_individual")) or ((target.get("entity_type") or "").strip().casefold() == "individual")

            if is_individual:
                ssn9 = normalize_ssn_digits(target.get("ssn", ""))
                if not ssn9:
                    messagebox.showerror("Link", "This person client has no SSN.\n\nAdd SSN first before linking.")
                    return
            else:
                ein9 = normalize_ein_digits(target.get("ein", ""))
                if not ein9:
                    messagebox.showerror("Link", "This business client has no EIN.\n\nAdd EIN first before linking.")
                    return

            # Prefill relations record
            initial = self._build_link_record_from_client(client_idx, target_idx)
            # ensure dialog shows correct link selection
            initial["linked_client_id"] = self._client_link_id(target_idx)
            initial["linked_client_label"] = self._client_label(target_idx)
            initial["role"] = "linked"

            stable_id = self._client_link_id(target_idx)
            if not stable_id:
                messagebox.showerror("Link", "Missing required SSN/EIN. Add it first before linking.")
                return

            initial["linked_client_id"] = stable_id

            # Open LinkDialog prefilled
            d = LinkDialog(self, "Linked Person", initial=initial, link_candidates=cands)
            win.destroy()
            self.wait_window(d)

            if d.result:
                newp = ensure_relation_dict(d.result)

                # Add/replace on this client (avoid duplicates by linked_client_id)
                changed = self._upsert_link_record(client_idx, "relations", newp)

                # Bidirectional mirror (always ensure)
                target_idx2 = self._find_client_idx_by_id_or_ein(newp.get("linked_client_id", ""))
                if target_idx2 is not None:
                    changed |= self._sync_bidirectional_link(client_idx, target_idx2, add=True)

                if changed:
                    save_clients(self.items)
                    self.populate()
                    self._update_suggestions()

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(6, 0))
        ttk.Button(btns, text="Cancel", command=lambda: (popup.hide(), win.destroy())).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(btns, text="Link", command=on_ok).pack(side=tk.RIGHT)

        def on_keyrelease(e=None):
            if e and e.keysym in ("Up","Down","Return","Escape","Prior","Next"):
                return
            refresh_popup()

        ent.bind("<KeyRelease>", on_keyrelease)
        ent.bind("<FocusIn>", lambda e: refresh_popup())
        ent.bind("<Return>", lambda e: on_ok())
        ent.bind("<Escape>", lambda e: (popup.hide(), win.destroy(), "break"))

        win.after(50, ent.focus_set)

    # ---------- Sorting ----------
    def sort_by(self, col, descending=False):
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        def _key(v):
            s = str(v[0])
            try: return float(s.replace(",",""))
            except Exception: return s.lower()
        data.sort(key=_key, reverse=descending)
        for idx, item in enumerate(data):
            self.tree.move(item[1], "", idx)
        self.tree.heading(col, command=lambda c=col: self.sort_by(c, not descending))

    # ---------- Navigation ----------
    def _on_main_double_click(self, _evt=None):
        idx = self.selected_index()
        if idx is None: return
        self._ac.hide()
        # remember the row we’re leaving from
        try:
            c = self.items[idx]
            self._return_focus_key = (c.get("ein") or c.get("name") or "").strip()
        except Exception:
            self._return_focus_key = None
        self.navigate("detail", idx, push=True)

    def _ensure_main_page(self):
        """Recreate main page and tree if they were destroyed."""
        if not hasattr(self, "page_main") or not self.page_main.winfo_exists():
            self._build_main_page()
        elif not hasattr(self, "tree") or not self.tree.winfo_exists():
            for w in self.page_main.winfo_children():
                w.destroy()
            self._build_main_page()

    def _ensure_search_page(self):
        if not hasattr(self, "page_search") or not self.page_search.winfo_exists():
            self._build_search_page()
        elif not hasattr(self, "tree") or not self.tree.winfo_exists():
            for w in self.page_search.winfo_children():
                w.destroy()
            self._build_search_page()

    def _ensure_notes_page(self):
        # Always (re)create against current page_host if needed.
        self.page_notes = self.notes.ensure(self.page_host)
        try:
            self.notes.refresh()
        except Exception:
            pass


    def _is_valid_person_payload(self, data):
        return is_valid_person_payload(data)

    def _page_key(self, state):
        """Return a comparable key for a (kind, payload) state."""
        if not isinstance(state, tuple) or len(state) != 2:
            return ("main", None)
        kind, data = state
        if kind == "detail":
            try:
                return (kind, int(data))
            except Exception:
                return (kind, None)
        elif kind == "person":
            try:
                ci, role_key, pidx = data
                return (kind, (int(ci), str(role_key), int(pidx)))
            except Exception:
                return (kind, None)
        else:
            # 'main', 'search', or future kinds
            return (str(kind), data)
    
    def _compress_stack(self, stack):
        """Collapse consecutive duplicate pages in-place (A-A-B-A -> A-B-A)."""
        if not stack:
            return stack
        out = [stack[0]]
        for s in stack[1:]:
            if self._page_key(s) != self._page_key(out[-1]):
                out.append(s)
        stack[:] = out
        return stack
    
    def _compress_history(self):
        """Compress both history and future to remove consecutive dups."""
        if hasattr(self, "_history"):
            self._compress_stack(self._history)
        if hasattr(self, "_future"):
            self._compress_stack(self._future)

    # ---------- History helpers (class methods) ----------
    def _describe_state(self, state):
        kind, data = state if isinstance(state, tuple) and len(state) == 2 else ("main", None)
        if kind == "main":
            return "Home"
        if kind == "search":
            q = getattr(self, "q", None)
            return f"Search: {q.get().strip()}" if q and q.get().strip() else "Search"
        if kind == "detail":
            try:
                idx = int(data)
                name = ""
                if 0 <= idx < len(getattr(self, "items", [])):
                    name = self.items[idx].get("name", "")
                return f"Detail: {name or idx}"
            except Exception:
                return "Detail"
        if kind == "person":
            try:
                ci, role_key, pidx = data
                base = ""
                if 0 <= int(ci) < len(getattr(self, "items", [])):
                    base = self.items[int(ci)].get("name", "")
                return f"Person: {base or ci} / {role_key} / #{pidx}"
            except Exception:
                return "Person"
        return str(kind).title()
    
    def _render_current_page(self):
        self._clear_page_host()
        kind, data = getattr(self, "_current_page", ("main", None))

        if kind == "main":
            self._ensure_main_page()
            self.page_main.pack(fill=tk.BOTH, expand=True)

        elif kind == "search":
            self._ensure_search_page()
            self.page_search.pack(fill=tk.BOTH, expand=True)
            self.populate()

        elif kind == "notes":
            # IMPORTANT: make back/forward able to render Notes
            if not hasattr(self, "page_notes") or not self.page_notes.winfo_exists():
                # your navigate() uses either self._ensure_notes_page() or self.notes.ensure()
                if hasattr(self, "_ensure_notes_page"):
                    self._ensure_notes_page()
                else:
                    self.page_notes = self.notes.ensure(self.page_host)
            self.page_notes.pack(fill=tk.BOTH, expand=True)
            try:
                if hasattr(self.page_notes, "refresh"):
                    self.page_notes.refresh()
            except Exception:
                pass

        elif kind == "taxes":
            ChecklistPage(app=self).ensure(self.page_host)

        elif kind == "reports":
            ReportsPage(app=self).ensure(self.page_host)

        elif kind == "actions":
            self._ensure_actions_page()
            if data:
                try:
                    self._actions_page.preselect(str(data))
                except Exception:
                    pass
            self.page_actions.pack(fill=tk.BOTH, expand=True)

        elif kind == "detail":
            self._build_detail_page(int(data))

        else:  # person / unknown
            if hasattr(self, "_is_valid_person_payload") and self._is_valid_person_payload(data):
                ci, role_key, pidx = data
                self._build_person_page(int(ci), str(role_key), int(pidx))
            else:
                self._ensure_main_page()
                self.page_main.pack(fill=tk.BOTH, expand=True)

        self._update_nav_buttons()

    
    def _show_back_menu(self, e):
        if not getattr(self, "_history", []):
            return
        m = tk.Menu(self, tearoff=False)
        for i, state in enumerate(reversed(self._history)):
            label = self._describe_state(state)
            idx = len(self._history) - 1 - i
            m.add_command(label=label, command=lambda _idx=idx: self._jump_back_to(_idx))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()
    
    def _show_forward_menu(self, e):
        if not getattr(self, "_future", []):
            return
        m = tk.Menu(self, tearoff=False)
        for j, state in enumerate(self._future):
            label = self._describe_state(state)
            m.add_command(label=label, command=lambda _j=j: self._jump_forward_to(_j))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()
    
    def _jump_back_to(self, idx_in_history: int):
        if not (0 <= idx_in_history < len(self._history)):
            return
        chosen = self._history[idx_in_history]
        tail_after = self._history[idx_in_history+1:]
        self._future = tail_after + [getattr(self, "_current_page", ("main", None))] + self._future
        self._history = self._history[:idx_in_history]
        if hasattr(self, "_compress_history"): self._compress_history()
        self._current_page = chosen
        self._render_current_page()
    
    def _jump_forward_to(self, j_in_future: int):
        if not (0 <= j_in_future < len(self._future)):
            return
        chosen = self._future[j_in_future]
        before = self._future[:j_in_future]
        if hasattr(self, "_current_page"):
            self._history.append(self._current_page)
        self._history.extend(before)
        self._future = self._future[j_in_future+1:]
        if hasattr(self, "_compress_history"): self._compress_history()
        self._current_page = chosen
        self._render_current_page()

    def _focus_last_in_search(self):
        if self._last_viewed_idx is None: return
        if not hasattr(self, "tree") or not self.tree.winfo_exists(): return
        if not (0 <= self._last_viewed_idx < len(self.items)): return

        target = self.items[self._last_viewed_idx]
        wanted = (target.get("name") or target.get("dba") or target.get("ein") or "").strip().lower()
        if not wanted: return

        try:
            name_col_idx = self.COLS.index("name")
        except ValueError:
            name_col_idx = 0

        for iid in self.tree.get_children(""):
            vals = self.tree.item(iid, "values") or []
            if not vals: continue
            display = str(vals[name_col_idx]).strip().lower() if name_col_idx < len(vals) else ""
            if display == wanted:
                self.tree.selection_set(iid)
                self.tree.focus(iid)
                self.tree.see(iid)
                break


    def navigate(self, page: str, idx=None, payload=None, push: bool = True, replace: bool = False):
        target = None
        if page == "main":
            target = ("main", None)
        elif page == "search":
            target = ("search", None)
        elif page == "notes":
            target = ("notes", None)
        elif page == "taxes":
            if self._last_viewed_idx is not None and 0 <= self._last_viewed_idx < len(self.items):
                c = self.items[self._last_viewed_idx]
                key = c.get("ein") or c.get("name") or c.get("dba")
                if key and hasattr(self, "checklist") and hasattr(self.checklist, "focus_client_by_key"):
                    self.checklist.focus_client_by_key(key) 
            target = ("taxes", None)
        elif page == "detail":
            # accept either idx or payload for backwards-compat
            sel = idx if idx is not None else payload
            if sel is None:
                # nothing to open; just ignore
                return
            try:
                sel = int(sel)
            except Exception:
                # if it can't be coerced, abort quietly
                return
            self._last_viewed_idx = sel
            target = ("detail", sel)
        elif page == "person":
            if self._is_valid_person_payload(payload):
                cmp_idx, role_key, pidx = payload
                target = ("person", (int(cmp_idx), str(role_key), int(pidx)))
        elif page == "actions":
            # payload can be a preselected tool key or None
            target = ("actions", payload)

        else:
            target = ("main", None)
    
        if not hasattr(self, "_current_page"):
            self._current_page = ("main", None)
    
        if self._page_key(getattr(self, "_current_page", ("main", None))) == self._page_key(target):
            # Still render to ensure UI matches requested page
            self._current_page = target
        else:
            if replace:
                self._current_page = target
            elif push:
                if hasattr(self, "_current_page"):
                    # Only push if different from last on history
                    if not self._history or self._page_key(self._history[-1]) != self._page_key(self._current_page):
                        self._history.append(self._current_page)
                else:
                    self._history.append(("main", None))
                self._future.clear()
                self._current_page = target
            else:
                self._current_page = target

        self._compress_history()
    
        self._clear_page_host()
    
        kind, data = self._current_page
        
        self.log.info("navigate(kind=%s, push=%s)", kind, push)
        if kind == "main":
            self._ensure_main_page()
            self.page_main.pack(fill=tk.BOTH, expand=True)
        elif kind == "search":
            self._ensure_search_page()
            self.page_search.pack(fill=tk.BOTH, expand=True)
            self.populate()
            key = getattr(self, "_return_focus_key", None)
            if key:
                self._focus_client_in_search(key)
                self._return_focus_key = None
        elif kind == "notes":
            self._ensure_notes_page()
            self.page_notes.pack(fill=tk.BOTH, expand=True)
        elif kind == "taxes":
            ChecklistPage(app=self).ensure(self.page_host)
        elif kind == "reports":
            ReportsPage(app=self).ensure(self.page_host)

        elif kind == "detail":
            self._build_detail_page(data)
        elif kind == "actions":
            self._ensure_actions_page()
            # preselect tool if provided
            if data:
                try:
                    self._actions_page.preselect(str(data))
                except Exception:
                    pass
            self.page_actions.pack(fill=tk.BOTH, expand=True)

        else:
            self._build_person_page(data[0], data[1], data[2])
    
        self._update_nav_buttons()

    def save_clients_data(self):
        """Persist current self.items using the existing save_clients() helper."""
        try:
            save_clients(self.items)
        except Exception:
            pass

    def select_detail_tab(self, title: str):
        """
        Called from NotePage to switch to the Logs tab if your detail page uses a ttk.Notebook.
        Safe no-op if not found.
        """
        nb = getattr(self, "detail_notebook", None) or getattr(self, "detail_nb", None) or getattr(self, "notebook", None)
        if not nb or not hasattr(nb, "tabs"):
            return
        try:
            for tab_id in nb.tabs():
                try:
                    t = nb.tab(tab_id, "text")
                except Exception:
                    t = ""
                if (t or "").strip().casefold() == (title or "").strip().casefold():
                    nb.select(tab_id)
                    return
        except Exception:
            return


    def go_home(self, push=True):
        self.navigate("main", push=push)

    def nav_back(self):
        self.log.info("nav_back()")
        if not getattr(self, "_history", None):
            return
        if not self._history:
            return

        prev = self._history.pop()
        self._future.append(self._current_page)
        self._current_page = prev

        self._render_current_page()

    def nav_forward(self):
        self.log.info("nav_forward()")
        if not getattr(self, "_future", None):
            return
        if not self._future:
            return

        nxt = self._future.pop()
        self._history.append(self._current_page)
        self._current_page = nxt

        self._render_current_page()


    def _update_nav_buttons(self):
        self.btn_back["state"] = tk.NORMAL if bool(getattr(self, "_history", [])) else tk.DISABLED
        self.btn_fwd["state"]  = tk.NORMAL if bool(getattr(self, "_future", [])) else tk.DISABLED

    def _find_client_idx_by_id_or_ein(self, link_id: str) -> int | None:
        """
        link_id formats we allow:
        - "client:<client_id>"     (legacy/support)
        - "ein:<9digits>"          (business)
        - "ssn:<9digits>"          (individual)
        Returns client index in self.items or None.
        """
        link_id = (link_id or "").strip()
        if not link_id:
            return None

        try:
            kind, val = link_id.split(":", 1)
        except ValueError:
            return None

        kind = kind.strip().lower()
        val = (val or "").strip()

        if kind == "client":
            if not val:
                return None
            for i, c in enumerate(self.items):
                if str(c.get("id", "") or "").strip() == val:
                    return i
            return None

        if kind == "ein":
            ein9 = normalize_ein_digits(val)
            if not ein9:
                return None
            for i, c in enumerate(self.items):
                if normalize_ein_digits(c.get("ein", "")) == ein9:
                    return i
            return None

        if kind == "ssn":
            ssn9 = normalize_ssn_digits(val)
            if not ssn9:
                return None
            for i, c in enumerate(self.items):
                if normalize_ssn_digits(c.get("ssn", "")) == ssn9:
                    return i
            return None

        return None


    def open_linked_target(self, linked_client_id: str) -> bool:
        """
        If linked_client_id resolves to a client, navigate to its Detail page.
        Returns True if navigation happened.
        """
        idx = self._find_client_idx_by_id_or_ein(linked_client_id)
        if idx is None:
            return False
        self.navigate("detail", idx, push=True)
        return True
    
    def link_clients(self, a_id: str, b_id: str, link: bool, role: str = ""):
        """
        Bidirectional linking using stable ids:
        - Individuals use ssn:<9>
        - Businesses use ein:<9>

        role:
        - If linking A -> B and B is a business: A-side role becomes "business"
        - If linking A -> B and B is an individual: A-side role uses `role` (can be "", spouse, employee, etc.)
        - Reciprocal B -> A follows the same rule (if A is business -> "business" else `role`)
        Unlink:
        - removes links in BOTH directions regardless of role
        """
        a_id = str(a_id or "").strip()
        b_id = str(b_id or "").strip()
        if not a_id or not b_id or a_id == b_id:
            return

        ai, a = self._resolve_client_key(a_id)
        bi, b = self._resolve_client_key(b_id)

        # If unlinking, missing/stale ids should not crash the UI
        if a is None or b is None:
            if not link:
                return
            raise ValueError("Could not resolve one or both clients for linking.")

        def _ensure_relations_list(c: dict):
            if c.get("relations") is None or not isinstance(c.get("relations"), list):
                c["relations"] = []
            return c["relations"]

        def _remove_link(off_list: list, target_id: str):
            tid = str(target_id or "").strip()
            if not tid:
                return
            keep = []
            for o in off_list:
                if not isinstance(o, dict):
                    continue
                if str(o.get("linked_client_id") or "").strip() == tid:
                    continue
                keep.append(o)
            off_list[:] = keep

        def _upsert_link(off_list: list, record: dict):
            record = ensure_relation_dict(record)
            tid = str(record.get("linked_client_id") or "").strip()
            if not tid:
                return
            for i, o in enumerate(off_list):
                if isinstance(o, dict) and str(o.get("linked_client_id") or "").strip() == tid:
                    off_list[i] = record
                    return
            off_list.append(record)

        def _is_individual_client(c: dict) -> bool:
            return bool(c.get("is_individual")) or ((c.get("entity_type") or "").strip().casefold() == "individual")

        def _pick_contact_from_client(c: dict) -> tuple[str, str]:
            """
            Best-effort: use c['email']/c['phone'] if present, else first relations entry.
            """
            email = (c.get("email") or "").strip() if isinstance(c.get("email"), str) else ""
            phone = (c.get("phone") or "").strip() if isinstance(c.get("phone"), str) else ""

            if email and phone:
                return email, phone

            try:
                rels = c.get("relations", []) or []
                if rels:
                    o0 = ensure_relation_dict(rels[0])
                    if not email:
                        email = (o0.get("email") or "").strip()
                    if not phone:
                        phone = (o0.get("phone") or "").strip()
            except Exception:
                pass

            return email, phone

        def _split_name(full: str) -> tuple[str, str, str]:
            parts = [p for p in re.split(r"\s+", (full or "").strip()) if p]
            if len(parts) == 0:
                return ("", "", "")
            if len(parts) == 1:
                return (parts[0], "", "")
            if len(parts) == 2:
                return (parts[0], "", parts[1])
            return (parts[0], " ".join(parts[1:-1]), parts[-1])

        def _record_from_client(src: dict, src_id_key: str, role_value: str) -> dict:
            """
            Build a relations-style dict referencing src, stored inside the OTHER client.
            """
            src_label = (src.get("name") or "").strip() or src_id_key
            is_ind = _is_individual_client(src)

            email, phone = _pick_contact_from_client(src)

            first_name = middle_name = last_name = ""
            nickname = ""

            if is_ind:
                nickname = (src.get("dba") or "").strip()
                fn, mn, ln = _split_name(src.get("name") or "")
                first_name, middle_name, last_name = fn, mn, ln

            rec = {
                "name": src_label,
                "nickname": nickname or "",
                "first_name": first_name,
                "middle_name": middle_name,
                "last_name": last_name,
                "email": email,
                "phone": phone,
                "addr1": (src.get("addr1") or "").strip(),
                "addr2": (src.get("addr2") or "").strip(),
                "city":  (src.get("city") or "").strip(),
                "state": (src.get("state") or "").strip(),
                "zip":   (src.get("zip") or "").strip(),
                "dob":   (src.get("dob") or "").strip() if isinstance(src.get("dob"), str) else "",
                "role": role_value or "",
                "linked_client_id": src_id_key,
                "linked_client_label": src_label,
            }
            return ensure_relation_dict(rec)

        # Compute stable IDs for both sides (must be ssn:<9> or ein:<9>)
        a_key = self._canonical_client_key(a, ai)
        b_key = self._canonical_client_key(b, bi)

        # If linking, require stable IDs on both ends.
        if link:
            if not a_key or not b_key:
                raise ValueError("Cannot link: one or both clients are missing required SSN/EIN (stable id).")

            a_rels = _ensure_relations_list(a)
            b_rels = _ensure_relations_list(b)

            # Decide roles on each side based on the TARGET being business/individual
            role_a_to_b = "business" if not _is_individual_client(b) else (role or "")
            role_b_to_a = "business" if not _is_individual_client(a) else (role or "")

            _upsert_link(a_rels, _record_from_client(b, b_key, role_a_to_b))
            _upsert_link(b_rels, _record_from_client(a, a_key, role_b_to_a))

        else:
            # Unlink is best-effort; if keys missing, still attempt via provided ids
            # Prefer computed keys if available.
            if not a_key:
                a_key = a_id
            if not b_key:
                b_key = b_id

            a_rels = _ensure_relations_list(a)
            b_rels = _ensure_relations_list(b)

            _remove_link(a_rels, b_key)
            _remove_link(b_rels, a_key)

        # Persist + refresh UI
        save_clients(self.items)
        self.populate()
        self._update_suggestions()


# -------------------- Entrypoint --------------------
def main():
    if sys.platform.startswith("win"):
        if not enforce_single_instance("Vertex"):
            return

    log = get_logger("launcher")
    log.info("Launching Client Manager main()")
    root = tk.Tk()

    if enforce_major_update_on_startup(
        root,
        app_name=APP_NAME,
        app_version=APP_VERSION,
        github_api_latest=GITHUB_API_LATEST,
        github_releases_url=GITHUB_RELEASES_URL,
        update_asset_name=UPDATE_ASSET_NAME,
    ):
        root.mainloop()
        return
    
    def resource_path(rel: str) -> str:
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        return str(base / rel)

    # Must be set early in the process on Windows for consistent taskbar identity/icon
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Vertex.Vertex")
        except Exception:
            pass

    # After root is created:
    try:
        ico = resource_path("logo.ico")
        root.iconbitmap(default=ico)
    except Exception:
        pass

    root.title("Vertex")
    root.geometry("1440x720")
    root.minsize(1440, 720)

    try:
        from styles.new_ui import NewUI
        NewUI.install(root)
        LOG.info("NewUI installed")
    except Exception as e:
        LOG.warning("NewUI not applied: %s", e)

    try:
        style = ttk.Style(); style.theme_use("clam")
    except Exception:
        pass
    
    App(root)
    root.mainloop()

if __name__ == "__main__":
    # Tool-runner mode for PyInstaller builds:
    # We spawn the same EXE with "--run-tool <script.py>" so tools run with the bundled Python/pandas.
    if "--run-tool" in sys.argv:
        import runpy
        from pathlib import Path

        try:
            i = sys.argv.index("--run-tool")
            script_name = sys.argv[i + 1]
        except Exception:
            # Bad invocation → fail fast with a clear message
            raise SystemExit("Usage: vertex.exe --run-tool <script.py> [args...]")

        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        tool_path = base / script_name
        if not tool_path.exists():
            raise SystemExit(f"Tool not found: {tool_path}")

        # Rebuild argv so the tool feels like it was launched directly
        sys.argv = [str(tool_path)] + sys.argv[i + 2 :]

        # Run the tool as __main__
        runpy.run_path(str(tool_path), run_name="__main__")
        raise SystemExit(0)

    main()
