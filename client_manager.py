from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import os, sys, json, re, webbrowser, subprocess, datetime, urllib.request, urllib.error, ssl, urllib.parse
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import tkinter.font as tkfont
from tkinter.scrolledtext import ScrolledText
import csv

APP_NAME = "Vertex"

# 🔢 bump this each time you ship a new version
APP_VERSION = "0.1.0"

# 🔗 set this to your real GitHub repo once you create it,
# e.g. "seunghyunyang/lineup-client-manager"
GITHUB_REPO = "shyang9711/vertex"
GITHUB_RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases"
GITHUB_API_LATEST   = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

# Optional: name of the EXE asset in GitHub Releases if you want auto-download
UPDATE_ASSET_NAME = "vertex.exe"

# pages
from functions.pages.dashboard_page import DashboardPage
from functions.pages.profile_tab import init_profile_tab
from functions.pages.documents_tab import init_documents_tab
from functions.pages.logs_tab import init_logs_tab
from functions.pages.checklist_page import ChecklistPage
from functions.pages.action_page import ActionRunnerPage
from functions.pages.reports_page import ReportsPage

from functions.models.taskbar_model import TaskbarModel
from functions.utils.app_logging import get_logger

# NewUI preference from styles/, fallback to functions/
try:
    from styles.new_ui import NewUI
except Exception:
    NewUI = None

LOG = get_logger("client_manager")

# -------------------- Constants --------------------
ENTITY_TYPES = [
    "", "Individual / Sole Proprietor", "Partnership", "LLC",
    "S-Corporation", "Corporation (C-Corp)", "Exempt Organization",
    "Trust / Estate", "Nonprofit", "Other"
]
US_STATES = [
    "", "AL","AK","AZ","AR","CA","CO","CT","DE","DC","FL","GA","HI","ID","IL","IN",
    "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA",
    "WA","WV","WI","WY"
]

# -------------------- Storage paths --------------------
def app_dir() -> Path:
    """
    Base directory for the app:
    - When frozen (PyInstaller EXE): folder containing the EXE
    - When running from source: folder containing client_manager.py
    """
    if getattr(sys, "frozen", False):
        # PyInstaller onefile / onedir
        return Path(sys.executable).resolve().parent
    try:
        return Path(__file__).parent.resolve()
    except NameError:
        return Path(os.getcwd()).resolve()


APP_DIR = app_dir()

DATA_ROOT   = APP_DIR / "data"
CLIENTS_DIR = DATA_ROOT / "clients"

CLIENTS_DIR.mkdir(parents=True, exist_ok=True)

DATA_DIR  = CLIENTS_DIR
DATA_FILE = CLIENTS_DIR / "clients.json"

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

def tokenize(s: str) -> List[str]:
    if s is None: return []
    s = str(s).lower().strip()
    parts = re.split(r"[^a-z0-9@._\-]+", s)
    return [p for p in parts if p]

def norm_text(s: str) -> str:
    return " ".join(tokenize(s))

def compose_person_name(first: str, middle: str, last: str, nickname: str) -> str:
    parts = [first.strip(), middle.strip(), last.strip()]
    base = " ".join([p for p in parts if p])
    if nickname.strip():
        return f'{base} ("{nickname.strip()}")' if base else nickname.strip()
    return base

def ensure_officer_dict(x) -> Dict[str, str]:
    """
    Normalize personnel dict.
    If only 'name' exists, move it to first_name (migration) and compose display 'name'.
    """
    if not isinstance(x, dict):
        x = {"name": str(x).strip()}

    o = {
        "name":        str(x.get("name","")).strip(),
        "first_name":  str(x.get("first_name","")).strip(),
        "middle_name": str(x.get("middle_name","")).strip(),
        "last_name":   str(x.get("last_name","")).strip(),
        "nickname":    str(x.get("nickname","")).strip(),
        "email":       str(x.get("email","")).strip(),
        "phone":       str(x.get("phone","")).strip(),
        "addr1":       str(x.get("addr1","")).strip(),
        "addr2":       str(x.get("addr2","")).strip(),
        "city":        str(x.get("city","")).strip(),
        "state":       str(x.get("state","")).strip(),
        "zip":         str(x.get("zip","")).strip(),
        "dob":         str(x.get("dob","")).strip(),
    }
    if o["name"] and not (o["first_name"] or o["middle_name"] or o["last_name"] or o["nickname"]):
        o["first_name"] = o["name"]
    composed = compose_person_name(o["first_name"], o["middle_name"], o["last_name"], o["nickname"])
    if composed:
        o["name"] = composed
    return o

def display_officer_name(o: Dict[str, str]) -> str:
    o = ensure_officer_dict(o)
    return o.get("name","").strip()

def officers_to_display_lines(officers: List[Dict[str,str]]) -> List[str]:
    return [display_officer_name(o) for o in (officers or []) if display_officer_name(o)]

def officers_to_flat_emails(officers: List[Dict[str,str]]) -> List[str]:
    seen, out = set(), []
    for o in officers or []:
        e = str(ensure_officer_dict(o).get("email","")).strip()
        if e and e not in seen:
            seen.add(e); out.append(e)
    return out

def officers_to_flat_phones(officers: List[Dict[str,str]]) -> List[str]:
    seen, out = set(), []
    for o in officers or []:
        p = str(ensure_officer_dict(o).get("phone","")).strip()
        if p and p not in seen:
            seen.add(p); out.append(p)
    return out

def is_valid_person_payload(data):
    return isinstance(data, (tuple, list)) and len(data) == 3 and data[0] is not None

# ---- Quarter / Tax helpers ---------------------------------------------------
def today_date() -> datetime.date:
    return datetime.date.today()

def quarter_start(d: datetime.date) -> datetime.date:
    q = (d.month - 1) // 3
    first_month = q*3 + 1
    return datetime.date(d.year, first_month, 1)

def new_quarter_started(last_checked_iso: str | None) -> bool:
    try:
        if not last_checked_iso:
            return True
        last = datetime.date.fromisoformat(last_checked_iso)
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
            c.setdefault("file_location","")
            c.setdefault("memo","")

            # Company extended fields
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
            officers = c.get("officers")
            if officers is None:
                legacy = c.get("owner","")
                officers = [legacy] if legacy else []
            norm_offs: List[Dict[str,str]] = []
            if isinstance(officers, list):
                for x in officers:
                    norm_offs.append(ensure_officer_dict(x))
            elif isinstance(officers, dict):
                norm_offs.append(ensure_officer_dict(officers))
            else:
                norm_offs = [ensure_officer_dict(officers)] if officers else []
            c["officers"] = norm_offs
            c.pop("owner", None)

            # Optional employees list (if present, normalize same way)
            emps = c.get("employees", [])
            if isinstance(emps, list):
                c["employees"] = [ensure_officer_dict(x) for x in emps]
            elif isinstance(emps, dict):
                c["employees"] = [ensure_officer_dict(emps)]
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

        # Normalize officers/employees
        cc["officers"]  = [ensure_officer_dict(o) for o in cc.get("officers", [])]
        cc["employees"] = [ensure_officer_dict(o) for o in cc.get("employees", [])]

        # Ensure keys exist
        for k, v in {
            "name":"", "dba":"", "ein":"", "file_location":"", "memo":"",
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


def export_all_to_json(path: Path, clients: List[Dict[str, Any]]) -> None:
    """
    Export ALL program data to a single JSON file chosen by the user.
    For now this includes only clients; later we can add tasks, match_rules, etc.
    """
    payload: Dict[str, Any] = {
        "clients": _normalize_clients_for_io(clients),
        # "tasks": tasks_list,
        # "match_rules": rules_list,
    }
    try:
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        LOG.exception("Error exporting data: %s", e)
        messagebox.showerror("Export Error", f"Couldn't export data:\n{e}")


def import_all_from_json(path: Path, clients: List[Dict[str, Any]]) -> int:
    """
    Import data from JSON and merge into the program.

    - For CLIENTS: only add if no existing client has the same EIN.
    - Existing items are NOT overwritten.
    Returns: number of new clients added.
    """
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        LOG.exception("Error reading import file: %s", e)
        messagebox.showerror("Import Error", f"Couldn't read file:\n{e}")
        return 0

    # Accept either { "clients": [...] } or a bare list
    if isinstance(data, dict):
        raw_clients = data.get("clients", [])
    elif isinstance(data, list):
        raw_clients = data
    else:
        messagebox.showerror("Import Error", "Unrecognized data format in JSON.")
        return 0

    if not isinstance(raw_clients, list):
        messagebox.showerror("Import Error", "Expected 'clients' to be a list.")
        return 0

    # Index existing by EIN
    existing_by_ein: Dict[str, Dict[str, Any]] = {}
    for c in clients:
        ein = (c.get("ein") or "").strip()
        if ein:
            existing_by_ein[ein] = c

    added = 0
    for raw in raw_clients:
        if not isinstance(raw, dict):
            continue
        ein = (raw.get("ein") or "").strip()
        if not ein:
            # Import rule: if EIN is missing, don't auto-create; keeps behavior predictable.
            continue
        if ein in existing_by_ein:
            # already have this EIN → skip, do NOT overwrite
            continue

        c = dict(raw)
        # basic normalization like load_clients
        c.setdefault("name", "")
        c.setdefault("dba", "")
        c.setdefault("file_location", "")
        c.setdefault("memo", "")
        c.setdefault("addr1", "")
        c.setdefault("addr2", "")
        c.setdefault("city", "")
        c.setdefault("state", "")
        c.setdefault("zip", "")
        c.setdefault("acct_mgr", "")
        c.setdefault("edd_number", "")
        c.setdefault("sales_tax_account", "")
        c.setdefault("entity_type", "")
        c.setdefault("ui_rate", "")
        c.setdefault("sales_tax_rate", "")
        c.setdefault("other_tax_rates", "")
        c.setdefault("tax_rates_last_checked", "")
        c.setdefault("logs", [])
        c.setdefault("officers", [])
        c.setdefault("employees", [])
        c["officers"] = [ensure_officer_dict(o) for o in c.get("officers", [])]
        c["employees"] = [ensure_officer_dict(o) for o in c.get("employees", [])]

        clients.append(c)
        existing_by_ein[ein] = c
        added += 1

    return added

def _parse_version(s: str) -> tuple[int, int, int]:
    """
    Turn strings like 'v0.1.3', '0.2', '1.0.0-beta' into (major, minor, patch).
    Non-numeric parts are ignored.
    """
    nums = re.findall(r"\d+", str(s))
    major = int(nums[0]) if len(nums) > 0 else 0
    minor = int(nums[1]) if len(nums) > 1 else 0
    patch = int(nums[2]) if len(nums) > 2 else 0
    return (major, minor, patch)


def _is_newer_version(latest: str, current: str) -> bool:
    return _parse_version(latest) > _parse_version(current)


def check_for_updates(parent: tk.Misc | None = None):
    """
    Check GitHub Releases for a newer version.
    - If no releases exist yet: show a friendly message, NOT a failure.
    - If newer exists:
        * When running as EXE, download a new EXE next to the current one as -LATEST.exe
        * When running from source, just open the Releases page.
    Never touches DATA_ROOT / data/ folder.
    """
    # --- call GitHub API for latest release ---
    try:
        req = urllib.request.Request(
            GITHUB_API_LATEST,
            headers={"User-Agent": APP_NAME},
        )
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, context=ctx, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        # Most likely: 404 = no releases yet
        if e.code == 404:
            messagebox.showinfo(
                "Updates",
                "No releases found on GitHub yet.\n\n"
                "Once you create a release for this app, "
                "the update checker will compare versions."
            )
            return
        messagebox.showerror("Update Check Failed", f"HTTP error from GitHub:\n{e}")
        return
    except Exception as e:
        messagebox.showerror("Update Check Failed", f"Could not contact GitHub:\n{e}")
        return

    tag = str(data.get("tag_name") or "").strip()
    if tag.lower().startswith("v"):
        tag = tag[1:]

    if not tag:
        messagebox.showinfo(
            "Updates",
            "Latest release has no tag_name.\n"
            "Use tags like v0.1.0 on GitHub releases."
        )
        return

    if not _is_newer_version(tag, APP_VERSION):
        messagebox.showinfo(
            "Up to date",
            f"You are running version {APP_VERSION}, which is the latest release."
        )
        return

    # --- Newer version exists ---
    if not getattr(sys, "frozen", False):
        # Running from source: just open Releases page
        if messagebox.askyesno(
            "Update available",
            f"Current version: {APP_VERSION}\nLatest version: {tag}\n\n"
            "Open the GitHub Releases page?"
        ):
            webbrowser.open(GITHUB_RELEASES_URL)
        return

    # Running as EXE: try to download a new EXE next to the current one
    assets = data.get("assets") or []
    exe_name = os.path.basename(sys.executable)
    url = None

    # Prefer asset that matches our exe name
    for a in assets:
        if a.get("name") == exe_name:
            url = a.get("browser_download_url")
            break
    # Or fallback to UPDATE_ASSET_NAME
    if not url and UPDATE_ASSET_NAME:
        for a in assets:
            if a.get("name") == UPDATE_ASSET_NAME:
                url = a.get("browser_download_url")
                break

    if not url:
        # No downloadable EXE found; just send user to Releases page
        if messagebox.askyesno(
            "Update available",
            f"Current version: {APP_VERSION}\nLatest version: {tag}\n\n"
            "No EXE asset found in the release.\n"
            "Open Releases page in your browser?"
        ):
            webbrowser.open(GITHUB_RELEASES_URL)
        return

    app_folder = Path(sys.executable).resolve().parent
    dest = app_folder / f"{exe_name.rsplit('.', 1)[0]}-LATEST.exe"

    try:
        with urllib.request.urlopen(url, context=ssl.create_default_context(), timeout=30) as resp:
            chunk_size = 1024 * 64
            with open(dest, "wb") as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
    except Exception as e:
        messagebox.showerror(
            "Update Download Failed",
            f"Could not download the latest EXE.\n\n{e}"
        )
        return

    messagebox.showinfo(
        "Update downloaded",
        f"Downloaded latest version to:\n{dest}\n\n"
        "Close this app and run the new EXE.\n"
        "Your data folder (data/) was not modified."
    )


def show_about_dialog(parent: tk.Misc | None = None):
    msg = (
        f"{APP_NAME}\n"
        f"Version {APP_VERSION}\n\n"
        f"Data folder:\n{DATA_ROOT}"
    )
    messagebox.showinfo("About", msg, parent=parent)

# -------------------- UI infrastructure --------------------
class AutocompletePopup(tk.Toplevel):
    """Small dropdown under an Entry, like Google suggestions (doesn't steal focus)."""
    def __init__(self, master, anchor_entry: tk.Entry, on_choose):
        super().__init__(master)
        self.withdraw()
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.anchor = anchor_entry
        self.on_choose = on_choose

        self.configure(bg=NewUI.BORDER)
        self.listbox = tk.Listbox(self, height=8, activestyle="none",
                                  bd=0, highlightthickness=0,
                                  relief="flat", font=("Segoe UI", 10))
        self.listbox.pack(fill="both", expand=True, padx=1, pady=1)
        self.listbox.bind("<ButtonRelease-1>", self._on_click_choose)
        self.listbox.bind("<Button-1>", self._on_mouse_down)
        self.listbox.bind("<Double-Button-1>", self._choose)

        self.listbox.unbind("<Up>")
        self.listbox.unbind("<Down>")

        self.listbox.bind("<Up>", self._lb_up)
        self.listbox.bind("<Down>", self._lb_down)

        self.listbox.bind("<Return>", self._choose)
        self.bind("<FocusOut>", self._maybe_hide)
        self.listbox.bind("<FocusOut>", self._maybe_hide)
        self.listbox.bind("<Escape>", lambda e: self.hide())

    def show(self, items: List[str]):
        self.listbox.delete(0, tk.END)
        for s in items[:20]:
            self.listbox.insert(tk.END, s)
        if not items:
            self.hide(); return
        x = self.anchor.winfo_rootx()
        y = self.anchor.winfo_rooty() + self.anchor.winfo_height()
        w = self.anchor.winfo_width()
        h = min(256, 22 * len(items))
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()
        if self.listbox.size() > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)

    def move_selection(self, delta: int):
        if not self.winfo_viewable(): return
        if self.listbox.size() == 0: return
        cur = self.listbox.curselection()
        i = cur[0] if cur else 0
        i = max(0, min(self.listbox.size()-1, i + delta))
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(i)
        self.listbox.activate(i)

    def current_text(self) -> Optional[str]:
        cur = self.listbox.curselection()
        if not cur: return None
        return self.listbox.get(cur[0])

    def _lb_up(self, event=None):
        self.move_selection(-1)
        return "break"

    def _lb_down(self, event=None):
        self.move_selection(+1)
        return "break"

    def _choose(self, *_):
        txt = self.current_text()
        if txt is None: return
        self.on_choose(txt)
        self.hide()

    def hide(self):
        if self.winfo_viewable():
            self.withdraw()

    def focus_listbox(self):
        self.listbox.focus_set()

    def _on_mouse_down(self, e):
        i = self.listbox.nearest(e.y)
        if 0 <= i < self.listbox.size():
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(i)
            self.listbox.activate(i)

    def _on_click_choose(self, e):
        # let the Listbox update selection, then choose
        self.after(1, self._choose)

    def _maybe_hide(self, _e=None):
        w = self.focus_get()
        # If the focus is not inside this toplevel, hide
        if not w or not str(w).startswith(str(self)):
            self.hide()

class ScrollFrame(ttk.Frame):
    """Simple vertical scrollable frame for long dialogs."""
    def __init__(self, master, height=520):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0, height=height)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        def _on_mousewheel(event):
            # Bail if the canvas is gone
            if not self.canvas.winfo_exists():
                return
            if event.delta:  # Windows / macOS
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            else:  # X11: Button-4/5
                if event.num == 4:
                    self.canvas.yview_scroll(-3, "units")
                elif event.num == 5:
                    self.canvas.yview_scroll(+3, "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind_all("<Button-4>", _on_mousewheel)
        self.canvas.bind_all("<Button-5>", _on_mousewheel)

# -------------------- Officer dialog (reused for editing) --------------------
class OfficerDialog(tk.Toplevel):
    def __init__(self, master, title="Officer", initial: Optional[Dict[str,str]]=None):
        super().__init__(master)
        self.title(title); self.resizable(False, False); self.result=None
        init = ensure_officer_dict(initial or {})

        self.v_first = tk.StringVar(value=init.get("first_name",""))
        self.v_mid   = tk.StringVar(value=init.get("middle_name",""))
        self.v_last  = tk.StringVar(value=init.get("last_name",""))
        self.v_nick  = tk.StringVar(value=init.get("nickname",""))
        self.v_email = tk.StringVar(value=init.get("email",""))
        self.v_phone = tk.StringVar(value=init.get("phone",""))
        self.v_addr1 = tk.StringVar(value=init.get("addr1",""))
        self.v_addr2 = tk.StringVar(value=init.get("addr2",""))
        self.v_city  = tk.StringVar(value=init.get("city",""))
        self.v_state = tk.StringVar(value=init.get("state",""))
        self.v_zip   = tk.StringVar(value=init.get("zip",""))
        self.v_dob   = tk.StringVar(value=init.get("dob",""))

        frm = ttk.Frame(self, padding=12); frm.grid(row=0, column=0, sticky="nsew")

        def row(lbl, var, r, c=0, width=42, colspan=1):
            ttk.Label(frm, text=lbl).grid(row=r, column=c, sticky="w", pady=(0,2))
            ent = ttk.Entry(frm, textvariable=var, width=width)
            ent.grid(row=r+1, column=c, sticky="we", pady=(0,6), columnspan=colspan); return ent

        e1 = row("First Name", self.v_first, 0, 0, 28)
        row("Middle Name", self.v_mid, 0, 1, 18)
        row("Last Name", self.v_last, 0, 2, 28)
        row("Nickname", self.v_nick, 0, 3, 18)

        row("Email", self.v_email, 2, 0, 40, colspan=2)
        row("Phone", self.v_phone, 2, 2, 24, colspan=2)

        row("Address 1", self.v_addr1, 4, 0, 50, colspan=4)
        row("Address 2", self.v_addr2, 6, 0, 50, colspan=4)

        ttk.Label(frm, text="City").grid(row=8, column=0, sticky="w", pady=(0,2))
        ent_city = ttk.Entry(frm, textvariable=self.v_city, width=32)
        ent_city.grid(row=9, column=0, sticky="we", pady=(0,6), columnspan=2)

        ttk.Label(frm, text="State").grid(row=8, column=2, sticky="w", pady=(0,2))
        cb_state = ttk.Combobox(frm, values=US_STATES, textvariable=self.v_state, width=8)
        cb_state.grid(row=9, column=2, sticky="we", pady=(0,6))

        ttk.Label(frm, text="Zip").grid(row=8, column=3, sticky="w", pady=(0,2))
        ent_zip = ttk.Entry(frm, textvariable=self.v_zip, width=12)
        ent_zip.grid(row=9, column=3, sticky="we", pady=(0,6))

        row("DOB (YYYY-MM-DD)", self.v_dob, 10, 0, 20)

        btns = ttk.Frame(frm); btns.grid(row=12, column=0, columnspan=4, sticky="e", pady=(6,0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda _e: self._save())
        self.bind("<Escape>", lambda _e: self.destroy())
        self.after(50, e1.focus_set); self.grab_set(); self.transient(master)

    def _save(self):
        first = self.v_first.get().strip()
        last  = self.v_last.get().strip()
        nick  = self.v_nick.get().strip()
        if not (first or last or nick):
            messagebox.showerror("Required", "Enter at least First or Last name (or Nickname)."); return
        name = compose_person_name(first, self.v_mid.get(), last, nick)
        self.result = {
            "name": name,
            "first_name": first,
            "middle_name": self.v_mid.get().strip(),
            "last_name": last,
            "nickname": nick,
            "email": self.v_email.get().strip(),
            "phone": self.v_phone.get().strip(),
            "addr1": self.v_addr1.get().strip(),
            "addr2": self.v_addr2.get().strip(),
            "city":  self.v_city.get().strip(),
            "state": self.v_state.get().strip(),
            "zip":   self.v_zip.get().strip(),
            "dob":   self.v_dob.get().strip(),
        }
        self.destroy()

# -------------------- Client dialog (Add/Edit) with scroll --------------------
class ClientDialog(tk.Toplevel):
    def __init__(self, master, title: str, initial: Optional[Dict[str, Any]] = None):
        super().__init__(master)
        self.title(title); self.resizable(True, True); self.result=None
        self.minsize(880, 600)
        self.maxsize(880, 1200)

        self._initial = dict(initial or {})
        init = self._initial
        self.v_name    = tk.StringVar(value=init.get("name",""))
        self.v_dba     = tk.StringVar(value=init.get("dba",""))
        self.v_entity  = tk.StringVar(value=init.get("entity_type",""))
        self.v_ein     = tk.StringVar(value=init.get("ein",""))
        self.v_edd     = tk.StringVar(value=init.get("edd_number",""))
        self.v_sales   = tk.StringVar(value=init.get("sales_tax_account",""))
        self.v_acctmgr = tk.StringVar(value=init.get("acct_mgr",""))
        self.v_path    = tk.StringVar(value=init.get("file_location",""))
        self.v_addr1 = tk.StringVar(value=init.get("addr1",""))
        self.v_addr2 = tk.StringVar(value=init.get("addr2",""))
        self.v_city  = tk.StringVar(value=init.get("city",""))
        self.v_state = tk.StringVar(value=init.get("state",""))
        self.v_zip   = tk.StringVar(value=init.get("zip",""))
        self.v_ui_rate   = tk.StringVar(value=init.get("ui_rate",""))
        self.v_stx_rate  = tk.StringVar(value=init.get("sales_tax_rate",""))
        self.v_other_tax = tk.StringVar(value=init.get("other_tax_rates",""))

        self.memo_init = init.get("memo","")

        sf = ScrollFrame(self, height=520); sf.pack(fill="both", expand=True)
        frm = sf.inner

        def row(lbl, var, r, c=0, width=52, colspan=1):
            ttk.Label(frm, text=lbl).grid(row=r, column=c, sticky="w", pady=(0,2))
            ent = ttk.Entry(frm, textvariable=var, width=width)
            ent.grid(row=r+1, column=c, sticky="we", pady=(0,6), columnspan=colspan); return ent

        e1 = row("Name", self.v_name, 0, 0, 40, 2)
        row("DBA", self.v_dba, 0, 2, 30, 2)

        ttk.Label(frm, text="Entity Type").grid(row=2, column=0, sticky="w", pady=(0,2))
        cb_entity = ttk.Combobox(frm, values=ENTITY_TYPES, textvariable=self.v_entity, width=36)
        cb_entity.grid(row=3, column=0, sticky="we", pady=(0,6))

        row("EIN (e.g., 12-3456789)", self.v_ein, 2, 1, 26)
        row("EDD Number", self.v_edd, 2, 2, 20)
        row("Sales Tax Account", self.v_sales, 2, 3, 20)

        ttk.Label(frm, text="Account Manager").grid(row=4, column=0, sticky="w", pady=(0,2))
        mgr_cb, self.v_acctmgr = self.master._make_mgr_combobox(frm, initial_name=self.v_acctmgr.get())
        mgr_cb.grid(row=5, column=0, sticky="we", pady=(0,6), columnspan=2)

        row("File Location (folder path or URL)", self.v_path, 4, 2, 40, 2)

        ttk.Label(frm, text="Company Address").grid(row=6, column=0, sticky="w", pady=(8,2))
        row("Address 1", self.v_addr1, 7, 0, 50, 4)
        row("Address 2", self.v_addr2, 9, 0, 50, 4)

        ttk.Label(frm, text="City").grid(row=11, column=0, sticky="w", pady=(0,2))
        ent_city = ttk.Entry(frm, textvariable=self.v_city, width=32)
        ent_city.grid(row=12, column=0, sticky="we", pady=(0,6), columnspan=2)

        ttk.Label(frm, text="State").grid(row=11, column=2, sticky="w", pady=(0,2))
        cb_state = ttk.Combobox(frm, values=US_STATES, textvariable=self.v_state, width=10)
        cb_state.grid(row=12, column=2, sticky="we", pady=(0,6))

        ttk.Label(frm, text="Zip").grid(row=11, column=3, sticky="w", pady=(0,2))
        ent_zip = ttk.Entry(frm, textvariable=self.v_zip, width=14)
        ent_zip.grid(row=12, column=3, sticky="we", pady=(0,6))

        ttk.Label(frm, text="Tax Rates").grid(row=14, column=0, sticky="w", pady=(8,2))
        row("UI Rate (%)", self.v_ui_rate, 15, 0, 12)
        row("Sales Tax Rate (%)", self.v_stx_rate, 15, 1, 16)
        row("Other Tax Rates (text)", self.v_other_tax, 15, 2, 36, 2)

        ttk.Label(frm, text="Officers / Personnel").grid(row=17, column=0, sticky="w", pady=(6,2))
        self.off_cols = ("name","first_name","middle_name","last_name","nickname",
                         "email","phone","addr1","addr2","city","state","zip","dob")
        self.off_tree = ttk.Treeview(frm, columns=self.off_cols, show="headings", height=4, selectmode="browse")

        for col, label, w in (("name","Person",220),("email","Email",220),("phone","Phone",120)):
            self.off_tree.heading(col, text=label); self.off_tree.column(col, width=w, anchor="w")
        for col in ("first_name","middle_name","last_name","nickname","addr1","addr2","city","state","zip","dob"):
            self.off_tree.heading(col, text=col); self.off_tree.column(col, width=0, stretch=False)

        self.off_tree.grid(row=18, column=0, columnspan=4, sticky="nsew")
        off_btns = ttk.Frame(frm); off_btns.grid(row=19, column=0, sticky="w", pady=(4,6), columnspan=4)
        ttk.Button(off_btns, text="Add", command=self._off_add).pack(side=tk.LEFT)
        ttk.Button(off_btns, text="Edit", command=self._off_edit).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(off_btns, text="Remove", command=self._off_remove).pack(side=tk.LEFT, padx=(6,0))

        ttk.Label(frm, text="Memo").grid(row=21, column=0, sticky="w", pady=(4,2))
        self.memo_txt = ScrolledText(frm, width=72, height=6, wrap="word")
        self.memo_txt.grid(row=22, column=0, columnspan=4, sticky="we", pady=(0,6))
        if self.memo_init: self.memo_txt.insert("1.0", self.memo_init)

        for o in init.get("officers", []):
            o = ensure_officer_dict(o)
            vals = (
                display_officer_name(o), o["first_name"], o["middle_name"], o["last_name"], o["nickname"],
                o["email"], o["phone"], o["addr1"], o["addr2"], o["city"], o["state"], o["zip"], o["dob"]
            )
            self.off_tree.insert("", "end", values=vals)

        btns = ttk.Frame(self); btns.pack(fill="x", pady=(6,6))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda _e: self._save()); self.bind("<Escape>", lambda _e: self.destroy())
        self.after(50, e1.focus_set); self.grab_set(); self.transient(master)

    def _off_add(self):
        d = OfficerDialog(self, "Add Personnel"); self.wait_window(d)
        if d.result:
            o = ensure_officer_dict(d.result)
            vals = (
                display_officer_name(o), o["first_name"], o["middle_name"], o["last_name"], o["nickname"],
                o["email"], o["phone"], o["addr1"], o["addr2"], o["city"], o["state"], o["zip"], o["dob"]
            )
            self.off_tree.insert("", "end", values=vals)

    def _off_edit(self):
        sel = self.off_tree.selection()
        if not sel:
            messagebox.showinfo("Edit", "Select a personnel row to edit."); return
        v = self.off_tree.item(sel[0], "values")
        init = {
            "name":        v[0],
            "first_name":  v[1], "middle_name": v[2], "last_name": v[3], "nickname": v[4],
            "email":       v[5], "phone": v[6],
            "addr1":       v[7], "addr2": v[8], "city": v[9], "state": v[10], "zip": v[11],
            "dob":         v[12],
        }
        d = OfficerDialog(self, "Edit Personnel", init); self.wait_window(d)
        if d.result:
            o = ensure_officer_dict(d.result)
            new_vals = (
                display_officer_name(o), o["first_name"], o["middle_name"], o["last_name"], o["nickname"],
                o["email"], o["phone"], o["addr1"], o["addr2"], o["city"], o["state"], o["zip"], o["dob"]
            )
            self.off_tree.item(sel[0], values=new_vals)

    def _off_remove(self):
        sel = self.off_tree.selection()
        if not sel:
            messagebox.showinfo("Remove", "Select a personnel row to remove."); return
        self.off_tree.delete(sel[0])

    def _gather_officers(self) -> List[Dict[str,str]]:
        offs = []
        for iid in self.off_tree.get_children():
            v = self.off_tree.item(iid, "values")
            offs.append({
                "name":        v[0],
                "first_name":  v[1], "middle_name": v[2], "last_name": v[3], "nickname": v[4],
                "email":       v[5], "phone": v[6],
                "addr1":       v[7], "addr2": v[8], "city": v[9], "state": v[10], "zip": v[11],
                "dob":         v[12],
            })
        return offs

    def _save(self):
        name = self.v_name.get().strip()
        if not name:
            messagebox.showerror("Required", "Name is required."); return

        officers = self._gather_officers()
        memo     = self.memo_txt.get("1.0", "end").strip()

        self.result = {
            "name": name,
            "dba": self.v_dba.get().strip(),
            "entity_type": self.v_entity.get().strip(),
            "ein": self.v_ein.get().strip(),
            "edd_number": self.v_edd.get().strip(),
            "sales_tax_account": self.v_sales.get().strip(),
            "acct_mgr": self.v_acctmgr.get().strip(),
            "file_location": self.v_path.get().strip(),

            "addr1": self.v_addr1.get().strip(),
            "addr2": self.v_addr2.get().strip(),
            "city":  self.v_city.get().strip(),
            "state": self.v_state.get().strip(),
            "zip":   self.v_zip.get().strip(),

            "ui_rate": self.v_ui_rate.get().strip(),
            "sales_tax_rate": self.v_stx_rate.get().strip(),
            "other_tax_rates": self.v_other_tax.get().strip(),
            "tax_rates_last_checked": date.today().isoformat(),

            "officers": officers,
            "employees": self._initial.get("employees", []),  # keep if existed
            "memo": memo,
            "logs": self._initial.get("logs", []),
        }
        self.destroy()

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
            "officers",
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

        self.items: List[Dict[str, Any]] = load_clients()

        self.style = ttk.Style()
        try: self.style.theme_use("clam")
        except Exception: pass
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.base_row_px = int(self.default_font.metrics("linespace") * 1.35)

        style = ttk.Style()
        style.configure("Treeview", rowheight=40)
        self._last_viewed_idx = None


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

            # New Company menu item
            on_new_company=self.on_new,

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
            on_export_data=self._export_data_dialog,

            # Update
            on_check_updates=self._check_for_updates,
            on_about=self._show_about,
        )

        # Top bar
        topbar = ttk.Frame(self, style="Card.TFrame", padding=(10, 8)); topbar.pack(side=tk.TOP, fill=tk.X)
        self.btn_back = ttk.Button(topbar, text="◀ Back", command=self.nav_back, state=tk.DISABLED, width=10, style="NewUI.TButton")
        self.btn_fwd  = ttk.Button(topbar, text="Forward ▶", command=self.nav_forward, state=tk.DISABLED, width=10, style="NewUI.TButton")
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

        # Main page
        self._build_main_page()

        # Autocomplete
        self._ac = AutocompletePopup(self, self.search_entry, on_choose=self._open_from_suggestion)

        # Hide suggestions when clicking anywhere outside the popup or the entry
        def _dismiss_ac_on_click(e):
            if not self._ac or not self._ac.winfo_exists() or not self._ac.winfo_viewable():
                return
            w = e.widget
            # Is the click inside the popup?
            in_popup = str(w).startswith(str(self._ac))
            # Is the click on the entry (keep it open if the user clicks back into the entry)?
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
        self._mgr_filter_active = set()   # set of selected manager names (exact-case)
        self._mgr_menu = None

    # Updating Software    
    def _check_for_updates(self):
        check_for_updates(self.winfo_toplevel())

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

    def _account_managers_path(self) -> str:
        return os.path.join(self._clients_folder(), "account_managers.json")
    
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
        # Allow typing-to-search without committing invalid text
        cb.configure(state="normal")
        cb["values"] = self._account_manager_names()
        def _validate_and_lock(event=None):
            names = self._account_manager_names()
            cur = (v.get() or "").strip()
            if cur:
                # If not already in roster, add it and SAVE immediately
                exists_ci = any(n.casefold() == cur.casefold() for n in names)
                if not exists_ci:
                    # Append to in-memory roster
                    am = getattr(self, "account_managers", []) or []
                    am.append({"name": cur})
                    self.account_managers = self._normalize_acct_mgr_list(am)
                    # WRITE TO clients/account_managers.json now
                    self._save_account_managers(self.account_managers)
                    # Refresh dropdown values everywhere
                    cb["values"] = self._account_manager_names()
                    # If you have a Managers ▾ filter menu open, its labels can be rebuilt on next open
                    # (optional) self._refresh_mgr_menu_checks()
                else:
                    # Snap to the canonical cased name
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
        Ensure list of dicts with keys: name, email, phone.
        Accepts list[str] or list[dict].
        """
        out = []
        for x in (items or []):
            if isinstance(x, dict):
                out.append({
                    "name":  str(x.get("name","")).strip(),
                    "email": str(x.get("email","")).strip(),
                    "phone": str(x.get("phone","")).strip(),
                })
            else:
                s = str(x).strip()
                if s:
                    out.append({"name": s, "email": "", "phone": ""})
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
                # optional status line update if you have self.status (StringVar)
                if hasattr(self, "status"):
                    self.status.set(f"Saved {len(lst)} account manager(s) → {path}")
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Account Managers", f"Failed to save:\n{e}")

            
    def _mgr_filter_names_ci(self):
        """Case-insensitive set for matching."""
        return { (n or "").casefold() for n in self._mgr_filter_active }

    def _toggle_mgr_filter(self, name: str | None):
        """
        name == None  -> toggle 'All'
        else          -> toggle a specific manager name
        """
        if name is None:
            # 'All' means clear specific selections
            if self._mgr_filter_active:
                self._mgr_filter_active.clear()
            else:
                # If nothing is selected, selecting All is also a no-op (empty set = All)
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
            elif c in ("emails","officers"):
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

    # ---------- Company Detail (Notebook) ----------
    def _build_detail_page(self, idx: int):
        for w in self.page_host.winfo_children():
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
        
        ttk.Button(actions, text="Edit Company", style="NewUI.Accent.TButton",
                   command=lambda i=idx: self._detail_edit(i)).pack(side="top", fill="x")
        ttk.Button(actions, text="Delete Company", style="NewUI.Accent.TButton",
                   command=lambda i=idx: self._delete_company(i)).pack(side="top", fill="x", pady=(6,0))
        ttk.Button(actions, text="Open File Location", style="NewUI.Accent.TButton",
                   command=lambda: self.open_path(c.get("file_location",""))).pack(side="top", fill="x", pady=(6,0))
        
        # Column sizing so name (col 0) hugs content, middle info (col 1) expands, actions (col 2) stays tight.
        header.grid_columnconfigure(0, weight=0)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=0)

        page.bind_all("<Control-e>", lambda e, i=idx: self._detail_edit(i))


        nb = ttk.Notebook(page); nb.pack(fill="both", expand=True, padx=6, pady=6)

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
        ttk.Label(ana, text=f"Officers: {len(c.get('officers',[]))}").pack(anchor="w")
        ttk.Label(ana, text=f"Employees: {len(c.get('employees',[]))}").pack(anchor="w")
        ttk.Label(ana, text=f"Has EIN: {'Yes' if c.get('ein') else 'No'}").pack(anchor="w")
        ttk.Label(ana, text=f"Entity Type: {c.get('entity_type','') or '—'}").pack(anchor="w")

        footer = ttk.Frame(page, padding=(8,2)); footer.pack(side=tk.BOTTOM, fill=tk.X)

        self._current_detail_idx = idx
        self._current_page = ("detail", idx)

    # ---------- Personnel Detail Page ----------
    def _build_person_page(self, company_idx: int, role_key: str, person_idx: int):
        for w in self.page_host.winfo_children():
            self._clear_page_host()

        c = self.items[company_idx]
        people = c.get(role_key, [])
        if person_idx < 0 or person_idx >= len(people):
            self.navigate("detail", company_idx, replace=True); return
        p = ensure_officer_dict(people[person_idx])

        page = ttk.Frame(self.page_host)
        page.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(page, padding=(8,8,8,4)); header.pack(side=tk.TOP, fill=tk.X)

        title_box = ttk.Frame(header); title_box.pack(side=tk.LEFT, padx=12, fill=tk.X, expand=True)
        ttk.Label(title_box, text=p.get("name","") or "(Person)", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(title_box, text=f"{'Officer' if role_key=='officers' else 'Employee'} — {c.get('name','')}", foreground="#555").pack(anchor="w", pady=(2,0))

        actions = ttk.Frame(header); actions.pack(side=tk.RIGHT)
        ttk.Button(actions, text="Edit", command=lambda: self._edit_person(company_idx, role_key, person_idx)).pack(side=tk.TOP, fill=tk.X)
        ttk.Button(actions, text="Back to Company", command=lambda: self.navigate('detail', company_idx, push=True)).pack(side=tk.TOP, fill=tk.X, pady=(6,0))

        nb = ttk.Notebook(page); nb.pack(fill="both", expand=True, padx=6, pady=6)

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

        self._current_page = ("person", (company_idx, role_key, person_idx))

    def _edit_person(self, company_idx: int, role_key: str, person_idx: int):
        c = self.items[company_idx]
        p = ensure_officer_dict(c.get(role_key, [])[person_idx])
        d = OfficerDialog(self, f"Edit {'Officer' if role_key=='officers' else 'Employee'}", p)
        self.wait_window(d)
        if d.result:
            c[role_key][person_idx] = ensure_officer_dict(d.result)
            save_clients(self.items)
            self.navigate("person", (company_idx, role_key, person_idx), replace=True)

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
        """Import data from a JSON file and merge in any new clients (by EIN)."""
        path_str = filedialog.askopenfilename(
            title="Import Data (JSON)",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path_str:
            return

        path = Path(path_str)
        self.log.info("Importing data from %s", path)
        added_clients = import_all_from_json(path, self.items)
        if added_clients:
            save_clients(self.items)
            self.populate()  # refresh search/dashboard
        if hasattr(self, "status"):
            self.status.set(f"Imported {added_clients} new client(s) from {path.name}")

    def _export_all_dialog(self):
        """Export ALL data (currently clients) to a JSON file chosen by the user."""
        path_str = filedialog.asksaveasfilename(
            title="Export All Data (JSON)",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path_str:
            return

        path = Path(path_str)
        self.log.info("Exporting all data to %s", path)
        export_all_to_json(path, self.items)
        if hasattr(self, "status"):
            self.status.set(f"Exported all data to {path}")
            
    def _export_data_dialog(self):
        """
        Compatibility wrapper for TaskbarModel's on_export_data hook.
        Currently just calls _export_all_dialog().
        """
        return self._export_all_dialog()

    # ---------- Misc ----------
    def _show_context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.menu.tk_popup(event.x_root, event.y_root)

    def _set_row_height_for_items(self, items: List[Dict[str, Any]]):
        max_lines = 1
        for c in items:
            officers_lines = max(1, len(officers_to_display_lines(c.get("officers", []))))
            emails_lines   = max(1, len(officers_to_flat_emails(c.get("officers", []))) or 1)
            phones_lines   = max(1, len(officers_to_flat_phones(c.get("officers",[]))) or 1)
            max_lines = max(max_lines, officers_lines, emails_lines, phones_lines)
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

        if not q:
            base = list(self.items)
        else:
            q_norm = norm_text(q)
            q_digits = "".join(PHONE_DIGITS_RE.findall(q))
            last10 = q_digits[-10:] if len(q_digits) >= 4 else q_digits
            last9  = q_digits[-9:]  if len(q_digits) >= 4 else q_digits

            res = []
            for c in self.items:
                officer_phones = officers_to_flat_phones(c.get("officers",[]))
                phones_digits_full = ["".join(PHONE_DIGITS_RE.findall(p or "")) for p in officer_phones]
                phones_norm_last10 = [normalize_phone_digits(p) for p in officer_phones]
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
                    " ".join(norm_text(ensure_officer_dict(o).get("name","")) for o in c.get("officers",[])),
                    " ".join(norm_text(ensure_officer_dict(o).get("first_name","")) for o in c.get("officers",[])),
                    " ".join(norm_text(ensure_officer_dict(o).get("last_name","")) for o in c.get("officers",[])),
                    " ".join(norm_text(ensure_officer_dict(o).get("nickname","")) for o in c.get("officers",[])),
                    " ".join(norm_text(ensure_officer_dict(o).get("email","")) for o in c.get("officers",[])),
                    norm_text(c.get("file_location","")),
                    norm_text(c.get("memo","")),
                ])
                text_hit = all(tok in hay for tok in q_norm.split()) if q_norm else False

                if phone_hit or ein_hit or text_hit:
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
            officer_lines = officers_to_display_lines(c.get("officers",[]))
            emails_lines  = officers_to_flat_emails(c.get("officers",[]))
            phones_lines  = officers_to_flat_phones(c.get("officers",[]))
            vals = (
                c.get("name",""),
                c.get("dba",""),
                "\n".join(officer_lines),
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
    
    def _focus_company_in_search(self, key: str):
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

    def on_new(self):
        dlg = ClientDialog(self, "New Client"); self.wait_window(dlg)
        if dlg.result:
            self.items.append(dlg.result); save_clients(self.items)
            self.populate(); self._update_suggestions()

    def on_edit(self):
        idx = getattr(self, "_current_detail_idx", None)
        if self._current_page[0] == "main" or idx is None:
            idx = self.selected_index()
            if idx is None:
                messagebox.showinfo("Edit", "Select a row to edit."); return
        dlg = ClientDialog(self, "Edit Client", self.items[idx]); self.wait_window(dlg)
        if dlg.result:
            self.items[idx] = dlg.result; save_clients(self.items)
            self.populate(); self._update_suggestions()
            if self._current_page[0] != "main":
                self.navigate(self._current_page[0], self._current_page[1], replace=True)

    def _detail_edit(self, idx: int):
        self.on_edit()

    def on_delete(self):
        self._delete_company()

    # ---------- Actions ----------
    def copy_emails(self):
        items = self.filtered_items()
        seen, acc = set(), []
        for c in items:
            for e in officers_to_flat_emails(c.get("officers",[])):
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

    def _delete_company(self, idx: int | None = None):
        if idx is None:
            if getattr(self, "_current_page", ("main", None))[0] != "main":
                idx = getattr(self, "_current_detail_idx", None)
            if idx is None:
                idx = self.selected_index()

        if idx is None or idx < 0 or idx >= len(self.items):
            messagebox.showinfo("Delete", "Select a company to delete first.")
            return

        c = self.items[idx]
        if not messagebox.askyesno(
            "Confirm Delete",
            f"Delete this company?\n\n{c.get('name','')}  (DBA: {c.get('dba','') or '—'})"
        ):
            return

        del self.items[idx]
        save_clients(self.items)

        # After deletion, return to main and refresh
        self.navigate("search", None, push=False, replace=True)
        self.populate()
        self._update_suggestions()
        self.status.set("Company deleted.")

    def _on_delete_from_tree(self, event):
        tv = event.widget
        if not isinstance(tv, ttk.Treeview):
            return
        sel = tv.selection()
        if not sel:
            return
        idx = tv.index(sel[0])
        self._delete_company_at_index(idx, tv)

    def _delete_company_at_index(self, idx, tv):
        pass

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
        elif page == "taxes":
            if self._last_viewed_idx is not None and 0 <= self._last_viewed_idx < len(self.items):
                c = self.items[self._last_viewed_idx]
                key = c.get("ein") or c.get("name") or c.get("dba")
                if key and hasattr(self, "checklist") and hasattr(self.checklist, "focus_company_by_key"):
                    self.checklist.focus_company_by_key(key) 
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
                self._focus_company_in_search(key)
                self._return_focus_key = None

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

    def nav_back(self):
        self.log.info("nav_back()")
        if not getattr(self, "_history", None):
            return
        prev = self._history.pop()
        self._future.append(self._current_page)
        self._current_page = prev

        self._clear_page_host()

        kind, data = prev
        if kind == "main":
            self._ensure_main_page()
            self.page_main.pack(fill=tk.BOTH, expand=True)

        elif kind == "search":
            self._ensure_search_page()
            self.page_search.pack(fill=tk.BOTH, expand=True)
            self.populate()
            key = getattr(self, "_return_focus_key", None)
            if key:
                self._focus_company_in_search(key)
                # clear it so it doesn’t keep re-firing
                self._return_focus_key = None

        elif kind == "taxes":
            ChecklistPage(app=self).ensure(self.page_host)
        elif kind == "detail":
            self._build_detail_page(data)
        else:
            if self._is_valid_person_payload(data):
                ci, role_key, pidx = data
                self._build_person_page(int(ci), str(role_key), int(pidx))
            else:
                self._ensure_main_page()
                self.page_main.pack(fill=tk.BOTH, expand=True)
        self._update_nav_buttons()

    def nav_forward(self):
        self.log.info("nav_forward()")
        if not getattr(self, "_future", None):
            return
        if not self._future:
            return
        nxt = self._future.pop()
        self._history.append(self._current_page)
        self._current_page = nxt

        self._clear_page_host()

        kind, data = nxt
        if kind == "main":
            self._ensure_main_page()
            self.page_main.pack(fill=tk.BOTH, expand=True)

        elif kind == "search":
            self._ensure_search_page()
            self.page_search.pack(fill=tk.BOTH, expand=True)
            self.populate()
        elif kind == "taxes":
            ChecklistPage(app=self).ensure(self.page_host)
        elif kind == "detail":
            self._build_detail_page(data)
        else:
            if self._is_valid_person_payload(data):
                ci, role_key, pidx = data
                self._build_person_page(int(ci), str(role_key), int(pidx))
            else:	
                # Stale/bad entry in history → safe fallback
                self._ensure_main_page()
                self.page_main.pack(fill=tk.BOTH, expand=True)
        self._update_nav_buttons()

    def _update_nav_buttons(self):
        self.btn_back["state"] = tk.NORMAL if bool(getattr(self, "_history", [])) else tk.DISABLED
        self.btn_fwd["state"]  = tk.NORMAL if bool(getattr(self, "_future", [])) else tk.DISABLED

# -------------------- Entrypoint --------------------
def main():
    log = get_logger("launcher")
    log.info("Launching Client Manager main()")
    root = tk.Tk()
    root.title("Client Manager")
    root.geometry("1280x800")
    root.minsize(1920, 720)

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
    main()
