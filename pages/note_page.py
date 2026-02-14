from __future__ import annotations

from typing import Dict, Any, List, Tuple
from datetime import datetime
import re
import tkinter as tk
from tkinter import ttk

try:
    from vertex.utils.app_logging import get_logger
except ModuleNotFoundError:
    from utils.app_logging import get_logger

LOG = get_logger("note_page")

# Keep tokenization consistent with client_manager.py, but self-contained
def tokenize(s: str) -> List[str]:
    s = (s or "").strip().lower()
    # Keep letters/digits plus common business-name symbols.
    # IMPORTANT: put '-' as '\-' or at the END to avoid "bad character range" errors.
    parts = re.split(r"[^a-z0-9@._&\-]+", s)
    return [p for p in parts if p]

def norm_text(s: str) -> str:
    return " ".join(tokenize(s))

def _parse_dt(val: Any) -> datetime:
    if isinstance(val, datetime):
        return val
    if not val:
        return datetime.min
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y-%m-%dT%H:%M:%S",
    ):
        try:
            return datetime.strptime(str(val), fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(str(val).replace("Z", "+00:00"))
    except Exception:
        return datetime.min

class NotePage:
    """
    Global Notes page.
    Shows ALL client logs across all companies.
    Newest first.
    Double-click:
      - on ✓ column: toggle done
      - elsewhere: open client and focus Logs tab
    """

    def __init__(self, app):
        self.app = app
        self.page = None
        self.tree = None

        self._iid_to_row: Dict[str, Tuple[int, int]] = {}  # iid -> (client_idx, log_idx)

        self.q = tk.StringVar()
        self.mgr = tk.StringVar(value="All")
        self.done = tk.StringVar(value="All")  # All | Open | Done

    def ensure(self, parent: tk.Widget):
        if self.page and self.page.winfo_exists():
            self.refresh()
            return self.page

        self.page = ttk.Frame(parent)

        # Header
        header = ttk.Frame(self.page, padding=(10, 8))
        header.pack(fill=tk.X)

        ttk.Label(header, text="Notes", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT)

        ttk.Label(header, text="Manager:").pack(side=tk.LEFT, padx=(14, 4))
        self.cmb_mgr = ttk.Combobox(header, textvariable=self.mgr, width=18, state="readonly")
        self.cmb_mgr.pack(side=tk.LEFT)

        ttk.Label(header, text="Status:").pack(side=tk.LEFT, padx=(10, 4))
        self.cmb_done = ttk.Combobox(header, textvariable=self.done, width=10, state="readonly",
                                     values=("All", "Open", "Done"))
        self.cmb_done.pack(side=tk.LEFT)

        ttk.Label(header, text="Search").pack(side=tk.RIGHT, padx=(6, 4))
        ent = ttk.Entry(header, textvariable=self.q, width=42)
        ent.pack(side=tk.RIGHT)
        ent.bind("<KeyRelease>", lambda _e: self.refresh())

        self.cmb_mgr.bind("<<ComboboxSelected>>", lambda _e: self.refresh())
        self.cmb_done.bind("<<ComboboxSelected>>", lambda _e: self.refresh())

        # Tree (NO user column) with scrollbar
        cols = ("done", "date", "client", "manager", "note")
        tree_frm = ttk.Frame(self.page)
        tree_frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        tree_frm.columnconfigure(0, weight=1)
        tree_frm.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_frm, columns=cols, show="headings", height=22)
        tree_vsb = ttk.Scrollbar(tree_frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_vsb.grid(row=0, column=1, sticky="ns")

        self.tree.heading("done", text="✓")
        self.tree.heading("date", text="Date")
        self.tree.heading("client", text="Company")
        self.tree.heading("manager", text="Manager")
        self.tree.heading("note", text="Note")

        self.tree.column("done", width=44, anchor="center", stretch=False)
        self.tree.column("date", width=170, anchor="w", stretch=False)
        self.tree.column("client", width=260, anchor="w", stretch=True)
        self.tree.column("manager", width=140, anchor="w", stretch=False)
        self.tree.column("note", width=900, anchor="w", stretch=True)

        # double click behavior (toggle if done-col, otherwise open)
        self.tree.bind("<Double-1>", self._on_double_click)

        self._rebuild_mgr_values()
        self.refresh()
        return self.page

    def _parse_ts(self, ts: str) -> datetime:
        ts = (ts or "").strip()
        if not ts:
            return datetime.min
        try:
            # ISO like "2025-12-15T08:12:33"
            return datetime.fromisoformat(ts.replace("Z", "+00:00"))
        except Exception:
            pass
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(ts, fmt)
            except Exception:
                pass
        return datetime.min

    def _flatten_logs(self):
        rows = []
        for client_idx, c in enumerate(getattr(self.app, "items", []) or []):
            cname = c.get("name", "")
            mgr = c.get("acct_mgr", "")
            for log_idx, lg in enumerate(c.get("logs", []) or []):
                ts = str(lg.get("ts", "")).strip()
                rows.append({
                    "client_idx": client_idx,
                    "log_idx": log_idx,
                    "date": ts,
                    "_dt": self._parse_ts(ts),
                    "client": cname,
                    "manager": mgr,
                    "text": lg.get("text", ""),
                    "done": bool(lg.get("done")),
                })

        rows.sort(key=lambda r: r["_dt"] or datetime.min, reverse=True)
        return rows

    def _rebuild_mgr_values(self):
        mgrs = []
        for c in getattr(self.app, "items", []) or []:
            m = (c.get("acct_mgr", "") or "").strip()
            if m:
                mgrs.append(m)
        mgrs = sorted(set(mgrs), key=lambda s: s.casefold())
        values = ["All"] + mgrs
        try:
            self.cmb_mgr.configure(values=values)
        except Exception:
            pass
        if self.mgr.get() not in values:
            self.mgr.set("All")

    def refresh(self):
        if not self.tree:
            return

        self._rebuild_mgr_values()

        q_norm = norm_text(self.q.get())
        mgr = self.mgr.get()
        done = self.done.get()

        self.tree.delete(*self.tree.get_children())
        self._iid_to_row.clear()

        for r in self._flatten_logs():
            if mgr != "All" and r["manager"] != mgr:
                continue
            if done == "Done" and not r["done"]:
                continue
            if done == "Open" and r["done"]:
                continue

            if q_norm:
                hay = " ".join([
                    norm_text(r["client"]),
                    norm_text(r["manager"]),
                    norm_text(r["text"]),
                    norm_text(r["date"]),
                ])
                # AND filtering: every token must match
                if not all(tok in hay for tok in q_norm.split()):
                    continue

            iid = self.tree.insert(
                "",
                "end",
                values=(
                    "✓" if r["done"] else "",
                    r["date"],
                    r["client"],
                    r["manager"],
                    r["text"],
                ),
            )
            self._iid_to_row[iid] = (r["client_idx"], r["log_idx"])

    def _on_double_click(self, e):
        if not self.tree:
            return
        iid = self.tree.identify_row(e.y)
        if not iid:
            return

        col = self.tree.identify_column(e.x)  # "#1" is first displayed column
        if col == "#1":
            self._toggle_done(iid)
        else:
            self._open_log(iid)

    def _toggle_done(self, iid: str):
        row = self._iid_to_row.get(iid)
        if not row:
            return
        client_idx, log_idx = row

        try:
            lg = self.app.items[client_idx]["logs"][log_idx]
        except Exception:
            return

        lg["done"] = not bool(lg.get("done"))

        # persist if your app has helper
        if hasattr(self.app, "save_clients_data"):
            try:
                self.app.save_clients_data()
            except Exception:
                pass

        self.refresh()

    def _open_log(self, iid: str):
        row = self._iid_to_row.get(iid)
        if not row:
            return
        client_idx, _log_idx = row

        # IMPORTANT: positional arg for data (your navigate() expects data, not idx=)
        self.app.navigate("detail", client_idx, push=True)

        # try to focus Logs tab if available
        if hasattr(self.app, "select_detail_tab"):
            try:
                self.app.select_detail_tab("Logs")
            except Exception:
                pass
