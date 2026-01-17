from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))
        
import json
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import re, unicodedata

try:
    from functions.utils.app_logging import get_logger
except ModuleNotFoundError:
    from utils.app_logging import get_logger
LOG = get_logger("checklist")

try:
    from styles.new_ui import NewUI
except Exception:
    class NewUI:
        @staticmethod
        def install(root): pass
        @staticmethod
        def stripe_tree(tree): pass

# ----- Columns -----
client_COL = "client"
COLS = [
    "Jan","Feb","Mar","Q1",
    "Apr","May","Jun","Q2",
    "Jul","Aug","Sep","Q3",
    "Oct","Nov","Dec","Q4",
]
ALL_COLS = [client_COL] + COLS

CELL_OPTIONS = ["--", "To do", "Done"]

_LEGACY_MAP = {
    "To do ⏳": "To do",
    "Done ✅": "Done",
}

def _store_to_ui(v: str) -> str:
    v = (v or "").strip()
    v = _LEGACY_MAP.get(v, v)
    return "--" if v == "" else (v if v in ("To do", "Done") else "--")

def _ui_to_store(v: str) -> str:
    v = (v or "").strip()
    if v in ("--", ""): return ""
    if v in ("To do", "To do ⏳"): return "To do"
    if v in ("Done", "Done ✅"):   return "Done"
    return ""


# Keys used inside monthly_state.json (one section per tab)
KEY_MONTHLY_TAX = "checklist_monthly_tax"
KEY_SALES_TAX   = "checklist_sales_tax"
KEY_FINSTAT     = "checklist_financial_statement"

def _paths():
    """
    Returns:
      clients_path: functions/data/clients/clients.json
      state_path:   functions/data/monthly_data/monthly_state.json
    """
    here = Path(__file__).resolve()
    pages_dir     = here.parent
    functions_dir = pages_dir.parent
    data_dir      = functions_dir / "data"

    clients_path = data_dir / "clients" / "clients.json"
    monthly_dir  = data_dir / "monthly_data"
    monthly_dir.mkdir(parents=True, exist_ok=True)
    state_path = monthly_dir / "monthly_state.json"
    return clients_path, state_path


def _load_clients():
    """
    Read clients.json and return list of (label, key) where:
      label = 'DBA — Name' if DBA exists else Name (fallback to EIN if name empty)
      key   = EIN if present else Name (fallback to DBA)
    """
    clients_path, _ = _paths()
    try:
        raw = json.loads(clients_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        messagebox.showwarning("Missing clients.json",
                               f"Couldn't find:\n{clients_path}\n\nProceeding with an empty list.")
        return []
    except Exception as e:
        messagebox.showerror("Failed to parse clients.json",
                             f"Path: {clients_path}\n\n{e}")
        return []

    out = []
    for rec in (raw or []):
        if not isinstance(rec, dict):
            continue
        name = (rec.get("name") or "").strip()
        dba  = (rec.get("dba")  or "").strip()
        ein  = (rec.get("ein")  or "").strip()
        if not name and not dba and not ein:
            continue
        label = f"{dba} — {name}" if dba else (name or ein or "(Unnamed)")
        key = ein or name or dba
        if key:
            out.append((label, key))
    out.sort(key=lambda t: t[0].lower())
    return out


def _read_state_blob():
    """Read the whole monthly_state.json blob (dict)."""
    _, state_path = _paths()
    if not state_path.exists():
        return {}
    try:
        return json.loads(state_path.read_text(encoding="utf-8")) or {}
    except Exception:
        return {}


def _write_state_blob(blob: dict):
    """Write the whole monthly_state.json blob (dict)."""
    _, state_path = _paths()
    try:
        state_path.write_text(json.dumps(blob, indent=2), encoding="utf-8")
        return True, None
    except Exception as e:
        return False, e

def _load_tab_year_state(section_key: str, year: int):
    blob = _read_state_blob()
    section = blob.get(section_key, {})
    years = section.get("years", {})
    ys = years.get(str(year), {})
    companies = ys.get("companies", [])
    cells = ys.get("cells", {})

    if not isinstance(companies, list): companies = []
    if not isinstance(cells, dict): cells = {}

    # Make sure each client row has all columns and normalize for UI
    for k in companies:
        row = cells.setdefault(k, {})
        for c in COLS:
            row[c] = _store_to_ui(row.get(c, ""))

    return companies, cells


def _save_tab_year_state(section_key: str, year: int, companies: list[str], cells: dict):
    """Persist current year state. Converts UI labels to JSON values."""
    blob = _read_state_blob()
    section = blob.get(section_key, {})
    years = section.get("years", {})

    # Normalize cells to storage form
    norm_cells = {}
    for k in companies:
        src = cells.get(k, {})
        norm_cells[k] = {c: _ui_to_store(src.get(c, "")) for c in COLS}

    years[str(year)] = {"companies": list(companies), "cells": norm_cells}
    section["years"] = years
    blob[section_key] = section
    return _write_state_blob(blob)

class _ChecklistTab(ttk.Frame):
    """
    One tab instance holding its own grid of companies + month/quarter columns,
    using a distinct section and YEAR in monthly_state.json.
    """
    def __init__(self, master, title: str, storage_key: str, app=None):
        super().__init__(master, padding=8)
        self.app = app
        self.storage_key = storage_key
        self.current_year = datetime.now().year
        self._pending_focus_key = None

        # Data
        self._clients = _load_clients()             # list[(label, key)]
        self._label_by_key = {k: lbl for (lbl, k) in self._clients}
        self._keys, self._cells = _load_tab_year_state(self.storage_key, self.current_year)

        # UI: header row
        top = ttk.Frame(self); top.pack(fill="x")
        ttk.Label(top, text=title, font=("Segoe UI", 12, "bold")).pack(side="left")

        # Year select
        right_box = ttk.Frame(top); right_box.pack(side="right")
        ttk.Label(right_box, text="Year:").pack(side="left", padx=(0,4))
        years = list(range(2022, datetime.now().year + 1))
        self._year_var = tk.IntVar(value=self.current_year)
        self._year_combo = ttk.Combobox(
            right_box, width=8, state="readonly",
            values=[str(y) for y in years], textvariable=self._year_var
        )
        self._year_combo.pack(side="left")
        self._year_combo.bind("<<ComboboxSelected>>", self._on_year_change)

        self._msg = ttk.Label(top, text="", foreground="#4CAF50")
        self._msg.pack(side="right", padx=(8,0))

        # Controls
        bar = ttk.Frame(self); bar.pack(fill="x", pady=(8,6))
        ttk.Label(bar, text="client:").pack(side="left")
        self._combo = ttk.Combobox(
            bar, width=48, state="normal",
            values=[lbl for (lbl, _k) in self._clients]
        )
        self._combo.pack(side="left", padx=(6,6))
        self._combo.bind("<KeyRelease>", self._on_type_filter)

        ttk.Button(bar, text="Add", style="Accent.TButton", command=self._on_add).pack(side="left")
        ttk.Button(bar, text="Remove Selected", command=self._on_remove).pack(side="left", padx=(6,0))
        ttk.Button(bar, text="Copy to Year", command=self._on_copy_to_year).pack(side="left", padx=(6,0))
        ttk.Button(bar, text="Reload", command=self._on_reload).pack(side="right")
        ttk.Button(bar, text="Save", command=self._on_save).pack(side="right", padx=(6,8))

        # Table (client + months/quarters)
        try:
            s = ttk.Style()
            s.configure("Treeview", rowheight=40)
        except Exception:
            pass

        table = ttk.Frame(self); table.pack(fill="both", expand=True)
        
        # LEFT tree: frozen client column
        self._ltree = ttk.Treeview(
            table,
            columns=(client_COL,),
            show="headings",
            height=18,
            selectmode="extended"
        )
        self._ltree.heading(client_COL, text=client_COL)
        self._ltree.column(client_COL, width=240, anchor="w", stretch=False)

        # --- Double-click opens client detail page
        def _on_client_dbl(e=None):
            # Which row did we double-click?
            iid = self._ltree.identify_row(e.y) if e else None
            if not iid:
                return

            # Map row -> client key (populated in _refresh_tree)
            key = getattr(self, "_ltree_iid_to_key", {}).get(iid)
            if not key or not getattr(self, "app", None) or not getattr(self.app, "items", None):
                return

            # Find the index in app.items
            target_idx = None
            for i, c in enumerate(self.app.items):
                ein  = (c.get("ein")  or "")
                name = (c.get("name") or "")
                dba  = (c.get("dba")  or "")
                if key and (key == ein or key == name or key == dba):
                    target_idx = i
                    break

            if target_idx is not None:
                try:
                    self.app.navigate("detail", target_idx, push=True)
                except Exception as ex:
                    print("[checklist] navigate(detail) failed:", ex)  # tiny breadcrumb

        self._ltree.bind("<Double-1>", _on_client_dbl)

        # RIGHT tree: months/quarters (scrollable)
        self._rtree = ttk.Treeview(
            table,
            columns=COLS,
            show="headings",
            height=18,
            selectmode="extended"
        )
        for c in COLS:
            self._rtree.heading(c, text=c)
            self._rtree.column(c, width=120, minwidth=100, anchor="center", stretch=False)
        
        # Scrollbars
        vsb = ttk.Scrollbar(table, orient="vertical")
        hsb = ttk.Scrollbar(table, orient="horizontal")
        
        # Layout (grid)
        self._ltree.grid(row=0, column=0, sticky="ns")
        self._rtree.grid(row=0, column=1, sticky="nsew")
        vsb.grid(row=0, column=2, sticky="ns")
        hsb.grid(row=1, column=1, sticky="ew")
        
        table.grid_columnconfigure(1, weight=1)
        table.grid_rowconfigure(0, weight=1)
        
        # Sync vertical scroll between the two trees
        def _sync_y(*args):
            self._ltree.yview(*args)
            self._rtree.yview(*args)
        
        def _on_rtree_yscroll(*args):
            _sync_y(*args)
            self._schedule_render()
        
        vsb.configure(command=_on_rtree_yscroll)
        self._ltree.configure(yscrollcommand=vsb.set)
        self._rtree.configure(yscrollcommand=lambda *a: (vsb.set(*a), self._schedule_render()))
        
        # Horizontal scroll only for RIGHT tree
        self._rtree.configure(xscrollcommand=hsb.set)
        hsb.configure(command=self._rtree.xview)

    
        self._rtree.bind("<Button-1>", self._on_cell_click)
        
        # Reposition editor when horizontal scroll moves
        def _on_xscroll(*args):
            self._rtree.xview(*args)
            self._reposition_editor()
            self._render_editors()
        
        self._rtree.configure(xscrollcommand=lambda *a: (hsb.set(*a), self._reposition_editor()))
        hsb.configure(command=_on_xscroll)
        
        # Keep vertical scroll in sync (and reposition editor)
        def _on_yscroll(*args):
            self._ltree.yview(*args)
            self._rtree.yview(*args)
            self._reposition_editor()
            self._render_editors()
        
        vsb.configure(command=_on_yscroll)
        self._ltree.configure(yscrollcommand=vsb.set)
        self._rtree.configure(yscrollcommand=lambda *a: (vsb.set(*a), self._reposition_editor()))
        
        # Move editor on size changes
        self._rtree.bind("<Configure>", lambda e: self._reposition_editor())

        # Keep a reference so helpers can set it
        self._vbar = vsb
        
        # Vertical sync: both trees report through our sync callbacks
        self._ltree.configure(yscrollcommand=self._sync_from_l)
        self._rtree.configure(yscrollcommand=self._sync_from_r)
        vsb.configure(command=self._yview)
        
        # Mouse wheel on either tree scrolls both
        self._bind_mousewheel(self._ltree)
        self._bind_mousewheel(self._rtree)

        self._ltree.bind("<<TreeviewSelect>>", self._sync_selection_from_left)
        self._rtree.bind("<<TreeviewSelect>>", self._sync_selection_from_right)
        
        # Optional: Delete key removes the selected row
        self._ltree.bind("<Delete>", lambda e: self._on_remove())
        self._rtree.bind("<Delete>", lambda e: self._on_remove())

        # stripe
        try:
            NewUI.stripe_tree(self._ltree)
            NewUI.stripe_tree(self._rtree)
        except Exception:
            pass
        
        # Always-on editors management (for RIGHT tree cells only)
        self._cell_widgets = {}  # (rid, col_name) -> ttk.Combobox
        
        # Re-render editors on size/scroll/selection changes
        for w in (self._rtree,):
            w.bind("<Configure>", lambda e: self._schedule_render())
            w.bind("<Expose>",    lambda e: self._schedule_render())
            w.bind("<ButtonRelease-1>", lambda e: self._schedule_render())
            w.bind("<MouseWheel>", lambda e: self._schedule_render())
            w.bind("<<TreeviewSelect>>", lambda e: self._schedule_render())

        try:
            self._style = ttk.Style()
            self._style.configure("CellEmpty.TCombobox")
            self._style.configure("CellTodo.TCombobox", fieldbackground="#FFF7CC")  # light yellow
            self._style.configure("CellDone.TCombobox", fieldbackground="#DFF7DF")  # light green
        except Exception:
            self._style = None
        
        self._refresh_tree()
        self._toast("Loaded ✓")

    def set_pending_focus_key(self, key: str):
        """Set a one-shot focus target; applied after next refresh/render."""
        self._pending_focus_key = key
        self.after_idle(self._apply_pending_focus)

    def _apply_pending_focus(self):
        """Scroll/select the pending client in the LEFT tree (if any), then clear it."""
        key = getattr(self, "_pending_focus_key", None)
        if not key or not getattr(self, "_ltree", None) or not self._ltree.winfo_exists():
            return

        wanted_key = (key or "").strip()
        wanted_label = (self._label_by_key.get(wanted_key) or wanted_key).strip().lower()

        # Try exact label
        for iid in self._ltree.get_children(""):
            vals = self._ltree.item(iid, "values") or []
            label = str(vals[0]).strip().lower() if vals else ""
            if label == wanted_label:
                self._ltree.selection_set(iid); self._ltree.focus(iid); self._ltree.see(iid)
                self._pending_focus_key = None     # <-- one-shot: clear after success
                return

        # Fallback: raw key
        low_key = wanted_key.lower()
        for iid in self._ltree.get_children(""):
            vals = self._ltree.item(iid, "values") or []
            label = str(vals[0]).strip().lower() if vals else ""
            if label == low_key:
                self._ltree.selection_set(iid); self._ltree.focus(iid); self._ltree.see(iid)
                self._pending_focus_key = None     # <-- clear after success
                return

        
    def focus_left_by_key(self, key: str):
        """Find row with this logical key in the left tree and bring it into view."""
        if not key or not hasattr(self, "_ltree") or not self._ltree.winfo_exists():
            return
        # The left tree shows the label; keep both exact key & label match attempts
        wanted_key = (key or "").strip()
        wanted_label = (self._label_by_key.get(wanted_key) or wanted_key).strip().lower()

        for iid in self._ltree.get_children(""):
            vals = self._ltree.item(iid, "values") or []
            label = str(vals[0]).strip().lower() if vals else ""
            # Exact label match first
            if label == wanted_label:
                self._ltree.selection_set(iid)
                self._ltree.focus(iid)
                self._ltree.see(iid)
                return

        # Fallback: try raw key string in case label mapping differs
        for iid in self._ltree.get_children(""):
            vals = self._ltree.item(iid, "values") or []
            label = str(vals[0]).strip().lower() if vals else ""
            if label == wanted_key.strip().lower():
                self._ltree.selection_set(iid)
                self._ltree.focus(iid)
                self._ltree.see(iid)
                return

    # --- Year change ---
    def _on_year_change(self, _e=None):
        try:
            y = int(self._year_combo.get())
        except Exception:
            y = self.current_year
        self.current_year = y
        self._keys, self._cells = _load_tab_year_state(self.storage_key, self.current_year)
        self._sort_keys()
        self._refresh_tree()
        self._render_editors()
        self.after_idle(self._apply_pending_focus)
        self._toast(f"Year {y} loaded")

    # --- Type filter (keep caret, no auto-dropdown) ---
    def _on_type_filter(self, event):
        if event.keysym in (
            "Up","Down","Left","Right","Home","End",
            "Prior","Next","Tab","Return","Escape"
        ):
            return

        text = (self._combo.get() or "")
        try:
            insert_idx = self._combo.index(tk.INSERT)
        except Exception:
            insert_idx = None

        all_labels = [lbl for (lbl, _k) in self._clients]
        if text.strip():
            low = text.strip().lower()
            vals = [lbl for lbl in all_labels if low in lbl.lower()]
        else:
            vals = all_labels

        self._combo.configure(values=vals)

        try:
            if insert_idx is not None:
                self._combo.icursor(insert_idx)
            if hasattr(self._combo, "selection_clear"):
                self._combo.selection_clear()
            else:
                self._combo.selection_range(0, 0)
        except Exception:
            pass

    def _on_copy_to_year(self):
        import tkinter.simpledialog as sd
        from datetime import datetime

        try:
            # Ask for target year (bounded 2022..current)
            years = list(range(2022, datetime.now().year + 1))
            tgt = sd.askinteger(
                "Copy to Year",
                f"Copy {len(self._keys)} companies from {self.current_year} to which year?\n"
                f"(Existing companies in the target year will NOT be overwritten.)",
                minvalue=min(years),
                maxvalue=max(years),
                initialvalue=self.current_year
            )
            if not tgt:
                return

            # Load target
            tgt_keys, tgt_cells = _load_tab_year_state(self.storage_key, tgt)
            before = set(tgt_keys)

            # Union without overwriting existing client rows
            for k in self._keys:
                if k not in before:
                    tgt_keys.append(k)
                    tgt_cells.setdefault(k, {c: "" for c in COLS})

            ok, err = _save_tab_year_state(self.storage_key, tgt, tgt_keys, tgt_cells)
            if ok:
                added = len(set(tgt_keys) - before)
                self._toast(f"Copied ✓ {added} new compan{'y' if added == 1 else 'ies'} to {tgt}")
            else:
                messagebox.showerror("Save failed", f"Could not write monthly_state.json\n\n{err}")
        except Exception as e:
            messagebox.showerror("Copy failed", f"{e}")

    # --- Add / Remove / Save / Reload ---
    def _on_add(self):
        label_typed = (self._combo.get() or "").strip()
        if not label_typed:
            messagebox.showinfo("Pick a client", "Choose a client from the dropdown (or type to filter).")
            return

        pair = next(((lbl, k) for (lbl, k) in self._clients if lbl == label_typed), None)
        if not pair:
            matches = [(lbl, k) for (lbl, k) in self._clients if label_typed.lower() in lbl.lower()]
            if len(matches) == 1:
                pair = matches[0]
            else:
                messagebox.showerror("Not found", f"Could not resolve a single client for:\n{label_typed}")
                return

        label, key = pair
        if key in self._keys:
            self._toast("Already added")
            return

        self._keys.append(key)
        self._sort_keys()
        self._cells.setdefault(key, {c: "--" for c in COLS})
        self._refresh_tree()
        self._render_editors()

        ok, err = _save_tab_year_state(self.storage_key, self.current_year, self._keys, self._cells)
        if ok:
            self._toast(f"Saved ✓ Added: {label}")
        else:
            messagebox.showerror("Save failed", f"Could not write monthly_state.json\n\n{err}")

    def _on_remove(self):
        """Remove the selected client from current year, persist, and keep selection sane."""
        idx, key = self._get_current_selection_key()
        if key is None:
            # use your toast label if you have one
            try: self._toast("Select a row first")
            except Exception: pass
            return
    
        # Remove from model
        try:
            if key in self._cells:
                self._cells.pop(key, None)
            if key in self._keys:
                self._keys.remove(key)
        except Exception:
            pass
    
        # Persist
        try:
            _save_tab_year_state(self.storage_key, self.current_year, self._keys, self._cells)
        except Exception:
            pass
    
        # Refresh UI
        self._refresh_tree()
    
        # Restore a sensible selection (same index, or previous row)
        try:
            new_idx = max(0, min(idx, len(self._keys) - 1))
            if self._keys:
                r_rows = self._rtree.get_children("")
                l_rows = self._ltree.get_children("")
                if 0 <= new_idx < len(r_rows):
                    self._rtree.selection_set(r_rows[new_idx])
                    self._rtree.see(r_rows[new_idx])
                if 0 <= new_idx < len(l_rows):
                    self._ltree.selection_set(l_rows[new_idx])
                    self._ltree.see(l_rows[new_idx])
        except Exception:
            pass
    
        try: self._toast("Removed")
        except Exception: pass


    def _on_reload(self):
        self._keys, self._cells = _load_tab_year_state(self.storage_key, self.current_year)
        self._sort_keys()
        self._refresh_tree()
        self._render_editors()
        self._toast("Loaded ✓")

    def _on_save(self):
        ok, err = _save_tab_year_state(self.storage_key, self.current_year, self._keys, self._cells)
        if ok:
            self._toast("Saved ✓")
        else:
            messagebox.showerror("Save failed", f"Could not write monthly_state.json\n\n{err}")

    # --- Tree refresh & always-on editors ---
    def _refresh_tree(self):
        # clear
        for t in (self._ltree, self._rtree):
            for iid in t.get_children(""):
                t.delete(iid)
    
        self._rid_to_key = {}
        self._ltree_iid_to_key = {}
        for key in self._keys:
            label = self._label_by_key.get(key, key)
            iid_left = self._ltree.insert("", "end", values=[label])
            self._ltree_iid_to_key[iid_left] = key 
    
            vals = [_store_to_ui(self._cells.get(key, {}).get(c, "")) for c in COLS]
            rid = self._rtree.insert("", "end", values=vals)
            self._rid_to_key[rid] = key
    
        # Force back to the first column (January)
        try:
            self._rtree.xview_moveto(0.0)
        except Exception:
            pass
    
        # Reset to top
        try:
            self._ltree.yview_moveto(0.0)
            self._rtree.yview_moveto(0.0)
        except Exception:
            pass
    
        self._schedule_render()
        self.after_idle(self._apply_pending_focus)

    def _schedule_render(self):
        # Debounce rapid events
        try:
            if getattr(self, "_render_after_id", None):
                self.after_cancel(self._render_after_id)
        except Exception:
            pass
        self._render_after_id = self.after(15, self._render_editors)


    def _render_editors(self):
        """Place combobox editors to match visible RIGHT-tree cells."""
        visible_keys = set()  # track which (rid, col) are visible this pass
    
        for rid in self._rtree.get_children(""):
            key = self._rid_to_key.get(rid)
            if not key:
                continue
    
            for col_name in COLS:  # use column identifiers, not numeric indexes
                bbox = self._rtree.bbox(rid, column=col_name)
                if not bbox:
                    # Cell isn't visible (scrolled out); we'll hide it later
                    continue
    
                x, y, w, h = bbox
                w_key = (rid, col_name)
                visible_keys.add(w_key)
    
                # Current UI value
                cur_ui = _store_to_ui(self._cells.get(key, {}).get(col_name, ""))
    
                # Choose highlight style
                if cur_ui == "To do":
                    style_name = "CellTodo.TCombobox"
                elif cur_ui == "Done":
                    style_name = "CellDone.TCombobox"
                else:
                    style_name = "CellEmpty.TCombobox"
    
                cb = self._cell_widgets.get(w_key)
                if cb is None:
                    cb = ttk.Combobox(
                        self._rtree,
                        state="readonly",
                        values=CELL_OPTIONS,       # ["--", "To do", "Done"]
                        style=style_name,
                        justify="center"
                    )
                    cb.set(cur_ui if cur_ui in CELL_OPTIONS else "--")
                    cb.place(x=x, y=y, width=w, height=h)
    
                    def _mk_commit(_rid=rid, _key=key, _col=col_name, _cb=cb):
                        def _commit(*_a):
                            ui_val = (_cb.get() or "").strip()
    
                            # Save immediately (updates RAM + JSON + tree cell)
                            normalized_ui = self._commit_cell(_key, _col, ui_val, iid=_rid)
    
                            # Update style based on new value
                            try:
                                if normalized_ui == "To do":
                                    new_style = "CellTodo.TCombobox"
                                elif normalized_ui == "Done":
                                    new_style = "CellDone.TCombobox"
                                else:
                                    new_style = "CellEmpty.TCombobox"
                                _cb.configure(style=new_style)
                            except Exception:
                                pass
                        return _commit
    
                    cb.bind("<<ComboboxSelected>>", _mk_commit())
                    self._cell_widgets[w_key] = cb
                else:
                    # Reposition + sync value + style every pass
                    cb.place(x=x, y=y, width=w, height=h)
                    want = cur_ui if cur_ui in CELL_OPTIONS else "--"
                    if cb.get() != want:
                        cb.set(want)
                    try:
                        if want == "To do":
                            cb.configure(style="CellTodo.TCombobox")
                        elif want == "Done":
                            cb.configure(style="CellDone.TCombobox")
                        else:
                            cb.configure(style="CellEmpty.TCombobox")
                    except Exception:
                        pass
    
        # Hide editors that are currently scrolled out of view
        for w_key, cb in list(self._cell_widgets.items()):
            if w_key not in visible_keys:
                try:
                    cb.place_forget()
                except Exception:
                    pass

    def _commit_cell(self, key: str, col_name: str, ui_value: str, iid=None):
        store_val = _ui_to_store(ui_value)
        ui_val = _store_to_ui(store_val)
    
        # RAM
        self._cells.setdefault(key, {c: "--" for c in COLS})
        self._cells[key][col_name] = ui_val
    
        # UI reflect
        if iid is not None:
            try:
                self._rtree.set(iid, col_name, ui_val)
            except Exception:
                pass
    
        # Persist
        ok, err = _save_tab_year_state(self.storage_key, self.current_year, self._keys, self._cells)
        if not ok:
            messagebox.showerror("Save failed", f"Could not write monthly_state.json\n\n{err}")
        return ui_val

    def _bind_mousewheel(self, widget):
        """Mouse wheel scrolls both trees together (Windows delta=120)."""
        def _on_mousewheel(e):
            steps = -int(e.delta / 120) if e.delta else 0
            if steps:
                self._yview("scroll", steps, "units")
                self._schedule_render() 
                return "break"
        widget.bind("<MouseWheel>", _on_mousewheel)
    
    def _yview(self, *args):
        """Scrollbar/keys -> move both trees."""
        try:
            self._ltree.yview(*args)
            self._rtree.yview(*args)
        except Exception:
            pass
        finally:
            self._schedule_render() 
    
    def _sync_from_l(self, first, last):
        """Left tree moved -> keep right + scrollbar in sync (no loops)."""
        if getattr(self, "_syncing", False):
            # Only update the visible scrollbar bounds
            try: self._vbar.set(first, last)
            except Exception: pass
            return
        self._syncing = True
        try:
            self._rtree.yview_moveto(first)
            self._vbar.set(first, last)
        finally:
            self._syncing = False
        self._schedule_render() 
    
    def _sync_from_r(self, first, last):
        """Right tree moved -> keep left + scrollbar in sync (no loops)."""
        if getattr(self, "_syncing", False):
            try: self._vbar.set(first, last)
            except Exception: pass
            return
        self._syncing = True
        try:
            self._ltree.yview_moveto(first)
            self._vbar.set(first, last)
        finally:
            self._syncing = False
        self._schedule_render() 

    def _normalize_label_for_sort(self, label: str):
        # casefold for case-insensitive, strip accents, keep digits but de-punctuate
        text = unicodedata.normalize("NFKD", label).casefold()
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = re.sub(r"[^0-9a-z]+", " ", text).strip()
        # Put number-leading names after letter-leading names
        return (text[:1].isdigit(), text)
    
    def _sort_keys(self):
        # Sort keys by the human-readable label, fallback to key if missing
        self._keys.sort(key=lambda k: self._normalize_label_for_sort(self._label_by_key.get(k, k)))

    def _selected_index_from(self, tree):
        """Return the row index (0-based) selected in a Treeview, or None."""
        sel = tree.selection()
        if not sel:
            return None
        # Get all top-level rows to compute index
        rows = tree.get_children("")
        try:
            return rows.index(sel[0])
        except Exception:
            return None
    
    def _get_current_selection_key(self):
        """Figure out the selected key from either right or left tree."""
        # Prefer right-tree selection (it maps directly via rid)
        sel = self._rtree.selection()
        if sel:
            rid = sel[0]
            key = self._rid_to_key.get(rid)
            if key is not None:
                idx = self._selected_index_from(self._rtree)
                return idx, key
        # Fall back to left-tree selection -> map by index into self._keys
        idx = self._selected_index_from(self._ltree)
        if idx is None:
            return None, None
        if 0 <= idx < len(self._keys):
            return idx, self._keys[idx]
        return None, None
    
    def _sync_selection_from_left(self, _evt=None):
        """Mirror selection from left to right."""
        idx = self._selected_index_from(self._ltree)
        if idx is None:
            self._rtree.selection_remove(self._rtree.selection())
            return
        rows = self._rtree.get_children("")
        if 0 <= idx < len(rows):
            self._rtree.selection_set(rows[idx])
            self._rtree.see(rows[idx])
    
    def _sync_selection_from_right(self, _evt=None):
        """Mirror selection from right to left."""
        idx = self._selected_index_from(self._rtree)
        if idx is None:
            self._ltree.selection_remove(self._ltree.selection())
            return
        rows = self._ltree.get_children("")
        if 0 <= idx < len(rows):
            self._ltree.selection_set(rows[idx])
            self._ltree.see(rows[idx])

    # --- Toast ---
    def _toast(self, text):
        try:
            self._msg.configure(text=text)
            self._msg.after(1200, lambda: self._msg.configure(text=""))
        except Exception:
            pass

    def _schedule_render(self, delay_ms: int = 10):
        """Throttle re-placement of editors so scrolling stays smooth."""
        try:
            if getattr(self, "_render_after", None):
                self.after_cancel(self._render_after)
        except Exception:
            pass
        self._render_after = self.after(delay_ms, self._render_editors)

    def _on_cell_click(self, event):
        rid = self._rtree.identify_row(event.y)
        if not rid:
            self._close_editor()
            return
        col_id = self._rtree.identify_column(event.x)  # "#1", "#2", ...
        try:
            idx = int(col_id.replace("#", "")) - 1
        except Exception:
            idx = -1
        if 0 <= idx < len(COLS):
            self._open_editor_at(rid, COLS[idx])
        else:
            self._close_editor()

    # ---- Editor helpers (now belong here) ----
    def _open_editor_at(self, rid, col_name):
        """Open or move the single editor to this cell, prefilled with current value."""
        key = self._rid_to_key.get(rid)
        if not key:
            return

        bbox = self._rtree.bbox(rid, column=col_name)
        if not bbox:
            self._close_editor()
            return

        x, y, w, h = bbox
        cur_ui = _store_to_ui(self._cells.get(key, {}).get(col_name, ""))

        # lazy-create editor
        if getattr(self, "_editor", None) is None:
            self._editor = ttk.Combobox(
                self._rtree,
                state="readonly",
                values=CELL_OPTIONS,
                justify="center"
            )
            # commit on selection
            def _commit(event=None, _rid=rid, _key=key, _col=col_name):
                ui_val = (self._editor.get() or "").strip()
                normalized_ui = self._commit_cell(_key, _col, ui_val, iid=_rid)
                try:
                    self._rtree.set(_rid, _col, normalized_ui)
                except Exception:
                    pass
                self._editor.set(normalized_ui)
            self._editor.bind("<<ComboboxSelected>>", _commit)

        # remember where we are
        self._editor_rid = rid
        self._editor_col = col_name
        self._editor_key = key

        # style by value
        try:
            if cur_ui == "To do":
                self._editor.configure(style="CellTodo.TCombobox")
            elif cur_ui == "Done":
                self._editor.configure(style="CellDone.TCombobox")
            else:
                self._editor.configure(style="CellEmpty.TCombobox")
        except Exception:
            pass

        self._editor.set(cur_ui if cur_ui in CELL_OPTIONS else "--")
        self._editor.place(x=x, y=y, width=w, height=h)

    def _reposition_editor(self):
        """If the editor is open, keep it glued to its cell after scroll/resize."""
        if getattr(self, "_editor", None) is None:
            return
        rid, col = getattr(self, "_editor_rid", None), getattr(self, "_editor_col", None)
        if not rid or not col or rid not in self._rtree.get_children(""):
            self._close_editor()
            return
        bbox = self._rtree.bbox(rid, column=col)
        if not bbox:
            self._editor.place_forget()
            return
        x, y, w, h = bbox
        self._editor.place(x=x, y=y, width=w, height=h)

    def _close_editor(self):
        if getattr(self, "_editor", None) is not None:
            try:
                self._editor.place_forget()
            except Exception:
                pass
        self._editor_rid = None
        self._editor_col = None
        self._editor_key = None

class ChecklistPage:
    """
    Page object compatible with App.navigate(...):
        ChecklistPage(app=self).ensure(self.page_host)
    """
    def __init__(self, app=None, **_kwargs):
        self.app = app
        self.frame = None
        self._pending_focus_key = None
        self.log = get_logger("checklist")

    def ensure(self, host):
        self.log.info("ensure() checklist host=%s", getattr(host, "winfo_name", lambda: "?")())
        if self.frame is None:
            self.frame = ttk.Frame(host, padding=8)
            self._build(self.frame)
        try:
            self.frame.pack_forget()
        except Exception as e:        
            LOG.warning("Editor reposition failed: %s", e)
        self.frame.pack(fill="both", expand=True)
        return self.frame

    def _build(self, root):
        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True)

        tabs = [
            ("Monthly Tax", KEY_MONTHLY_TAX),
            ("Sales Tax", KEY_SALES_TAX),
            ("Financial Statement", KEY_FINSTAT),
        ]
        for title, key in tabs:
            tab = _ChecklistTab(nb, title=title, storage_key=key, app=self.app)
            nb.add(tab, text=title)

    def focus_client_by_key(self, key: str):
        """One-shot: when returning from Detail, ask the ACTIVE tab to focus this client."""
        self.log.info("focus_client_by_key(%s)", key)
        if not key:
            return
        try:
            current = self._notebook.select()
            tab = self._notebook.nametowidget(current)
            if hasattr(tab, "set_pending_focus_key"):
                tab.set_pending_focus_key(key)   # one-shot for the current tab only
        except Exception:
            pass

    def scroll_top(self):
        """Scroll to the top of the active tab."""
        try:
            current = self._notebook.select()
            tab = self._notebook.nametowidget(current)
            if hasattr(tab, "_ltree") and tab._ltree.winfo_exists():
                tab._ltree.yview_moveto(0.0)
            if hasattr(tab, "_rtree") and tab._rtree.winfo_exists():
                tab._rtree.yview_moveto(0.0)
        except Exception:
            pass


# Standalone demo
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Checklist")
    try:
        NewUI.install(root)
    except Exception:
        pass
    page = ChecklistPage()
    page.ensure(root)
    root.minsize(1200, 640)
    root.mainloop()
