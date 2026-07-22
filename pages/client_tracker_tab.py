from __future__ import annotations

import sys
import pathlib
import datetime

if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

try:
    from vertex.utils.client_tracker import (
        ensure_tracker_fields,
        count_tracker_summary,
        collect_tax_years,
        add_file_request,
        update_file_request,
        add_annual_reminder,
        update_annual_reminder,
        add_client_issue,
        update_client_issue,
        append_tracker_log,
        roll_forward_file_requests,
        file_status_tag,
        reminder_status_tag,
        issue_status_tag,
        default_tax_year,
        FILE_REQUEST_STATUSES,
        REMINDER_STATUSES,
        ISSUE_STATUSES,
        CATEGORIES,
        today_str,
        now_ts,
    )
    from vertex.ui.dialogs.tracker_item_dialog import TrackerItemDialog
    from vertex.ui.components.scrollframe import ScrollFrame
except ModuleNotFoundError:
    from utils.client_tracker import (
        ensure_tracker_fields,
        count_tracker_summary,
        collect_tax_years,
        add_file_request,
        update_file_request,
        add_annual_reminder,
        update_annual_reminder,
        add_client_issue,
        update_client_issue,
        append_tracker_log,
        roll_forward_file_requests,
        file_status_tag,
        reminder_status_tag,
        issue_status_tag,
        default_tax_year,
        FILE_REQUEST_STATUSES,
        REMINDER_STATUSES,
        ISSUE_STATUSES,
        CATEGORIES,
        today_str,
        now_ts,
    )
    from ui.dialogs.tracker_item_dialog import TrackerItemDialog
    from ui.components.scrollframe import ScrollFrame

TRACKER_TAB_LABEL = "Tracker"
_ROW_SEL_BG = "#EEF2FF"
_ROW_BG = "#FFFFFF"
_HDR_BG = "#F3F4F6"
_ROW_ALT_BG = "#FAFAFA"

# Shared column layout — fixed widths keep header/body aligned; cols 4 & 9 stretch.
_COL_WIDTHS = (288, 88, 52, 108, 120, 92, 92, 56, 72, 100)
_COL_WEIGHTS = (0, 0, 0, 0, 1, 0, 0, 0, 0, 2)
_COL_FIXED = (True, True, True, True, False, True, True, True, True, False)

_FILE_HEADERS = (
    "Quick set", "Status", "Year", "Category", "File / Item",
    "Requested", "Received", "Repeat", "Priority", "Note",
)
_REM_HEADERS = (
    "Quick set", "Status", "Year", "Category", "Title",
    "Every yr", "—", "—", "Priority", "Note",
)
_ISSUE_HEADERS = (
    "Quick set", "Status", "Year", "Category", "Title",
    "Type", "Opened", "Closed", "Priority", "Note",
)

_PRIORITY_FG = {
    "Urgent": "#DC2626",
    "High": "#D97706",
    "Normal": "#374151",
    "Low": "#9CA3AF",
}

_STATUS_STYLE = {
    "tr_needed": ("#FEF3C7", "#92400E"),
    "tr_requested": ("#DBEAFE", "#1E40AF"),
    "tr_received": ("#D1FAE5", "#065F46"),
    "tr_waiting": ("#FFEDD5", "#9A3412"),
    "tr_open": ("#FEE2E2", "#991B1B"),
    "tr_active": ("#EEF2FF", "#3730A3"),
    "tr_closed": ("#F3F4F6", "#6B7280"),
    "tr_archived": ("#E5E7EB", "#9CA3AF"),
    "tr_default": (_ROW_BG, "#374151"),
}


def _configure_pill_styles(style: ttk.Style) -> None:
    style.configure(
        "Tracker.Pill.TButton",
        padding=(5, 2),
        background="#FFFFFF",
        foreground="#111827",
        borderwidth=1,
        relief="solid",
        bordercolor="#D1D5DB",
        lightcolor="#D1D5DB",
        darkcolor="#D1D5DB",
        font=("Segoe UI", 8),
    )
    style.map(
        "Tracker.Pill.TButton",
        background=[("active", "#F9FAFB"), ("pressed", "#F3F4F6")],
        bordercolor=[("active", "#9CA3AF"), ("pressed", "#6B7280")],
    )
    style.configure(
        "Tracker.Pill.Active.TButton",
        padding=(5, 2),
        background="#F8FAFC",
        foreground="#111827",
        borderwidth=2,
        relief="solid",
        bordercolor="#6366F1",
        lightcolor="#6366F1",
        darkcolor="#6366F1",
        font=("Segoe UI", 8),
    )
    style.map(
        "Tracker.Pill.Active.TButton",
        background=[("active", "#EEF2FF"), ("pressed", "#E0E7FF")],
    )


def _pill_style(active: bool) -> str:
    return "Tracker.Pill.Active.TButton" if active else "Tracker.Pill.TButton"


def _priority_fg(priority: str) -> str:
    return _PRIORITY_FG.get(str(priority or "Normal").strip(), "#374151")


def _status_style_for(kind: str, status: str) -> tuple[str, str]:
    st = str(status or "").strip()
    if kind == "file":
        tag = file_status_tag(st)
    elif kind == "reminder":
        tag = reminder_status_tag(st)
    else:
        tag = issue_status_tag(st)
    return _STATUS_STYLE.get(tag, _STATUS_STYLE["tr_default"])


def _make_row_columns(row: tk.Frame, bg: str) -> list[tk.Frame]:
    """Create one fixed-layout cell frame per column (header + data rows use the same helper)."""
    cells: list[tk.Frame] = []
    for col, (width, weight, fixed) in enumerate(zip(_COL_WIDTHS, _COL_WEIGHTS, _COL_FIXED)):
        if fixed:
            cell = tk.Frame(row, bg=bg, width=width, height=1)
            row.grid_columnconfigure(col, minsize=width, weight=0)
            cell.grid(row=0, column=col, sticky="nsew")
            cell.grid_propagate(False)
        else:
            cell = tk.Frame(row, bg=bg)
            row.grid_columnconfigure(col, minsize=width, weight=weight)
            cell.grid(row=0, column=col, sticky="nsew")
        cells.append(cell)
    return cells


def _make_cell_label(
    parent: tk.Frame,
    text: str,
    bg: str,
    *,
    bold: bool = False,
    fg: str = "#374151",
    anchor: str = "w",
) -> tk.Label:
    font = ("Segoe UI", 9, "bold") if bold else ("Segoe UI", 9)
    lbl = tk.Label(
        parent, text=text or "", anchor=anchor, bg=bg, fg=fg,
        font=font, padx=4, pady=6,
    )
    lbl.pack(fill="both", expand=True, anchor=anchor)
    return lbl


_SCROLLBAR_GUTTER = 17


def _build_table_panel(parent, headers: tuple[str, ...], height: int = 340) -> ScrollFrame:
    """Header + scroll body share fixed column cells; gutter matches vertical scrollbar."""
    panel = ttk.Frame(parent)
    panel.pack(fill="both", expand=True)

    top = tk.Frame(panel)
    top.pack(fill="x")
    hdr = tk.Frame(top, bg=_HDR_BG)
    hdr.pack(side="left", fill="x", expand=True)
    hdr_row = tk.Frame(hdr, bg=_HDR_BG)
    hdr_row.pack(fill="x")
    hdr_cells = _make_row_columns(hdr_row, _HDR_BG)
    for i, text in enumerate(headers):
        _make_cell_label(hdr_cells[i], text, _HDR_BG, bold=True, fg="#111827")
    tk.Frame(top, width=_SCROLLBAR_GUTTER, bg=_HDR_BG).pack(side="right", fill="y")

    ttk.Separator(panel, orient="horizontal").pack(fill="x")

    scroll = ScrollFrame(panel, height=height)
    scroll.pack(fill="both", expand=True)
    return scroll


def init_client_tracker_tab(
    notebook: ttk.Notebook,
    app,
    client: dict,
    save_clients_cb,
    client_idx: int | None = None,
):
    ensure_tracker_fields(client)

    tracker_tab = ttk.Frame(notebook, padding=8)
    notebook.add(tracker_tab, text=TRACKER_TAB_LABEL)

    summary_var = tk.StringVar(value="")
    ttk.Label(tracker_tab, textvariable=summary_var, font=("Segoe UI", 10, "bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 6)
    )

    filter_frm = ttk.Frame(tracker_tab)
    filter_frm.grid(row=1, column=0, sticky="ew", pady=(0, 6))
    filter_frm.columnconfigure(1, weight=1)

    v_search = tk.StringVar()
    v_year = tk.StringVar(value="All")
    v_status = tk.StringVar(value="All")
    v_category = tk.StringVar(value="All")

    ttk.Label(filter_frm, text="Search").grid(row=0, column=0, sticky="w", padx=(0, 4))
    ttk.Entry(filter_frm, textvariable=v_search, width=18).grid(row=0, column=1, sticky="ew", padx=(0, 8))
    ttk.Label(filter_frm, text="Year").grid(row=0, column=2, sticky="w", padx=(0, 4))
    cb_year = ttk.Combobox(filter_frm, textvariable=v_year, width=8, state="readonly")
    cb_year.grid(row=0, column=3, sticky="w", padx=(0, 8))
    ttk.Label(filter_frm, text="Status").grid(row=0, column=4, sticky="w", padx=(0, 4))
    cb_status = ttk.Combobox(filter_frm, textvariable=v_status, width=14, state="readonly")
    cb_status.grid(row=0, column=5, sticky="w", padx=(0, 8))
    ttk.Label(filter_frm, text="Category").grid(row=0, column=6, sticky="w", padx=(0, 4))
    cb_category = ttk.Combobox(filter_frm, textvariable=v_category, width=14, state="readonly")
    cb_category.grid(row=0, column=7, sticky="w")

    action_bar = ttk.Frame(tracker_tab)
    action_bar.grid(row=2, column=0, sticky="w", pady=(0, 6))

    inner_nb = ttk.Notebook(tracker_tab)
    inner_nb.grid(row=3, column=0, sticky="nsew")
    tracker_tab.grid_rowconfigure(3, weight=1)
    tracker_tab.grid_columnconfigure(0, weight=1)

    _style = ttk.Style(tracker_tab)
    _configure_pill_styles(_style)

    files_tab = ttk.Frame(inner_nb, padding=4)
    reminders_tab = ttk.Frame(inner_nb, padding=4)
    issues_tab = ttk.Frame(inner_nb, padding=4)
    inner_nb.add(reminders_tab, text="Annual Reminders")
    inner_nb.add(files_tab, text="File Requests")
    inner_nb.add(issues_tab, text="Issues / Notices")

    files_scroll = _build_table_panel(files_tab, _FILE_HEADERS)
    files_body = files_scroll.inner

    reminders_scroll = _build_table_panel(reminders_tab, _REM_HEADERS)
    reminders_body = reminders_scroll.inner

    issues_scroll = _build_table_panel(issues_tab, _ISSUE_HEADERS)
    issues_body = issues_scroll.inner

    files_selected_id: dict[str, str | None] = {"id": None}
    reminders_selected_id: dict[str, str | None] = {"id": None}
    issues_selected_id: dict[str, str | None] = {"id": None}
    files_row_refs: dict[str, dict] = {}
    reminders_row_refs: dict[str, dict] = {}
    issues_row_refs: dict[str, dict] = {}

    # ---- shared helpers (closures) ----

    def _refresh_logs_tab():
        try:
            refresher_map = getattr(app, "_logs_tab_refreshers", {}) or {}
            fn = refresher_map.get(id(client))
            if callable(fn):
                fn()
        except Exception:
            pass

    def _refresh_external_summaries():
        try:
            prof = getattr(app, "_detail_profile_frame", None)
            if prof and hasattr(prof, "_refresh_tracker_summary"):
                prof._refresh_tracker_summary()
        except Exception:
            pass
        try:
            dash = getattr(app, "dashboard", None)
            if dash and hasattr(dash, "_refresh_tracker_sidebar"):
                dash._refresh_tracker_sidebar()
        except Exception:
            pass

    def _persist():
        save_clients_cb(app.items)
        _refresh_summary()
        _refresh_current_tree()
        _refresh_logs_tab()
        _refresh_external_summaries()

    def _refresh_summary():
        s = count_tracker_summary(client)
        summary_var.set(
            f"Missing: {s['missing']} | Requested: {s['requested']} | Received: {s['received']} | "
            f"Reminders: {s['active_reminders']} | Open Issues: {s['open_issues']} | IPP Open: {s['ipp_open']}"
        )

    def _current_kind() -> str:
        sel = inner_nb.select()
        if sel == str(reminders_tab):
            return "reminder"
        if sel == str(issues_tab):
            return "issue"
        return "file"

    def _refresh_filter_combos():
        years = ["All"] + collect_tax_years(client)
        cb_year["values"] = years
        if v_year.get() not in years:
            v_year.set("All")
        kind = _current_kind()
        if kind == "file":
            statuses = ["All"] + list(FILE_REQUEST_STATUSES)
        elif kind == "reminder":
            statuses = ["All"] + list(REMINDER_STATUSES)
        else:
            statuses = ["All"] + list(ISSUE_STATUSES)
        cb_status["values"] = statuses
        if v_status.get() not in statuses:
            v_status.set("All")
        cats = ["All"] + list(CATEGORIES)
        cb_category["values"] = cats
        if v_category.get() not in cats:
            v_category.set("All")

    def _matches_filter(item: dict) -> bool:
        if not isinstance(item, dict):
            return False
        year_f = v_year.get()
        if year_f and year_f != "All":
            if str(item.get("tax_year", "") or "").strip() != year_f:
                return False
        status_f = v_status.get()
        if status_f and status_f != "All":
            if str(item.get("status", "") or "") != status_f:
                return False
        cat_f = v_category.get()
        if cat_f and cat_f != "All":
            if str(item.get("category", "") or "") != cat_f:
                return False
        q = v_search.get().strip().casefold()
        if q:
            hay = " ".join(str(item.get(k, "") or "") for k in item).casefold()
            if q not in hay:
                return False
        return True

    def _find_by_id(collection: str, item_id: str) -> dict | None:
        for it in client.get(collection) or []:
            if isinstance(it, dict) and str(it.get("id", "") or "") == item_id:
                return it
        return None

    def _selected_id(kind: str) -> str | None:
        if kind == "file":
            return files_selected_id.get("id")
        if kind == "reminder":
            return reminders_selected_id.get("id")
        return issues_selected_id.get("id")

    def _row_refs(kind: str) -> dict:
        if kind == "file":
            return files_row_refs
        if kind == "reminder":
            return reminders_row_refs
        return issues_row_refs

    def _selected_store(kind: str) -> dict:
        if kind == "file":
            return files_selected_id
        if kind == "reminder":
            return reminders_selected_id
        return issues_selected_id

    def _update_row_highlight(kind: str):
        sel = _selected_store(kind).get("id")
        for item_id, ref in _row_refs(kind).items():
            bg = _ROW_SEL_BG if item_id == sel else ref.get("base_bg", _ROW_BG)
            ref["frame"].configure(bg=bg)
            btn_parent = ref.get("btn_frm")
            if btn_parent:
                btn_parent.configure(bg=bg)
            for cell in ref.get("col_frames", []):
                if cell is ref.get("status_cell"):
                    continue
                try:
                    cell.configure(bg=bg)
                except tk.TclError:
                    pass
            for lbl in ref.get("labels", []):
                if lbl is ref.get("status_lbl"):
                    continue
                try:
                    lbl.configure(bg=bg)
                except tk.TclError:
                    pass
            st_bg, st_fg = _status_style_for(kind, ref.get("status", ""))
            if ref.get("status_cell"):
                ref["status_cell"].configure(bg=st_bg)
            if ref.get("status_lbl"):
                ref["status_lbl"].configure(bg=st_bg, fg=st_fg)

    def _toggle_selection(kind: str, item_id: str):
        store = _selected_store(kind)
        if store.get("id") == item_id:
            store["id"] = None
        else:
            store["id"] = item_id
        _update_row_highlight(kind)

    def _clear_selection(kind: str):
        _selected_store(kind)["id"] = None
        _update_row_highlight(kind)

    def _bind_row_select(row: tk.Frame, labels: list, kind: str, item_id: str, on_dbl):
        def _click(_e, iid=item_id):
            _toggle_selection(kind, iid)

        def _dbl(_e, iid=item_id):
            _selected_store(kind)["id"] = iid
            _update_row_highlight(kind)
            on_dbl(iid)

        for w in [row, *labels]:
            w.bind("<Button-1>", _click)
            w.bind("<Double-1>", _dbl)

    def _fill_row_cells(ref: dict, kind: str, it: dict):
        st = str(it.get("status", "") or "").strip()
        ref["status"] = st
        st_bg, st_fg = _status_style_for(kind, st)
        ref["status_lbl"].configure(text=st, bg=st_bg, fg=st_fg)
        if ref.get("status_cell"):
            ref["status_cell"].configure(bg=st_bg)

        pri = str(it.get("priority", "") or "Normal")
        ref["priority_lbl"].configure(text=pri, fg=_priority_fg(pri))

        ref["note_lbl"].configure(text=str(it.get("note", "") or ""))

        if kind == "file":
            ref["cells"]["year"].configure(text=str(it.get("tax_year", "") or ""))
            ref["cells"]["category"].configure(text=str(it.get("category", "") or ""))
            ref["cells"]["title"].configure(text=str(it.get("name", "") or ""))
            ref["cells"]["c5"].configure(text=str(it.get("requested_date", "") or ""))
            ref["cells"]["c6"].configure(text=str(it.get("received_date", "") or ""))
            ref["cells"]["c7"].configure(text="Yes" if it.get("repeat_next_year") else "No")
            for status, btn in ref["buttons"].items():
                btn.configure(style=_pill_style(st == status))

        elif kind == "reminder":
            ref["cells"]["year"].configure(text=str(it.get("tax_year", "") or ""))
            ref["cells"]["category"].configure(text=str(it.get("category", "") or ""))
            ref["cells"]["title"].configure(text=str(it.get("title", "") or ""))
            ref["cells"]["c5"].configure(text="Yes" if it.get("applies_every_year") else "No")
            ref["cells"]["c6"].configure(text="")
            ref["cells"]["c7"].configure(text="")
            ref["buttons"]["Active"].configure(style=_pill_style(st == "Active"))
            ref["buttons"]["Done"].configure(style=_pill_style(st == "Done This Year"))

        else:
            ref["cells"]["year"].configure(text=str(it.get("tax_year", "") or ""))
            ref["cells"]["category"].configure(text=str(it.get("category", "") or ""))
            ref["cells"]["title"].configure(text=str(it.get("title", "") or ""))
            ref["cells"]["c5"].configure(text=str(it.get("type", "") or ""))
            ref["cells"]["c6"].configure(text=str(it.get("opened_date", "") or ""))
            ref["cells"]["c7"].configure(text=str(it.get("closed_date", "") or ""))
            open_active = st in ("Open", "Waiting on Client", "Waiting on Government")
            closed_active = st in ("Closed", "Resolved", "Archived")
            ref["buttons"]["Open"].configure(style=_pill_style(open_active))
            ref["buttons"]["Closed"].configure(style=_pill_style(closed_active))

    def _build_data_row(parent, kind: str, it: dict, row_index: int) -> dict | None:
        item_id = str(it.get("id", "") or "").strip()
        if not item_id:
            return None

        base_bg = _ROW_BG if row_index % 2 == 0 else _ROW_ALT_BG
        row = tk.Frame(parent, bg=base_bg, padx=0, pady=0)
        row.pack(fill="x")
        col_frames = _make_row_columns(row, base_bg)

        btn_frm = tk.Frame(col_frames[0], bg=base_bg)
        btn_frm.pack(fill="x", padx=2, pady=4)
        buttons: dict[str, ttk.Button] = {}

        if kind == "file":
            st = str(it.get("status", "") or "").strip()
            for label, status in (("Needed", "Needed"), ("Requested", "Requested"), ("Received", "Received")):
                b = ttk.Button(
                    btn_frm, text=label, style=_pill_style(st == status), width=len(label),
                    command=lambda s=status, iid=item_id: _file_status_click(iid, s),
                )
                b.pack(side=tk.LEFT, padx=(0, 2))
                buttons[status] = b
        elif kind == "reminder":
            st = str(it.get("status", "") or "").strip()
            for label, key in (("Active", "Active"), ("Done", "Done")):
                active = (key == "Active" and st == "Active") or (key == "Done" and st == "Done This Year")
                b = ttk.Button(
                    btn_frm, text=label, style=_pill_style(active), width=len(label),
                    command=lambda k=key, iid=item_id: _reminder_quick_click(iid, k),
                )
                b.pack(side=tk.LEFT, padx=(0, 2))
                buttons[key] = b
        else:
            st = str(it.get("status", "") or "").strip()
            open_active = st in ("Open", "Waiting on Client", "Waiting on Government")
            closed_active = st in ("Closed", "Resolved", "Archived")
            b_open = ttk.Button(
                btn_frm, text="Open", style=_pill_style(open_active), width=5,
                command=lambda iid=item_id: _issue_status_click(iid, "Open"),
            )
            b_close = ttk.Button(
                btn_frm, text="Close", style=_pill_style(closed_active), width=5,
                command=lambda iid=item_id: _issue_status_click(iid, "Closed"),
            )
            b_open.pack(side=tk.LEFT, padx=(0, 2))
            b_close.pack(side=tk.LEFT, padx=(0, 2))
            buttons = {"Open": b_open, "Closed": b_close}

        st = str(it.get("status", "") or "").strip()
        st_bg, st_fg = _status_style_for(kind, st)
        status_cell = col_frames[1]
        status_cell.configure(bg=st_bg)
        status_lbl = _make_cell_label(status_cell, st, st_bg, fg=st_fg)

        cells = {}
        cells["year"] = _make_cell_label(col_frames[2], str(it.get("tax_year", "") or ""), base_bg)
        cells["category"] = _make_cell_label(col_frames[3], str(it.get("category", "") or ""), base_bg)

        title_text = str(it.get("name" if kind == "file" else "title", "") or "")
        cells["title"] = _make_cell_label(col_frames[4], title_text, base_bg)

        if kind == "file":
            c5 = str(it.get("requested_date", "") or "")
            c6 = str(it.get("received_date", "") or "")
            c7 = "Yes" if it.get("repeat_next_year") else "No"
        elif kind == "reminder":
            c5 = "Yes" if it.get("applies_every_year") else "No"
            c6 = ""
            c7 = ""
        else:
            c5 = str(it.get("type", "") or "")
            c6 = str(it.get("opened_date", "") or "")
            c7 = str(it.get("closed_date", "") or "")

        cells["c5"] = _make_cell_label(col_frames[5], c5, base_bg)
        cells["c6"] = _make_cell_label(col_frames[6], c6, base_bg)
        cells["c7"] = _make_cell_label(col_frames[7], c7, base_bg)

        pri = str(it.get("priority", "") or "Normal")
        priority_lbl = _make_cell_label(col_frames[8], pri, base_bg, fg=_priority_fg(pri))
        note_lbl = _make_cell_label(col_frames[9], str(it.get("note", "") or ""), base_bg)

        labels = list(cells.values()) + [status_lbl, priority_lbl, note_lbl]

        if kind == "file":
            on_dbl = lambda iid: _edit_file_by_id(iid)
        elif kind == "reminder":
            on_dbl = lambda iid: _edit_reminder_by_id(iid)
        else:
            on_dbl = lambda iid: _edit_issue_by_id(iid)

        _bind_row_select(row, labels, kind, item_id, on_dbl)

        ref = {
            "frame": row,
            "base_bg": base_bg,
            "col_frames": col_frames,
            "btn_frm": btn_frm,
            "buttons": buttons,
            "status_cell": status_cell,
            "status_lbl": status_lbl,
            "priority_lbl": priority_lbl,
            "note_lbl": note_lbl,
            "cells": cells,
            "labels": labels,
            "status": st,
        }
        return ref

    def _refresh_kind_list(kind: str, body: tk.Frame, collection: str, empty_msg: str):
        refs = _row_refs(kind)
        for w in body.winfo_children():
            w.destroy()
        refs.clear()
        rows = 0
        for it in client.get(collection) or []:
            if not isinstance(it, dict) or not _matches_filter(it):
                continue
            ref = _build_data_row(body, kind, it, rows)
            if ref:
                refs[str(it.get("id", "") or "")] = ref
                rows += 1
        if rows == 0:
            ttk.Label(body, text=empty_msg, foreground="#6B7280").pack(anchor="w", padx=8, pady=12)
        _update_row_highlight(kind)

    def _refresh_files_list():
        _refresh_kind_list("file", files_body, "file_requests", "(no file requests yet)")

    def _refresh_reminders_list():
        _refresh_kind_list("reminder", reminders_body, "annual_reminders", "(no annual reminders yet)")

    def _refresh_issues_list():
        _refresh_kind_list("issue", issues_body, "client_issues", "(no issues yet)")

    def _refresh_current_tree():
        _refresh_filter_combos()
        _refresh_files_list()
        _refresh_reminders_list()
        _refresh_issues_list()

    def refresh_tracker():
        _refresh_summary()
        _refresh_current_tree()

    try:
        if not hasattr(app, "_tracker_tab_refreshers"):
            app._tracker_tab_refreshers = {}
        app._tracker_tab_refreshers[id(client)] = refresh_tracker
    except Exception:
        pass

    def _sync_row(kind: str, item_id: str):
        ref = _row_refs(kind).get(item_id)
        cur = _find_by_id(
            {"file": "file_requests", "reminder": "annual_reminders", "issue": "client_issues"}[kind],
            item_id,
        )
        if ref and cur:
            _fill_row_cells(ref, kind, cur)

    def _apply_file_status(item_id: str, status: str, *, log: bool = True) -> bool:
        cur = _find_by_id("file_requests", item_id)
        if not cur:
            return False
        if str(cur.get("status", "") or "").strip() == status:
            return False
        patch: dict = {"status": status, "updated_ts": now_ts()}
        if status in ("Needed", "Requested"):
            patch["received_date"] = ""
        if status == "Requested":
            if not str(cur.get("requested_date", "") or "").strip():
                patch["requested_date"] = today_str()
        elif status == "Received":
            patch["received_date"] = today_str()
            if not str(cur.get("requested_date", "") or "").strip():
                patch["requested_date"] = today_str()
        update_file_request(client, item_id, patch)
        if log:
            name = f"{cur.get('tax_year', '')} {cur.get('name', '')}".strip()
            if status == "Requested":
                append_tracker_log(client, f"Requested tracker file: {name}")
            elif status == "Received":
                append_tracker_log(client, f"Received tracker file: {name}")
            elif status == "Needed":
                append_tracker_log(client, f"Set tracker file to Needed: {name}")
        return True

    def _apply_issue_status(item_id: str, status: str, *, log: bool = True) -> bool:
        cur = _find_by_id("client_issues", item_id)
        if not cur:
            return False
        cur_st = str(cur.get("status", "") or "").strip()
        if status == "Open" and cur_st in ("Open", "Waiting on Client", "Waiting on Government"):
            return False
        if status == "Closed" and cur_st in ("Closed", "Resolved"):
            return False
        patch: dict = {"status": status, "updated_ts": now_ts()}
        if status == "Open":
            patch["closed_date"] = ""
            if not str(cur.get("opened_date", "") or "").strip():
                patch["opened_date"] = today_str()
        elif status == "Closed":
            if not str(cur.get("closed_date", "") or "").strip():
                patch["closed_date"] = today_str()
        update_client_issue(client, item_id, patch)
        if log:
            title = cur.get("title", "")
            if status == "Open":
                append_tracker_log(client, f"Reopened tracker issue: {title}")
            elif status == "Closed":
                append_tracker_log(client, f"Closed tracker issue: {title}")
        return True

    def _persist_item(kind: str, item_id: str):
        save_clients_cb(app.items)
        _refresh_summary()
        _sync_row(kind, item_id)
        _refresh_logs_tab()
        _refresh_external_summaries()

    def _file_status_click(item_id: str, status: str):
        if _apply_file_status(item_id, status):
            _persist_item("file", item_id)

    def _issue_status_click(item_id: str, status: str):
        if _apply_issue_status(item_id, status):
            _persist_item("issue", item_id)

    def _reminder_quick_click(item_id: str, key: str):
        if key == "Active":
            cur = _find_by_id("annual_reminders", item_id)
            if not cur:
                return
            if str(cur.get("status", "") or "").strip() == "Active":
                return
            update_annual_reminder(client, item_id, {"status": "Active", "archived": False})
            append_tracker_log(client, f"Reactivated tracker reminder: {cur.get('title', '')}")
            _persist_item("reminder", item_id)
        elif key == "Done":
            cur = _find_by_id("annual_reminders", item_id)
            if not cur:
                return
            if str(cur.get("status", "") or "").strip() == "Done This Year":
                return
            update_annual_reminder(client, item_id, {"status": "Done This Year"})
            append_tracker_log(client, f"Marked reminder done this year: {cur.get('title', '')}")
            _persist_item("reminder", item_id)

    def _edit_file_by_id(item_id: str):
        cur = _find_by_id("file_requests", item_id)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit File Request", "file", initial=cur)
        if result:
            update_file_request(client, item_id, result)
            _persist()

    def _edit_reminder_by_id(item_id: str):
        cur = _find_by_id("annual_reminders", item_id)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit Annual Reminder", "reminder", initial=cur)
        if result:
            update_annual_reminder(client, item_id, result)
            _persist()

    def _edit_issue_by_id(item_id: str):
        cur = _find_by_id("client_issues", item_id)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit Issue", "issue", initial=cur)
        if result:
            update_client_issue(client, item_id, result)
            _persist()

    def _add_file():
        result = TrackerItemDialog.open(
            app.winfo_toplevel(), "Add File Request", "file",
            initial={"tax_year": default_tax_year()},
        )
        if result:
            add_file_request(client, result)
            _persist()

    def _edit_file():
        iid = _selected_id("file")
        if not iid:
            messagebox.showinfo("Edit", "Select a file request first.")
            return
        _edit_file_by_id(iid)

    def _archive_file():
        iid = _selected_id("file")
        if not iid:
            messagebox.showinfo("Archive", "Select a file request first.")
            return
        if messagebox.askyesno("Archive", "Archive the selected file request?"):
            update_file_request(client, iid, {"archived": True, "status": "Archived"})
            append_tracker_log(client, "Archived tracker file request.")
            _persist()

    def _mark_requested():
        iid = _selected_id("file")
        if not iid:
            messagebox.showinfo("Mark Requested", "Select a file request first.")
            return
        if _apply_file_status(iid, "Requested"):
            _persist_item("file", iid)

    def _mark_received():
        iid = _selected_id("file")
        if not iid:
            messagebox.showinfo("Mark Received", "Select a file request first.")
            return
        if _apply_file_status(iid, "Received"):
            _persist_item("file", iid)

    def _roll_forward():
        today = datetime.date.today()
        from_year = simpledialog.askstring(
            "Roll Forward", "Copy repeat-next-year file requests from tax year:",
            initialvalue=str(today.year - 1), parent=tracker_tab,
        )
        if not from_year:
            return
        to_year = simpledialog.askstring(
            "Roll Forward", "Copy to tax year:",
            initialvalue=str(today.year), parent=tracker_tab,
        )
        if not to_year:
            return
        n = roll_forward_file_requests(client, from_year.strip(), to_year.strip())
        messagebox.showinfo("Roll Forward", f"Copied {n} file request(s).", parent=tracker_tab)
        _persist()

    def _add_reminder():
        result = TrackerItemDialog.open(
            app.winfo_toplevel(), "Add Annual Reminder", "reminder",
            initial={"tax_year": default_tax_year()},
        )
        if result:
            add_annual_reminder(client, result)
            _persist()

    def _edit_reminder():
        iid = _selected_id("reminder")
        if not iid:
            messagebox.showinfo("Edit", "Select a reminder first.")
            return
        _edit_reminder_by_id(iid)

    def _archive_reminder():
        iid = _selected_id("reminder")
        if not iid:
            messagebox.showinfo("Archive", "Select a reminder first.")
            return
        if messagebox.askyesno("Archive", "Archive the selected reminder?"):
            cur = _find_by_id("annual_reminders", iid)
            update_annual_reminder(client, iid, {"archived": True, "status": "Archived"})
            if cur:
                append_tracker_log(client, f"Archived tracker reminder: {cur.get('title', '')}")
            _persist()

    def _mark_done_year():
        iid = _selected_id("reminder")
        if not iid:
            messagebox.showinfo("Mark Done", "Select a reminder first.")
            return
        _reminder_quick_click(iid, "Done")

    def _reactivate_reminder():
        iid = _selected_id("reminder")
        if not iid:
            messagebox.showinfo("Reactivate", "Select a reminder first.")
            return
        _reminder_quick_click(iid, "Active")

    def _add_issue():
        result = TrackerItemDialog.open(
            app.winfo_toplevel(), "Add Issue", "issue",
            initial={"tax_year": default_tax_year()},
        )
        if result:
            add_client_issue(client, result)
            _persist()

    def _edit_issue():
        iid = _selected_id("issue")
        if not iid:
            messagebox.showinfo("Edit", "Select an issue first.")
            return
        _edit_issue_by_id(iid)

    def _archive_issue():
        iid = _selected_id("issue")
        if not iid:
            messagebox.showinfo("Archive", "Select an issue first.")
            return
        if messagebox.askyesno("Archive", "Archive the selected issue?"):
            cur = _find_by_id("client_issues", iid)
            update_client_issue(client, iid, {"archived": True, "status": "Archived"})
            if cur:
                append_tracker_log(client, f"Archived tracker issue: {cur.get('title', '')}")
            _persist()

    def _mark_resolved():
        iid = _selected_id("issue")
        if not iid:
            messagebox.showinfo("Mark Resolved", "Select an issue first.")
            return
        cur = _find_by_id("client_issues", iid)
        if not cur:
            return
        patch = {"status": "Resolved"}
        if not str(cur.get("closed_date", "") or "").strip():
            patch["closed_date"] = today_str()
        update_client_issue(client, iid, patch)
        append_tracker_log(client, f"Resolved tracker issue: {cur.get('title', '')}")
        _persist_item("issue", iid)

    def _reopen_issue():
        iid = _selected_id("issue")
        if not iid:
            messagebox.showinfo("Reopen", "Select an issue first.")
            return
        if _apply_issue_status(iid, "Open"):
            _persist_item("issue", iid)

    _ACTION_SETS = {
        "file": [
            ("Add", _add_file), ("Edit", _edit_file), ("Archive", _archive_file),
            ("Mark Requested", _mark_requested), ("Mark Received", _mark_received),
            ("Roll Forward", _roll_forward),
        ],
        "reminder": [
            ("Add", _add_reminder), ("Edit", _edit_reminder), ("Archive", _archive_reminder),
            ("Mark Done This Year", _mark_done_year), ("Reactivate", _reactivate_reminder),
        ],
        "issue": [
            ("Add", _add_issue), ("Edit", _edit_issue), ("Archive", _archive_issue),
            ("Mark Resolved", _mark_resolved), ("Reopen", _reopen_issue),
        ],
    }

    def _sync_action_bar():
        for w in action_bar.winfo_children():
            w.destroy()
        for text, cmd in _ACTION_SETS.get(_current_kind(), []):
            ttk.Button(action_bar, text=text, command=cmd).pack(side=tk.LEFT, padx=(0, 6))

    v_search.trace_add("write", lambda *_: _refresh_current_tree())
    v_year.trace_add("write", lambda *_: _refresh_current_tree())
    v_status.trace_add("write", lambda *_: _refresh_current_tree())
    v_category.trace_add("write", lambda *_: _refresh_current_tree())

    def _on_inner_tab_changed(_e=None):
        _refresh_filter_combos()
        _sync_action_bar()

    inner_nb.bind("<<NotebookTabChanged>>", _on_inner_tab_changed)

    for scroll, kind, tab in (
        (files_scroll, "file", files_tab),
        (reminders_scroll, "reminder", reminders_tab),
        (issues_scroll, "issue", issues_tab),
    ):
        scroll.canvas.bind(
            "<Button-1>",
            lambda e, k=kind: _clear_selection(k) if e.widget is scroll.canvas else None,
        )
        tab.bind(
            "<Button-1>",
            lambda e, k=kind: _clear_selection(k) if e.widget is tab else None,
        )

    _sync_action_bar()
    refresh_tracker()
    return tracker_tab
