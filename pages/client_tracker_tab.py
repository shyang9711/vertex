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
        FILE_REQUEST_STATUSES,
        REMINDER_STATUSES,
        ISSUE_STATUSES,
        CATEGORIES,
        today_str,
        now_ts,
    )
    from vertex.ui.dialogs.tracker_item_dialog import TrackerItemDialog
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
        FILE_REQUEST_STATUSES,
        REMINDER_STATUSES,
        ISSUE_STATUSES,
        CATEGORIES,
        today_str,
        now_ts,
    )
    from ui.dialogs.tracker_item_dialog import TrackerItemDialog

TRACKER_TAB_LABEL = "Tracker"


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

    idx = client_idx if isinstance(client_idx, int) else None

    # ---- Summary row ----
    summary_var = tk.StringVar(value="")
    ttk.Label(tracker_tab, textvariable=summary_var, font=("Segoe UI", 10, "bold")).grid(
        row=0, column=0, sticky="w", pady=(0, 6)
    )

    # ---- Filter row ----
    filter_frm = ttk.Frame(tracker_tab)
    filter_frm.grid(row=1, column=0, sticky="ew", pady=(0, 6))
    filter_frm.columnconfigure(1, weight=1)

    v_search = tk.StringVar()
    v_year = tk.StringVar(value="All")
    v_status = tk.StringVar(value="All")
    v_category = tk.StringVar(value="All")

    ttk.Label(filter_frm, text="Search").grid(row=0, column=0, sticky="w", padx=(0, 4))
    search_entry = ttk.Entry(filter_frm, textvariable=v_search, width=18)
    search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))

    ttk.Label(filter_frm, text="Year").grid(row=0, column=2, sticky="w", padx=(0, 4))
    cb_year = ttk.Combobox(filter_frm, textvariable=v_year, width=8, state="readonly")
    cb_year.grid(row=0, column=3, sticky="w", padx=(0, 8))

    ttk.Label(filter_frm, text="Status").grid(row=0, column=4, sticky="w", padx=(0, 4))
    cb_status = ttk.Combobox(filter_frm, textvariable=v_status, width=14, state="readonly")
    cb_status.grid(row=0, column=5, sticky="w", padx=(0, 8))

    ttk.Label(filter_frm, text="Category").grid(row=0, column=6, sticky="w", padx=(0, 4))
    cb_category = ttk.Combobox(filter_frm, textvariable=v_category, width=14, state="readonly")
    cb_category.grid(row=0, column=7, sticky="w")

    # ---- Internal notebook ----
    inner_nb = ttk.Notebook(tracker_tab)
    inner_nb.grid(row=2, column=0, sticky="nsew")
    tracker_tab.grid_rowconfigure(2, weight=1)
    tracker_tab.grid_columnconfigure(0, weight=1)

    _style = ttk.Style(tracker_tab)
    _style.configure("Tracker.Treeview", rowheight=26)

    # Sub-tab frames
    files_tab = ttk.Frame(inner_nb, padding=4)
    reminders_tab = ttk.Frame(inner_nb, padding=4)
    issues_tab = ttk.Frame(inner_nb, padding=4)
    inner_nb.add(files_tab, text="File Requests")
    inner_nb.add(reminders_tab, text="Annual Reminders")
    inner_nb.add(issues_tab, text="Issues / Notices")

    def _make_tree(parent, columns: tuple[str, ...], headings: dict[str, str], widths: dict[str, int]):
        tv = ttk.Treeview(
            parent,
            columns=columns,
            show="headings",
            selectmode="browse",
            height=12,
            style="Tracker.Treeview",
        )
        for col in columns:
            tv.heading(col, text=headings.get(col, col))
            tv.column(col, width=widths.get(col, 100), anchor="w", stretch=(col in ("name", "title", "note")))
        yscr = ttk.Scrollbar(parent, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=yscr.set)
        tv.grid(row=0, column=0, sticky="nsew")
        yscr.grid(row=0, column=1, sticky="ns")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        return tv

    files_tv = _make_tree(
        files_tab,
        ("status", "year", "category", "name", "requested", "received", "repeat", "priority", "note"),
        {
            "status": "Status", "year": "Year", "category": "Category", "name": "Name",
            "requested": "Requested", "received": "Received", "repeat": "Repeat",
            "priority": "Priority", "note": "Note",
        },
        {"status": 100, "year": 56, "category": 100, "name": 140, "requested": 88, "received": 88,
         "repeat": 52, "priority": 72, "note": 160},
    )

    reminders_tv = _make_tree(
        reminders_tab,
        ("status", "category", "title", "priority", "every_year", "note"),
        {
            "status": "Status", "category": "Category", "title": "Title",
            "priority": "Priority", "every_year": "Every Year", "note": "Note",
        },
        {"status": 100, "category": 100, "title": 200, "priority": 72, "every_year": 72, "note": 180},
    )

    issues_tv = _make_tree(
        issues_tab,
        ("status", "year", "type", "title", "opened", "closed", "priority", "note"),
        {
            "status": "Status", "year": "Year", "type": "Type", "title": "Title",
            "opened": "Opened", "closed": "Closed", "priority": "Priority", "note": "Note",
        },
        {"status": 100, "year": 56, "type": 110, "title": 160, "opened": 88, "closed": 88,
         "priority": 72, "note": 140},
    )

    files_row_meta: dict[str, str] = {}
    reminders_row_meta: dict[str, str] = {}
    issues_row_meta: dict[str, str] = {}

    # ---- Control bars ----
    files_ctrl = ttk.Frame(files_tab)
    files_ctrl.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
    reminders_ctrl = ttk.Frame(reminders_tab)
    reminders_ctrl.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
    issues_ctrl = ttk.Frame(issues_tab)
    issues_ctrl.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _refresh_logs_tab():
        try:
            refresher_map = getattr(app, "_logs_tab_refreshers", {}) or {}
            fn = refresher_map.get(id(client))
            if callable(fn):
                fn()
        except Exception:
            pass

    def _persist():
        save_clients_cb(app.items)
        _refresh_summary()
        _refresh_current_tree()
        _refresh_logs_tab()

    def _refresh_summary():
        s = count_tracker_summary(client)
        summary_var.set(
            f"Missing: {s['missing']} | Requested: {s['requested']} | Received: {s['received']} | "
            f"Open Issues: {s['open_issues']} | IPP Open: {s['ipp_open']}"
        )

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

    def _current_kind() -> str:
        sel = inner_nb.select()
        if sel == str(reminders_tab):
            return "reminder"
        if sel == str(issues_tab):
            return "issue"
        return "file"

    def _current_tv() -> ttk.Treeview:
        k = _current_kind()
        if k == "reminder":
            return reminders_tv
        if k == "issue":
            return issues_tv
        return files_tv

    def _matches_filter(item: dict, kind: str) -> bool:
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
            parts = []
            for key in item:
                parts.append(str(item.get(key, "") or ""))
            hay = " ".join(parts).casefold()
            if q not in hay:
                return False

        return True

    def _refresh_files_tv():
        files_tv.delete(*files_tv.get_children())
        files_row_meta.clear()
        for it in client.get("file_requests") or []:
            if not isinstance(it, dict):
                continue
            if not _matches_filter(it, "file"):
                continue
            repeat = "Yes" if it.get("repeat_next_year") else "No"
            iid = files_tv.insert(
                "",
                "end",
                values=(
                    it.get("status", ""),
                    it.get("tax_year", ""),
                    it.get("category", ""),
                    it.get("name", ""),
                    it.get("requested_date", ""),
                    it.get("received_date", ""),
                    repeat,
                    it.get("priority", ""),
                    it.get("note", ""),
                ),
            )
            files_row_meta[iid] = str(it.get("id", "") or "")

    def _refresh_reminders_tv():
        reminders_tv.delete(*reminders_tv.get_children())
        reminders_row_meta.clear()
        for it in client.get("annual_reminders") or []:
            if not isinstance(it, dict):
                continue
            if not _matches_filter(it, "reminder"):
                continue
            every = "Yes" if it.get("applies_every_year") else "No"
            iid = reminders_tv.insert(
                "",
                "end",
                values=(
                    it.get("status", ""),
                    it.get("category", ""),
                    it.get("title", ""),
                    it.get("priority", ""),
                    every,
                    it.get("note", ""),
                ),
            )
            reminders_row_meta[iid] = str(it.get("id", "") or "")

    def _refresh_issues_tv():
        issues_tv.delete(*issues_tv.get_children())
        issues_row_meta.clear()
        for it in client.get("client_issues") or []:
            if not isinstance(it, dict):
                continue
            if not _matches_filter(it, "issue"):
                continue
            iid = issues_tv.insert(
                "",
                "end",
                values=(
                    it.get("status", ""),
                    it.get("tax_year", ""),
                    it.get("type", ""),
                    it.get("title", ""),
                    it.get("opened_date", ""),
                    it.get("closed_date", ""),
                    it.get("priority", ""),
                    it.get("note", ""),
                ),
            )
            issues_row_meta[iid] = str(it.get("id", "") or "")

    def _refresh_current_tree():
        _refresh_filter_combos()
        _refresh_files_tv()
        _refresh_reminders_tv()
        _refresh_issues_tv()

    def refresh_tracker():
        _refresh_summary()
        _refresh_current_tree()

    try:
        if not hasattr(app, "_tracker_tab_refreshers"):
            app._tracker_tab_refreshers = {}
        app._tracker_tab_refreshers[id(client)] = refresh_tracker
    except Exception:
        pass

    def _selected_item_id(tv: ttk.Treeview, meta: dict) -> str | None:
        sel = tv.selection()
        if not sel:
            return None
        return meta.get(sel[0]) or None

    def _find_by_id(collection: str, item_id: str) -> dict | None:
        for it in client.get(collection) or []:
            if isinstance(it, dict) and str(it.get("id", "") or "") == item_id:
                return it
        return None

    # ---- File actions ----
    def _add_file():
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Add File Request", "file")
        if not result:
            return
        add_file_request(client, result)
        _persist()

    def _edit_file():
        iid = _selected_item_id(files_tv, files_row_meta)
        if not iid:
            messagebox.showinfo("Edit", "Select a file request first.")
            return
        cur = _find_by_id("file_requests", iid)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit File Request", "file", initial=cur)
        if not result:
            return
        update_file_request(client, iid, result)
        _persist()

    def _archive_file():
        iid = _selected_item_id(files_tv, files_row_meta)
        if not iid:
            messagebox.showinfo("Archive", "Select a file request first.")
            return
        if not messagebox.askyesno("Archive", "Archive the selected file request?"):
            return
        update_file_request(client, iid, {"archived": True, "status": "Archived"})
        append_tracker_log(client, f"Archived tracker file request.")
        _persist()

    def _mark_requested():
        iid = _selected_item_id(files_tv, files_row_meta)
        if not iid:
            messagebox.showinfo("Mark Requested", "Select a file request first.")
            return
        cur = _find_by_id("file_requests", iid)
        if not cur:
            return
        patch = {"status": "Requested", "updated_ts": now_ts()}
        if not str(cur.get("requested_date", "") or "").strip():
            patch["requested_date"] = today_str()
        update_file_request(client, iid, patch)
        append_tracker_log(client, f"Requested tracker file: {cur.get('tax_year', '')} {cur.get('name', '')}".strip())
        _persist()

    def _mark_received():
        iid = _selected_item_id(files_tv, files_row_meta)
        if not iid:
            messagebox.showinfo("Mark Received", "Select a file request first.")
            return
        cur = _find_by_id("file_requests", iid)
        if not cur:
            return
        patch = {"status": "Received", "received_date": today_str(), "updated_ts": now_ts()}
        if not str(cur.get("requested_date", "") or "").strip():
            patch["requested_date"] = today_str()
        update_file_request(client, iid, patch)
        append_tracker_log(client, f"Received tracker file: {cur.get('tax_year', '')} {cur.get('name', '')}".strip())
        _persist()

    def _roll_forward():
        today = datetime.date.today()
        from_year = simpledialog.askstring(
            "Roll Forward",
            "Copy repeat-next-year file requests from tax year:",
            initialvalue=str(today.year - 1),
            parent=tracker_tab,
        )
        if not from_year:
            return
        to_year = simpledialog.askstring(
            "Roll Forward",
            "Copy to tax year:",
            initialvalue=str(today.year),
            parent=tracker_tab,
        )
        if not to_year:
            return
        n = roll_forward_file_requests(client, from_year.strip(), to_year.strip())
        messagebox.showinfo("Roll Forward", f"Copied {n} file request(s).", parent=tracker_tab)
        _persist()

    # ---- Reminder actions ----
    def _add_reminder():
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Add Annual Reminder", "reminder")
        if not result:
            return
        add_annual_reminder(client, result)
        _persist()

    def _edit_reminder():
        iid = _selected_item_id(reminders_tv, reminders_row_meta)
        if not iid:
            messagebox.showinfo("Edit", "Select a reminder first.")
            return
        cur = _find_by_id("annual_reminders", iid)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit Annual Reminder", "reminder", initial=cur)
        if not result:
            return
        update_annual_reminder(client, iid, result)
        _persist()

    def _archive_reminder():
        iid = _selected_item_id(reminders_tv, reminders_row_meta)
        if not iid:
            messagebox.showinfo("Archive", "Select a reminder first.")
            return
        if not messagebox.askyesno("Archive", "Archive the selected reminder?"):
            return
        cur = _find_by_id("annual_reminders", iid)
        update_annual_reminder(client, iid, {"archived": True, "status": "Archived"})
        if cur:
            append_tracker_log(client, f"Archived tracker reminder: {cur.get('title', '')}")
        _persist()

    def _mark_done_year():
        iid = _selected_item_id(reminders_tv, reminders_row_meta)
        if not iid:
            messagebox.showinfo("Mark Done", "Select a reminder first.")
            return
        cur = _find_by_id("annual_reminders", iid)
        if not cur:
            return
        update_annual_reminder(client, iid, {"status": "Done This Year"})
        append_tracker_log(client, f"Marked reminder done this year: {cur.get('title', '')}")
        _persist()

    def _reactivate_reminder():
        iid = _selected_item_id(reminders_tv, reminders_row_meta)
        if not iid:
            messagebox.showinfo("Reactivate", "Select a reminder first.")
            return
        cur = _find_by_id("annual_reminders", iid)
        if not cur:
            return
        update_annual_reminder(client, iid, {"status": "Active", "archived": False})
        append_tracker_log(client, f"Reactivated tracker reminder: {cur.get('title', '')}")
        _persist()

    # ---- Issue actions ----
    def _add_issue():
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Add Issue", "issue")
        if not result:
            return
        add_client_issue(client, result)
        _persist()

    def _edit_issue():
        iid = _selected_item_id(issues_tv, issues_row_meta)
        if not iid:
            messagebox.showinfo("Edit", "Select an issue first.")
            return
        cur = _find_by_id("client_issues", iid)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit Issue", "issue", initial=cur)
        if not result:
            return
        update_client_issue(client, iid, result)
        _persist()

    def _archive_issue():
        iid = _selected_item_id(issues_tv, issues_row_meta)
        if not iid:
            messagebox.showinfo("Archive", "Select an issue first.")
            return
        if not messagebox.askyesno("Archive", "Archive the selected issue?"):
            return
        cur = _find_by_id("client_issues", iid)
        update_client_issue(client, iid, {"archived": True, "status": "Archived"})
        if cur:
            append_tracker_log(client, f"Archived tracker issue: {cur.get('title', '')}")
        _persist()

    def _mark_resolved():
        iid = _selected_item_id(issues_tv, issues_row_meta)
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
        _persist()

    def _reopen_issue():
        iid = _selected_item_id(issues_tv, issues_row_meta)
        if not iid:
            messagebox.showinfo("Reopen", "Select an issue first.")
            return
        cur = _find_by_id("client_issues", iid)
        if not cur:
            return
        update_client_issue(client, iid, {"status": "Open", "closed_date": ""})
        append_tracker_log(client, f"Reopened tracker issue: {cur.get('title', '')}")
        _persist()

    # ---- Buttons ----
    for text, cmd in (
        ("Add", _add_file),
        ("Edit", _edit_file),
        ("Delete/Archive", _archive_file),
        ("Mark Requested", _mark_requested),
        ("Mark Received", _mark_received),
        ("Roll Forward", _roll_forward),
    ):
        ttk.Button(files_ctrl, text=text, command=cmd).pack(side=tk.LEFT, padx=(0, 6))

    for text, cmd in (
        ("Add", _add_reminder),
        ("Edit", _edit_reminder),
        ("Archive", _archive_reminder),
        ("Mark Done This Year", _mark_done_year),
        ("Reactivate", _reactivate_reminder),
    ):
        ttk.Button(reminders_ctrl, text=text, command=cmd).pack(side=tk.LEFT, padx=(0, 6))

    for text, cmd in (
        ("Add", _add_issue),
        ("Edit", _edit_issue),
        ("Archive", _archive_issue),
        ("Mark Resolved", _mark_resolved),
        ("Reopen", _reopen_issue),
    ):
        ttk.Button(issues_ctrl, text=text, command=cmd).pack(side=tk.LEFT, padx=(0, 6))

    def _on_filter_change(*_):
        _refresh_current_tree()

    v_search.trace_add("write", _on_filter_change)
    v_year.trace_add("write", _on_filter_change)
    v_status.trace_add("write", _on_filter_change)
    v_category.trace_add("write", _on_filter_change)

    def _on_inner_tab_changed(_e=None):
        _refresh_filter_combos()

    inner_nb.bind("<<NotebookTabChanged>>", _on_inner_tab_changed)

    def _on_edit_dbl(tv, meta, kind):
        def handler(_e=None):
            iid = _selected_item_id(tv, meta)
            if not iid:
                return
            if kind == "file":
                _edit_file()
            elif kind == "reminder":
                _edit_reminder()
            else:
                _edit_issue()
        return handler

    files_tv.bind("<Double-1>", _on_edit_dbl(files_tv, files_row_meta, "file"))
    reminders_tv.bind("<Double-1>", _on_edit_dbl(reminders_tv, reminders_row_meta, "reminder"))
    issues_tv.bind("<Double-1>", _on_edit_dbl(issues_tv, issues_row_meta, "issue"))

    refresh_tracker()
    return tracker_tab
