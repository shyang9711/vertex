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
_EMPTY_IID = "__empty__"
_ROW_SEL_BG = "#EEF2FF"
_ROW_BG = "#FFFFFF"


def _configure_tracker_tags(tv: ttk.Treeview) -> None:
    tv.tag_configure("empty", foreground="#6B7280")
    tv.tag_configure("tr_needed", background="#FEF3C7", foreground="#92400E")
    tv.tag_configure("tr_requested", background="#DBEAFE", foreground="#1E40AF")
    tv.tag_configure("tr_received", background="#D1FAE5", foreground="#065F46")
    tv.tag_configure("tr_waiting", background="#FFEDD5", foreground="#9A3412")
    tv.tag_configure("tr_open", background="#FEE2E2", foreground="#991B1B")
    tv.tag_configure("tr_active", background="#EEF2FF", foreground="#3730A3")
    tv.tag_configure("tr_closed", background="#F3F4F6", foreground="#6B7280")
    tv.tag_configure("tr_archived", background="#E5E7EB", foreground="#9CA3AF")
    tv.tag_configure("tr_default", foreground="")


def _configure_pill_styles(style: ttk.Style) -> None:
    """Rounded-outline pill buttons like the reference mockup."""
    style.configure(
        "Tracker.Pill.TButton",
        padding=(12, 5),
        background="#FFFFFF",
        foreground="#111827",
        borderwidth=1,
        relief="solid",
        bordercolor="#D1D5DB",
        lightcolor="#D1D5DB",
        darkcolor="#D1D5DB",
    )
    style.map(
        "Tracker.Pill.TButton",
        background=[("active", "#F9FAFB"), ("pressed", "#F3F4F6")],
        bordercolor=[("active", "#9CA3AF"), ("pressed", "#6B7280")],
    )
    style.configure(
        "Tracker.Pill.Active.TButton",
        padding=(12, 5),
        background="#F8FAFC",
        foreground="#111827",
        borderwidth=2,
        relief="solid",
        bordercolor="#6366F1",
        lightcolor="#6366F1",
        darkcolor="#6366F1",
    )
    style.map(
        "Tracker.Pill.Active.TButton",
        background=[("active", "#EEF2FF"), ("pressed", "#E0E7FF")],
    )


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
    _style.configure("Tracker.Treeview", rowheight=26)
    _configure_pill_styles(_style)

    files_tab = ttk.Frame(inner_nb, padding=4)
    reminders_tab = ttk.Frame(inner_nb, padding=4)
    issues_tab = ttk.Frame(inner_nb, padding=4)
    inner_nb.add(files_tab, text="File Requests")
    inner_nb.add(reminders_tab, text="Annual Reminders")
    inner_nb.add(issues_tab, text="Issues / Notices")

    files_scroll = ScrollFrame(files_tab, height=360)
    files_scroll.pack(fill="both", expand=True)
    files_inner = files_scroll.inner

    issues_scroll = ScrollFrame(issues_tab, height=360)
    issues_scroll.pack(fill="both", expand=True)
    issues_inner = issues_scroll.inner

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
            tv.column(
                col,
                width=widths.get(col, 100),
                anchor="w",
                stretch=(col in ("name", "title", "note")),
            )
        yscr = ttk.Scrollbar(parent, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=yscr.set)
        tv.grid(row=0, column=0, sticky="nsew")
        yscr.grid(row=0, column=1, sticky="ns")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        _configure_tracker_tags(tv)
        return tv

    reminders_tv = _make_tree(
        reminders_tab,
        ("status", "category", "title", "priority", "every_year", "note"),
        {
            "status": "Status", "category": "Category", "title": "Title",
            "priority": "Priority", "every_year": "Every Year", "note": "Note",
        },
        {"status": 100, "category": 100, "title": 200, "priority": 72, "every_year": 72, "note": 180},
    )

    files_selected_id: dict[str, str | None] = {"id": None}
    issues_selected_id: dict[str, str | None] = {"id": None}
    files_row_refs: dict[str, dict] = {}
    issues_row_refs: dict[str, dict] = {}
    reminders_row_meta: dict[str, str] = {}

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

    def _insert_empty(tv: ttk.Treeview, message: str, ncols: int):
        vals = [""] * ncols
        vals[0] = "—"
        if ncols > 1:
            vals[1] = message
        tv.insert("", "end", iid=_EMPTY_IID, values=tuple(vals), tags=("empty",))

    def _insert_empty(tv: ttk.Treeview, message: str, ncols: int):
        vals = [""] * ncols
        vals[0] = "—"
        if ncols > 1:
            vals[1] = message
        tv.insert("", "end", iid=_EMPTY_IID, values=tuple(vals), tags=("empty",))

    def _pill_style(active: bool) -> str:
        return "Tracker.Pill.Active.TButton" if active else "Tracker.Pill.TButton"

    def _update_file_row_highlight():
        sel = files_selected_id.get("id")
        for item_id, ref in files_row_refs.items():
            bg = _ROW_SEL_BG if item_id == sel else _ROW_BG
            ref["frame"].configure(bg=bg)
            for lbl in ref.get("labels", []):
                lbl.configure(bg=bg)

    def _update_issue_row_highlight():
        sel = issues_selected_id.get("id")
        for item_id, ref in issues_row_refs.items():
            bg = _ROW_SEL_BG if item_id == sel else _ROW_BG
            ref["frame"].configure(bg=bg)
            for lbl in ref.get("labels", []):
                lbl.configure(bg=bg)

    def _toggle_file_selection(item_id: str):
        if files_selected_id.get("id") == item_id:
            files_selected_id["id"] = None
        else:
            files_selected_id["id"] = item_id
        _update_file_row_highlight()

    def _toggle_issue_selection(item_id: str):
        if issues_selected_id.get("id") == item_id:
            issues_selected_id["id"] = None
        else:
            issues_selected_id["id"] = item_id
        _update_issue_row_highlight()

    def _clear_file_selection():
        files_selected_id["id"] = None
        _update_file_row_highlight()

    def _clear_issue_selection():
        issues_selected_id["id"] = None
        _update_issue_row_highlight()

    def _file_info_text(it: dict) -> str:
        parts = [
            str(it.get("tax_year", "") or ""),
            str(it.get("category", "") or ""),
            str(it.get("name", "") or ""),
            str(it.get("status", "") or ""),
        ]
        if it.get("requested_date"):
            parts.append(f"req {it.get('requested_date')}")
        if it.get("received_date"):
            parts.append(f"rcv {it.get('received_date')}")
        if it.get("note"):
            parts.append(str(it.get("note", ""))[:40])
        return "  ·  ".join(p for p in parts if p)

    def _issue_info_text(it: dict) -> str:
        parts = [
            str(it.get("tax_year", "") or ""),
            str(it.get("type", "") or ""),
            str(it.get("title", "") or ""),
            str(it.get("status", "") or ""),
        ]
        if it.get("note"):
            parts.append(str(it.get("note", ""))[:40])
        return "  ·  ".join(p for p in parts if p)

    def _sync_file_row_buttons(item_id: str):
        ref = files_row_refs.get(item_id)
        cur = _find_by_id("file_requests", item_id)
        if not ref or not cur:
            return
        st = str(cur.get("status", "") or "").strip()
        for status, btn in ref.get("buttons", {}).items():
            btn.configure(style=_pill_style(st == status))
        if ref.get("info"):
            ref["info"].configure(text=_file_info_text(cur))

    def _sync_issue_row_buttons(item_id: str):
        ref = issues_row_refs.get(item_id)
        cur = _find_by_id("client_issues", item_id)
        if not ref or not cur:
            return
        st = str(cur.get("status", "") or "").strip()
        open_active = st in ("Open", "Waiting on Client", "Waiting on Government")
        closed_active = st in ("Closed", "Resolved", "Archived")
        ref["buttons"]["Open"].configure(style=_pill_style(open_active))
        ref["buttons"]["Closed"].configure(style=_pill_style(closed_active))
        if ref.get("info"):
            ref["info"].configure(text=_issue_info_text(cur))

    def _build_file_row(it: dict):
        item_id = str(it.get("id", "") or "").strip()
        if not item_id:
            return
        st = str(it.get("status", "") or "").strip()
        row = tk.Frame(files_inner, bg=_ROW_BG, padx=2, pady=3)
        row.pack(fill="x")

        btn_frm = ttk.Frame(row)
        btn_frm.pack(side=tk.LEFT, padx=(0, 8))

        buttons: dict[str, ttk.Button] = {}
        for label, status in (("Needed", "Needed"), ("Requested", "Requested"), ("Received", "Received")):
            b = ttk.Button(
                btn_frm,
                text=label,
                style=_pill_style(st == status),
                width=max(8, len(label)),
                command=lambda s=status, iid=item_id: _file_status_click(iid, s),
            )
            b.pack(side=tk.LEFT, padx=(0, 4))
            buttons[status] = b

        info = tk.Label(row, text=_file_info_text(it), anchor="w", bg=_ROW_BG, fg="#374151")
        info.pack(side=tk.LEFT, fill="x", expand=True)

        def _row_body_click(_e, iid=item_id):
            _toggle_file_selection(iid)

        def _row_dbl_click(_e, iid=item_id):
            files_selected_id["id"] = iid
            _update_file_row_highlight()
            _edit_file()

        for w in (row, info):
            w.bind("<Button-1>", _row_body_click)
            w.bind("<Double-1>", _row_dbl_click)

        files_row_refs[item_id] = {"frame": row, "buttons": buttons, "info": info, "labels": [info]}

    def _build_issue_row(it: dict):
        item_id = str(it.get("id", "") or "").strip()
        if not item_id:
            return
        st = str(it.get("status", "") or "").strip()
        open_active = st in ("Open", "Waiting on Client", "Waiting on Government")
        closed_active = st in ("Closed", "Resolved", "Archived")

        row = tk.Frame(issues_inner, bg=_ROW_BG, padx=2, pady=3)
        row.pack(fill="x")

        btn_frm = ttk.Frame(row)
        btn_frm.pack(side=tk.LEFT, padx=(0, 8))

        btn_open = ttk.Button(
            btn_frm, text="Open", style=_pill_style(open_active), width=7,
            command=lambda iid=item_id: _issue_status_click(iid, "Open"),
        )
        btn_open.pack(side=tk.LEFT, padx=(0, 4))
        btn_close = ttk.Button(
            btn_frm, text="Close", style=_pill_style(closed_active), width=7,
            command=lambda iid=item_id: _issue_status_click(iid, "Closed"),
        )
        btn_close.pack(side=tk.LEFT, padx=(0, 4))

        info = tk.Label(row, text=_issue_info_text(it), anchor="w", bg=_ROW_BG, fg="#374151")
        info.pack(side=tk.LEFT, fill="x", expand=True)

        def _row_body_click(_e, iid=item_id):
            _toggle_issue_selection(iid)

        def _row_dbl_click(_e, iid=item_id):
            issues_selected_id["id"] = iid
            _update_issue_row_highlight()
            _edit_issue()

        for w in (row, info):
            w.bind("<Button-1>", _row_body_click)
            w.bind("<Double-1>", _row_dbl_click)

        issues_row_refs[item_id] = {
            "frame": row,
            "buttons": {"Open": btn_open, "Closed": btn_close},
            "info": info,
            "labels": [info],
        }

    def _refresh_files_list():
        for w in files_inner.winfo_children():
            w.destroy()
        files_row_refs.clear()
        rows = 0
        for it in client.get("file_requests") or []:
            if not isinstance(it, dict) or not _matches_filter(it):
                continue
            _build_file_row(it)
            rows += 1
        if rows == 0:
            ttk.Label(files_inner, text="(no file requests yet)", foreground="#6B7280").pack(anchor="w", pady=8)
        _update_file_row_highlight()

    def _refresh_reminders_tv():
        reminders_tv.delete(*reminders_tv.get_children())
        reminders_row_meta.clear()
        rows = 0
        for it in client.get("annual_reminders") or []:
            if not isinstance(it, dict) or not _matches_filter(it):
                continue
            rows += 1
            st = str(it.get("status", "") or "")
            iid = reminders_tv.insert(
                "",
                "end",
                values=(
                    st, it.get("category", ""), it.get("title", ""),
                    it.get("priority", ""), "Yes" if it.get("applies_every_year") else "No",
                    it.get("note", ""),
                ),
                tags=(reminder_status_tag(st),),
            )
            reminders_row_meta[iid] = str(it.get("id", "") or "")
        if rows == 0:
            _insert_empty(reminders_tv, "(no annual reminders yet)", 6)

    def _refresh_issues_list():
        for w in issues_inner.winfo_children():
            w.destroy()
        issues_row_refs.clear()
        rows = 0
        for it in client.get("client_issues") or []:
            if not isinstance(it, dict) or not _matches_filter(it):
                continue
            _build_issue_row(it)
            rows += 1
        if rows == 0:
            ttk.Label(issues_inner, text="(no issues yet)", foreground="#6B7280").pack(anchor="w", pady=8)
        _update_issue_row_highlight()

    def _refresh_current_tree():
        _refresh_filter_combos()
        _refresh_files_list()
        _refresh_reminders_tv()
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

    def _selected_file_id() -> str | None:
        return files_selected_id.get("id")

    def _selected_issue_id() -> str | None:
        return issues_selected_id.get("id")

    def _selected_reminder_id() -> str | None:
        sel = reminders_tv.selection()
        if not sel or sel[0] == _EMPTY_IID:
            return None
        return reminders_row_meta.get(sel[0]) or None

    def _find_by_id(collection: str, item_id: str) -> dict | None:
        for it in client.get(collection) or []:
            if isinstance(it, dict) and str(it.get("id", "") or "") == item_id:
                return it
        return None

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
            if status == "Requested":
                append_tracker_log(
                    client,
                    f"Requested tracker file: {cur.get('tax_year', '')} {cur.get('name', '')}".strip(),
                )
            elif status == "Received":
                append_tracker_log(
                    client,
                    f"Received tracker file: {cur.get('tax_year', '')} {cur.get('name', '')}".strip(),
                )
            elif status == "Needed":
                append_tracker_log(
                    client,
                    f"Set tracker file to Needed: {cur.get('tax_year', '')} {cur.get('name', '')}".strip(),
                )
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

    def _persist_file_item(item_id: str):
        save_clients_cb(app.items)
        _refresh_summary()
        _sync_file_row_buttons(item_id)
        _refresh_logs_tab()
        _refresh_external_summaries()

    def _persist_issue_item(item_id: str):
        save_clients_cb(app.items)
        _refresh_summary()
        _sync_issue_row_buttons(item_id)
        _refresh_logs_tab()
        _refresh_external_summaries()

    def _file_status_click(item_id: str, status: str):
        if _apply_file_status(item_id, status):
            _persist_file_item(item_id)

    def _issue_status_click(item_id: str, status: str):
        if _apply_issue_status(item_id, status):
            _persist_issue_item(item_id)

    def _add_file():
        result = TrackerItemDialog.open(
            app.winfo_toplevel(), "Add File Request", "file",
            initial={"tax_year": default_tax_year()},
        )
        if result:
            add_file_request(client, result)
            _persist()

    def _edit_file():
        iid = _selected_file_id()
        if not iid:
            messagebox.showinfo("Edit", "Select a file request first.")
            return
        cur = _find_by_id("file_requests", iid)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit File Request", "file", initial=cur)
        if result:
            update_file_request(client, iid, result)
            _persist()

    def _archive_file():
        iid = _selected_file_id()
        if not iid:
            messagebox.showinfo("Archive", "Select a file request first.")
            return
        if messagebox.askyesno("Archive", "Archive the selected file request?"):
            update_file_request(client, iid, {"archived": True, "status": "Archived"})
            append_tracker_log(client, "Archived tracker file request.")
            _persist()

    def _mark_requested():
        iid = _selected_file_id()
        if not iid:
            messagebox.showinfo("Mark Requested", "Select a file request first.")
            return
        if _apply_file_status(iid, "Requested"):
            _persist_file_item(iid)

    def _mark_received():
        iid = _selected_file_id()
        if not iid:
            messagebox.showinfo("Mark Received", "Select a file request first.")
            return
        if _apply_file_status(iid, "Received"):
            _persist_file_item(iid)

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
        iid = _selected_reminder_id()
        if not iid:
            messagebox.showinfo("Edit", "Select a reminder first.")
            return
        cur = _find_by_id("annual_reminders", iid)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit Annual Reminder", "reminder", initial=cur)
        if result:
            update_annual_reminder(client, iid, result)
            _persist()

    def _archive_reminder():
        iid = _selected_reminder_id()
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
        iid = _selected_reminder_id()
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
        iid = _selected_reminder_id()
        if not iid:
            messagebox.showinfo("Reactivate", "Select a reminder first.")
            return
        cur = _find_by_id("annual_reminders", iid)
        if not cur:
            return
        update_annual_reminder(client, iid, {"status": "Active", "archived": False})
        append_tracker_log(client, f"Reactivated tracker reminder: {cur.get('title', '')}")
        _persist()

    def _add_issue():
        result = TrackerItemDialog.open(
            app.winfo_toplevel(), "Add Issue", "issue",
            initial={"tax_year": default_tax_year()},
        )
        if result:
            add_client_issue(client, result)
            _persist()

    def _edit_issue():
        iid = _selected_issue_id()
        if not iid:
            messagebox.showinfo("Edit", "Select an issue first.")
            return
        cur = _find_by_id("client_issues", iid)
        if not cur:
            return
        result = TrackerItemDialog.open(app.winfo_toplevel(), "Edit Issue", "issue", initial=cur)
        if result:
            update_client_issue(client, iid, result)
            _persist()

    def _archive_issue():
        iid = _selected_issue_id()
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
        iid = _selected_issue_id()
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
        _persist_issue_item(iid)

    def _reopen_issue():
        iid = _selected_issue_id()
        if not iid:
            messagebox.showinfo("Reopen", "Select an issue first.")
            return
        if _apply_issue_status(iid, "Open"):
            _persist_issue_item(iid)

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

    def _on_filter_change(*_):
        _refresh_current_tree()

    v_search.trace_add("write", _on_filter_change)
    v_year.trace_add("write", _on_filter_change)
    v_status.trace_add("write", _on_filter_change)
    v_category.trace_add("write", _on_filter_change)

    def _on_inner_tab_changed(_e=None):
        _refresh_filter_combos()
        _sync_action_bar()

    inner_nb.bind("<<NotebookTabChanged>>", _on_inner_tab_changed)

    def _reminder_tree_release(e):
        row = reminders_tv.identify_row(e.y)
        if not row or row == _EMPTY_IID:
            reminders_tv.selection_remove(reminders_tv.selection())
            reminders_tv.focus("")
            return "break"
        cur_sel = reminders_tv.selection()
        if cur_sel and cur_sel[0] == row:
            reminders_tv.selection_remove(cur_sel)
            reminders_tv.focus("")
        else:
            reminders_tv.selection_set(row)
            reminders_tv.focus(row)

    def _clear_reminder_selection(_e=None):
        reminders_tv.selection_remove(reminders_tv.selection())
        reminders_tv.focus("")

    reminders_tv.unbind("<ButtonRelease-1>")
    reminders_tv.bind("<ButtonRelease-1>", _reminder_tree_release)
    reminders_tv.bind("<Double-1>", lambda _e: _edit_reminder() if _selected_reminder_id() else None)
    reminders_tab.bind("<Button-1>", lambda e: _clear_reminder_selection() if e.widget is reminders_tab else None)

    files_scroll.canvas.bind("<Button-1>", lambda e: _clear_file_selection() if e.widget is files_scroll.canvas else None)
    files_tab.bind("<Button-1>", lambda e: _clear_file_selection() if e.widget is files_tab else None)

    issues_scroll.canvas.bind("<Button-1>", lambda e: _clear_issue_selection() if e.widget is issues_scroll.canvas else None)
    issues_tab.bind("<Button-1>", lambda e: _clear_issue_selection() if e.widget is issues_tab else None)

    _sync_action_bar()
    refresh_tracker()
    return tracker_tab
