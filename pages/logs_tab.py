from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import tkinter as tk
from tkinter import ttk, messagebox
import datetime
try:
    from vertex.utils.app_logging import get_logger

except ModuleNotFoundError:
    from utils.app_logging import get_logger

LOG = get_logger("logs")

try:
    from styles.new_ui import NewUI
except Exception:
    class NewUI:
        @staticmethod
        def stripe_tree(tv): pass


# LogDialog moved to ui/dialogs/logdialog.py
try:
    from vertex.ui.dialogs.logdialog import LogDialog
except ModuleNotFoundError:
    from ui.dialogs.logdialog import LogDialog

# Tab label (used for navigation restore compatibility)
LOGS_TAB_LABEL = "Logs"

# --- Theme & tag helpers ---
def _is_dark(app) -> bool:
    return bool(getattr(app, "is_dark_mode", False)
                or getattr(app, "dark_mode", False)
                or (getattr(app, "theme", "") in ("dark", "Dark")))


def _ensure_row_tags(tv: ttk.Treeview, dark: bool):
    tv.tag_configure('done_light', foreground='#6B7280', background='#F3F4F6')
    tv.tag_configure('done_dark',  foreground='#9CA3AF', background='#1F2937')
    try:
        import tkinter.font as tkfont
        base = tkfont.nametofont("TkTextFont")
        active_bold = base.copy()
        active_bold.configure(weight="bold")
        tv.tag_configure('active_light', foreground='#111827', background='#EEF2FF', font=active_bold)
        tv.tag_configure('active_dark',  foreground='#E5E7EB', background='#312E81', font=active_bold)
    except Exception:
        tv.tag_configure('active_light', foreground='#111827', background='#EEF2FF')
        tv.tag_configure('active_dark',  foreground='#E5E7EB', background='#312E81')

    tv.tag_configure('task_light', foreground='#1F2937', background='#F9FAFB')
    tv.tag_configure('task_dark',  foreground='#E5E7EB', background='#111827')

    return {
        "done":   'done_dark' if dark else 'done_light',
        "active": 'active_dark' if dark else 'active_light',
        "task":   'task_dark' if dark else 'task_light',
    }


def _norm_minute(ts: str) -> str:
    """Normalize timestamps to minute precision for dedupe keys."""
    s = str(ts or "").strip().replace("T", " ").replace("Z", "")
    return s[:16]


def _parse_sort_ts(raw: str) -> datetime.datetime:
    s = str(raw or "").strip()
    if not s:
        return datetime.datetime.min
    s_iso = s.replace("Z", "+00:00")
    try:
        return datetime.datetime.fromisoformat(s_iso)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s[:19], fmt)
        except ValueError:
            pass
    return datetime.datetime.min


def _split_history_text(text: str) -> tuple[str, str]:
    """Split 'Task | note' style history lines (same as finish flow)."""
    t = (text or "").strip()
    if " | " in t:
        a, b = t.split(" | ", 1)
        return a.strip(), b.strip()
    return t, ""


def _build_merged_rows(client: dict) -> list[dict]:
    """
    Build unified rows for the Logs tab Treeview.

    Each row dict:
      status, task, time_disp, note_disp, sort_ts, tag, meta
    meta: {kind, log_index?} for memo actions; kind in memo|history|work|active
    """
    rows: list[dict] = []

    # Keys to skip redundant history log lines when a completed work_item exists
    completed_keys: set[tuple[str, str]] = set()
    for wi in client.get("work_items") or []:
        if not isinstance(wi, dict):
            continue
        if str(wi.get("status", "") or "").strip().lower() != "completed":
            continue
        tn = (wi.get("task_name") or "").strip()
        ca = wi.get("completed_at") or wi.get("updated_at") or ""
        if tn:
            completed_keys.add((tn.casefold(), _norm_minute(str(ca))))

    # --- Active session (task bar)
    aw = client.get("active_work") or {}
    active_wid = ""
    if isinstance(aw, dict) and aw:
        active_wid = str(aw.get("work_item_id", "") or "").strip()
        tname = str(aw.get("task_name") or "").strip() or "—"
        since = str(aw.get("started_at") or aw.get("created_at") or "").strip() or "—"
        rows.append({
            "status": "Active",
            "task": tname,
            "time_disp": since,
            "note_disp": "",
            "sort_ts": _parse_sort_ts(since),
            "tag": "task",
            "meta": {"kind": "active"},
        })

    # --- Work items (held / completed / stray active)
    for wi in client.get("work_items") or []:
        if not isinstance(wi, dict):
            continue
        st = str(wi.get("status", "") or "").strip().lower()
        wid = str(wi.get("id", "") or "").strip()
        tname = str(wi.get("task_name", "") or "").strip()
        if not tname:
            tname = "—"

        if st == "active":
            if active_wid and wid == active_wid:
                continue
            ts = str(wi.get("started_at") or wi.get("updated_at") or wi.get("created_at") or "").strip() or "—"
            rows.append({
                "status": "Active",
                "task": tname,
                "time_disp": ts,
                "note_disp": str(wi.get("note") or "").replace("\n", " ").strip(),
                "sort_ts": _parse_sort_ts(ts),
                "tag": "task",
                "meta": {"kind": "work", "work_item_id": wid},
            })
        elif st == "on_hold":
            ts = str(wi.get("held_at") or wi.get("updated_at") or "").strip() or "—"
            note = str(wi.get("note") or "").replace("\n", " ").strip()
            rows.append({
                "status": "On hold",
                "task": tname,
                "time_disp": ts,
                "note_disp": note,
                "sort_ts": _parse_sort_ts(ts),
                "tag": "task",
                "meta": {"kind": "work", "work_item_id": wid},
            })
        elif st == "completed":
            ts = str(wi.get("completed_at") or wi.get("updated_at") or "").strip() or "—"
            note = str(wi.get("note") or "").replace("\n", " ").strip()
            rows.append({
                "status": "Finished",
                "task": tname,
                "time_disp": ts,
                "note_disp": note,
                "sort_ts": _parse_sort_ts(ts),
                "tag": "task",
                "meta": {"kind": "work", "work_item_id": wid},
            })

    # --- Log entries (memos + history)
    for i, entry in enumerate(client.get("logs") or []):
        if not isinstance(entry, dict):
            continue
        lt = str(entry.get("log_type", "") or "").strip().lower()
        if not lt:
            lt = "memo"
        ts_raw = entry.get("ts", "")
        ts_disp = str(ts_raw or "").strip()
        if entry.get("edited"):
            ts_disp = f"{ts_disp} (Edited)" if ts_disp else "(Edited)"
        text = str(entry.get("text", "") or "").strip()
        sort_ts = _parse_sort_ts(str(ts_raw))

        if lt == "history":
            task_part, note_part = _split_history_text(text)
            key = (task_part.casefold(), _norm_minute(str(ts_raw)))
            if key in completed_keys:
                continue
            rows.append({
                "status": "Finished",
                "task": task_part or "—",
                "time_disp": ts_disp,
                "note_disp": note_part,
                "sort_ts": sort_ts,
                "tag": "task",
                "meta": {"kind": "history", "log_index": i},
            })
        else:
            # Memo (and any non-history log)
            done = bool(entry.get("done"))
            st_label = "N/A"
            if done:
                st_label = "N/A ✓"
            rows.append({
                "status": st_label,
                "task": text or "—",
                "time_disp": ts_disp,
                "note_disp": "",
                "sort_ts": sort_ts,
                "tag": "done" if done else "active",
                "meta": {"kind": "memo", "log_index": i},
            })

    rows.sort(key=lambda r: r["sort_ts"], reverse=True)
    return rows


def init_logs_tab(notebook: ttk.Notebook, app, client: dict, save_clients_cb):
    """Client Logs tab: merged tasks + memos/history in one list."""
    logs_tab = ttk.Frame(notebook, padding=8)
    notebook.add(logs_tab, text=LOGS_TAB_LABEL)

    DARK = _is_dark(app)
    ROW_TAGS = None  # set after tv created

    cols = ("status", "task", "time", "note")
    _merged_style = ttk.Style(logs_tab)
    _merged_style.configure("MergedLogs.Treeview", rowheight=26)
    tv = ttk.Treeview(
        logs_tab,
        columns=cols,
        show="headings",
        selectmode="browse",
        height=14,
        style="MergedLogs.Treeview",
    )

    tv.heading("status", text="Status")
    tv.heading("task", text="Task / memo")
    tv.heading("time", text="Time")
    tv.heading("note", text="Note")

    tv.column("status", width=100, anchor="w", stretch=False)
    tv.column("task", width=280, anchor="w", stretch=True)
    tv.column("time", width=180, anchor="w", stretch=False)
    tv.column("note", width=260, anchor="w", stretch=True)

    yscr = ttk.Scrollbar(logs_tab, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=yscr.set)

    tv.grid(row=0, column=0, sticky="nsew")
    yscr.grid(row=0, column=1, sticky="ns")

    logs_tab.grid_rowconfigure(0, weight=1)
    logs_tab.grid_columnconfigure(0, weight=1)

    ROW_TAGS = _ensure_row_tags(tv, DARK)
    row_meta: dict[str, dict] = {}

    def refresh_merged():
        tv.delete(*tv.get_children())
        row_meta.clear()
        merged = _build_merged_rows(client)
        tag_apply: list[tuple[str, str]] = []
        if not merged:
            tv.insert("", "end", values=("—", "(no entries yet)", "", ""))
        else:
            for r in merged:
                iid = tv.insert(
                    "",
                    "end",
                    values=(r["status"], r["task"], r["time_disp"], r["note_disp"]),
                )
                row_meta[iid] = r["meta"]
                tag_apply.append((iid, r.get("tag") or "task"))
        try:
            NewUI.stripe_tree(tv)
        except Exception:
            pass
        # Re-apply semantic tags on top of zebra (stripe_tree overwrites tags).
        for iid, tag_key in tag_apply:
            stripes = tuple(tv.item(iid, "tags") or ())
            semantic = ROW_TAGS.get(tag_key, ROW_TAGS["task"])
            tv.item(iid, tags=stripes + (semantic,))

    # Expose refresh for task bar / Hold / Finished
    try:
        if not hasattr(app, "_logs_tab_refreshers"):
            app._logs_tab_refreshers = {}
        app._logs_tab_refreshers[id(client)] = refresh_merged
    except Exception:
        pass

    def _meta_for_selection():
        sel = tv.selection()
        if not sel:
            return None
        return row_meta.get(sel[0])

    def _memo_log_index_from_iid(iid: str) -> int | None:
        m = row_meta.get(iid) or {}
        if m.get("kind") != "memo":
            return None
        idx = m.get("log_index")
        return idx if isinstance(idx, int) else None

    def selected_memo_log_index():
        sel = tv.selection()
        if not sel:
            return None
        return _memo_log_index_from_iid(sel[0])

    def edit_log():
        i = selected_memo_log_index()
        if i is None:
            messagebox.showinfo("Edit log", "Select a memo row (Status N/A) to edit.")
            return
        entry = (client.get("logs") or [])[i]
        d = LogDialog(app.winfo_toplevel(), "Edit log entry", initial=entry)
        app.wait_window(d)
        if d.result:
            merged = dict(d.result)
            merged["log_type"] = str(entry.get("log_type", "memo") or "memo").strip().lower()
            client["logs"][i] = merged
            save_clients_cb(app.items)
            refresh_merged()

    def delete_log():
        i = selected_memo_log_index()
        if i is None:
            messagebox.showinfo("Delete log", "Select a memo row (Status N/A) to delete.")
            return
        if not messagebox.askyesno("Delete log", "Delete the selected memo entry?"):
            return
        del client["logs"][i]
        save_clients_cb(app.items)
        refresh_merged()

    def toggle_done():
        sel = tv.selection()
        if not sel:
            messagebox.showinfo("Toggle", "Select a memo row to toggle done / not done.")
            return
        _toggle_done_by_iid(sel[0])

    def _toggle_done_by_iid(iid: str):
        i = _memo_log_index_from_iid(iid)
        if i is None:
            return
        client.setdefault("logs", [])
        if 0 <= i < len(client["logs"]):
            client["logs"][i]["done"] = not bool(client["logs"][i].get("done"))
            save_clients_cb(app.items)
            refresh_merged()
            # re-select same logical row if still present
            for child in tv.get_children():
                m = row_meta.get(child) or {}
                if m.get("kind") == "memo" and m.get("log_index") == i:
                    tv.selection_set(child)
                    tv.focus(child)
                    break

    # Controls row
    controls = ttk.Frame(logs_tab)
    controls.grid(row=1, column=0, sticky="we", pady=(6, 0))
    controls.grid_columnconfigure(0, weight=1)

    v_quick = tk.StringVar()
    ttk.Entry(controls, textvariable=v_quick).grid(row=0, column=0, sticky="we")

    def add_note():
        t = v_quick.get().strip()
        if not t:
            return
        client.setdefault("logs", []).append({
            "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "text": t,
            "done": False,
            "log_type": "memo",
        })
        save_clients_cb(app.items)
        v_quick.set("")
        refresh_merged()

    ttk.Button(controls, text="Add log", command=add_note).grid(row=0, column=1, padx=(6, 0))
    ttk.Button(controls, text="Edit", command=edit_log).grid(row=0, column=2, padx=(6, 0))
    ttk.Button(controls, text="Delete", command=delete_log).grid(row=0, column=3, padx=(6, 0))
    ttk.Button(controls, text="Toggle Done", command=toggle_done).grid(row=0, column=4, padx=(6, 0))

    def on_tree_double(_e):
        edit_log()

    def on_tree_release(e):
        row = tv.identify_row(e.y)
        col = tv.identify_column(e.x)
        if row:
            tv.selection_set(row)
            tv.focus(row)
            # First data column = Status — click toggles done for memos
            if col == "#1":
                _toggle_done_by_iid(row)
        else:
            tv.selection_remove(tv.selection())
            tv.focus("")

    tv.unbind("<Button-1>")
    tv.unbind("<ButtonRelease-1>")
    tv.unbind("<Double-1>")

    tv.bind("<Double-1>", on_tree_double)
    tv.bind("<ButtonRelease-1>", on_tree_release)

    def clear_selection(_e=None):
        tv.selection_remove(tv.selection())
        tv.focus("")

    logs_tab.bind("<Button-1>", lambda e: clear_selection() if e.widget is logs_tab else None)
    controls.bind("<Button-1>", lambda e: clear_selection())

    def _on_tab_changed(_e=None):
        clear_selection()

    notebook.bind("<<NotebookTabChanged>>", _on_tab_changed)

    menu = tk.Menu(tv, tearoff=False)
    menu.add_command(label="Toggle Done", command=toggle_done)
    menu.add_command(label="Edit", command=edit_log)
    menu.add_command(label="Delete", command=delete_log)

    def popup(ev):
        row = tv.identify_row(ev.y)
        if row:
            tv.selection_set(row)
            m = row_meta.get(row) or {}
            st = tk.NORMAL if m.get("kind") == "memo" else tk.DISABLED
            try:
                menu.entryconfigure(0, state=st)
                menu.entryconfigure(1, state=st)
                menu.entryconfigure(2, state=st)
            except Exception:
                pass
            menu.tk_popup(ev.x_root, ev.y_root)

    tv.bind("<Button-3>", popup)

    refresh_merged()
