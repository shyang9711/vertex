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
    # DONE rows (de-emphasized but readable)
    tv.tag_configure('done_light', foreground='#6B7280', background='#F3F4F6')  # gray-500 on gray-100
    tv.tag_configure('done_dark',  foreground='#9CA3AF', background='#1F2937')  # gray-400 on gray-800

    # ACTIVE (unchecked) rows (slightly emphasized)
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

    return {
        "done":   'done_dark' if dark else 'done_light',
        "active": 'active_dark' if dark else 'active_light',
    }


def _task_detail_cell(wi: dict) -> str:
    """Second-line detail for an on-hold work item row."""
    held = str(wi.get("held_at") or wi.get("updated_at") or "").strip()
    note = str(wi.get("note") or "").replace("\n", " ").strip()
    if note:
        tail = note if len(note) <= 48 else note[:45] + "…"
        return f"{held} — {tail}" if held else tail
    return held or "—"


def init_logs_tab(notebook: ttk.Notebook, app, client: dict, save_clients_cb):
    """Client Logs tab: working-on tasks summary + memo/history log list."""
    logs_tab = ttk.Frame(notebook, padding=8)
    notebook.add(logs_tab, text=LOGS_TAB_LABEL)

    DARK = _is_dark(app)

    # ---- Working on tasks (read-only; use task bar above to Start / Hold / Finished)
    ttk.Label(logs_tab, text="Working on tasks", font=("Segoe UI", 10, "bold")).grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
    )
    task_cols = ("status", "task", "detail")
    _task_style = ttk.Style(logs_tab)
    _task_style.configure("TasksTab.Treeview", rowheight=24)
    tasks_tv = ttk.Treeview(
        logs_tab,
        columns=task_cols,
        show="headings",
        selectmode="browse",
        height=6,
        style="TasksTab.Treeview",
    )
    tasks_tv.heading("status", text="Status")
    tasks_tv.heading("task", text="Task")
    tasks_tv.heading("detail", text="Since / note")

    tasks_tv.column("status", width=110, anchor="w", stretch=False)
    tasks_tv.column("task", width=220, anchor="w", stretch=True)
    tasks_tv.column("detail", width=360, anchor="w", stretch=True)

    tasks_yscr = ttk.Scrollbar(logs_tab, orient="vertical", command=tasks_tv.yview)
    tasks_tv.configure(yscrollcommand=tasks_yscr.set)
    tasks_tv.grid(row=1, column=0, sticky="nsew")
    tasks_yscr.grid(row=1, column=1, sticky="ns")

    def refresh_tasks_tv():
        tasks_tv.delete(*tasks_tv.get_children())
        active = client.get("active_work") or {}
        if isinstance(active, dict) and active:
            tname = str(active.get("task_name") or "").strip() or "—"
            since = str(active.get("started_at") or active.get("created_at") or "").strip() or "—"
            tasks_tv.insert("", "end", values=("In progress", tname, since))
        held: list[dict] = []
        for wi in client.get("work_items") or []:
            if not isinstance(wi, dict):
                continue
            if str(wi.get("status", "") or "").strip().lower() != "on_hold":
                continue
            if not str(wi.get("task_name", "") or "").strip():
                continue
            held.append(wi)
        held.sort(key=lambda w: str(w.get("held_at") or w.get("updated_at") or ""), reverse=True)
        for wi in held:
            tasks_tv.insert(
                "",
                "end",
                values=("On hold", str(wi.get("task_name") or "").strip(), _task_detail_cell(wi)),
            )
        if not tasks_tv.get_children():
            tasks_tv.insert("", "end", values=("—", "(no active or held tasks)", ""))
        try:
            NewUI.stripe_tree(tasks_tv)
        except Exception:
            pass

    ttk.Separator(logs_tab).grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)

    ttk.Label(logs_tab, text="Log entries (memos & task history)", font=("Segoe UI", 10, "bold")).grid(
        row=3, column=0, columnspan=2, sticky="w", pady=(0, 4)
    )

    # ---- Log Treeview with checkbox-ish Done column
    cols = ("done", "ts", "text")
    _notes_tab_style = ttk.Style(logs_tab)
    _notes_tab_style.configure("NotesTab.Treeview", rowheight=28)
    tv = ttk.Treeview(
        logs_tab, columns=cols, show="headings", selectmode="browse", height=10, style="NotesTab.Treeview"
    )

    ROW_TAGS = _ensure_row_tags(tv, DARK)

    tv.heading("done", text="✓")
    tv.heading("ts",   text="When")
    tv.heading("text", text="Entry")

    tv.column("done", width=40,  anchor="center", stretch=False)
    tv.column("ts",   width=200, anchor="w",      stretch=False)
    tv.column("text", width=700, anchor="w",      stretch=True)

    yscr = ttk.Scrollbar(logs_tab, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=yscr.set)

    tv.grid(row=4, column=0, sticky="nsew")
    yscr.grid(row=4, column=1, sticky="ns")

    logs_tab.grid_rowconfigure(4, weight=1)
    logs_tab.grid_columnconfigure(0, weight=1)

    def _ts_display(entry):
        ts = entry.get("ts", "")
        if entry.get("edited"):
            return f"{ts} (Edited)" if ts else "(Edited)"
        return ts

    def refresh_tv():
        tv.delete(*tv.get_children())
        iids = []
        for entry in (client.get("logs") or []):
            done_mark = "☑" if entry.get("done") else "☐"
            iid = tv.insert("", "end", values=(
                done_mark,
                _ts_display(entry),
                entry.get("text", ""),
            ))
            iids.append((iid, bool(entry.get("done"))))
        try:
            NewUI.stripe_tree(tv)
        except Exception:
            pass
        
        for iid, is_done in iids:
            tv.item(iid, tags=(ROW_TAGS["done"] if is_done else ROW_TAGS["active"],))

    def refresh_all():
        refresh_tasks_tv()
        refresh_tv()

    # Expose refresh so task bar / Hold / Finished can update tasks + log list immediately.
    try:
        if not hasattr(app, "_logs_tab_refreshers"):
            app._logs_tab_refreshers = {}
        app._logs_tab_refreshers[id(client)] = refresh_all
    except Exception:
        pass

    def selected_index():
        sel = tv.selection()
        if not sel:
            return None
        return tv.index(sel[0])

    def edit_log():
        LOG.info("Edit log clicked")
        i = selected_index()
        if i is None:
            messagebox.showinfo("Edit log", "Select a log row to edit.")
            return
        entry = (client.get("logs") or [])[i]
        d = LogDialog(app.winfo_toplevel(), "Edit log entry", initial=entry)
        app.wait_window(d)
        if d.result:
            merged = dict(d.result)
            merged["log_type"] = str(entry.get("log_type", "memo") or "memo").strip().lower()
            client["logs"][i] = merged
            save_clients_cb(app.items)
            refresh_all()

    def delete_log():
        LOG.info("Delete log clicked")
        i = selected_index()
        if i is None:
            messagebox.showinfo("Delete log", "Select a log row to delete.")
            return
        if not messagebox.askyesno("Delete log", "Delete the selected log entry?"):
            return
        del client["logs"][i]
        save_clients_cb(app.items)
        refresh_all()

    def toggle_done():
        LOG.info("Toggle done clicked")
        sel = tv.selection()
        if not sel:
            messagebox.showinfo("Toggle", "Select a row to toggle done / not done.")
            return
        _toggle_done_by_iid(sel[0])


    def _toggle_done_by_iid(iid: str):
        try:
            idx = tv.index(iid)
        except Exception:
            return
        client.setdefault("logs", [])
        if 0 <= idx < len(client["logs"]):
            client["logs"][idx]["done"] = not bool(client["logs"][idx].get("done"))
            save_clients_cb(app.items)
            refresh_all()
            try:
                new_iid = tv.get_children()[idx]
                tv.selection_set(new_iid)
                tv.focus(new_iid)
            except Exception:
                pass


    # Controls row
    controls = ttk.Frame(logs_tab)
    controls.grid(row=5, column=0, sticky="we", pady=(6, 0))
    controls.grid_columnconfigure(0, weight=1)

    # Quick-add entry (memo)
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
        refresh_all()

    ttk.Button(controls, text="Add log",    command=add_note).grid(row=0, column=1, padx=(6,0))
    ttk.Button(controls, text="Edit",   command=edit_log).grid(row=0, column=2, padx=(6,0))
    ttk.Button(controls, text="Delete", command=delete_log).grid(row=0, column=3, padx=(6,0))
    ttk.Button(controls, text="Toggle Done", command=toggle_done).grid(row=0, column=4, padx=(6,0))

    def on_tree_double(_e):
        edit_log()

    def on_tree_release(e):
        row = tv.identify_row(e.y)
        col = tv.identify_column(e.x)
        if row:
            tv.selection_set(row)
            tv.focus(row)
            if col == "#1":
                _toggle_done_by_iid(row)
        else:
            tv.selection_remove(tv.selection())
            tv.focus('')

    tv.unbind("<Button-1>")
    tv.unbind("<ButtonRelease-1>")
    tv.unbind("<Double-1>")

    tv.bind("<Double-1>", on_tree_double)
    tv.bind("<ButtonRelease-1>", on_tree_release)

    def clear_selection(_e=None):
        tv.selection_remove(tv.selection())
        tv.focus('')

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
            menu.tk_popup(ev.x_root, ev.y_root)

    tv.bind("<Button-3>", popup)

    refresh_all()
