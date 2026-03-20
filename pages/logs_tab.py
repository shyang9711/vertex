from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import tkinter as tk
from tkinter import ttk, messagebox
import datetime

try:
    from vertex.utils.app_logging import get_logger
    from vertex.utils.logs_merge import build_merged_rows
except ModuleNotFoundError:
    from utils.app_logging import get_logger
    from utils.logs_merge import build_merged_rows

LOG = get_logger("logs")

try:
    from styles.new_ui import NewUI
except Exception:
    class NewUI:
        @staticmethod
        def stripe_tree(tv): pass

try:
    from vertex.ui.dialogs.logdialog import LogDialog
except ModuleNotFoundError:
    from ui.dialogs.logdialog import LogDialog

LOGS_TAB_LABEL = "Logs"


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


def _work_item_status(client: dict, work_item_id: str) -> str:
    wid = str(work_item_id or "").strip()
    for wi in client.get("work_items") or []:
        if isinstance(wi, dict) and str(wi.get("id", "") or "").strip() == wid:
            return str(wi.get("status", "") or "").strip().lower()
    return ""


def init_logs_tab(
    notebook: ttk.Notebook,
    app,
    client: dict,
    save_clients_cb,
    client_idx: int | None = None,
):
    logs_tab = ttk.Frame(notebook, padding=8)
    notebook.add(logs_tab, text=LOGS_TAB_LABEL)

    idx = client_idx if isinstance(client_idx, int) else None

    DARK = _is_dark(app)
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
        merged = build_merged_rows(client)
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
                row_meta[iid] = dict(r["meta"])
                tag_apply.append((iid, r.get("tag") or "task"))
        try:
            NewUI.stripe_tree(tv)
        except Exception:
            pass
        for iid, tag_key in tag_apply:
            stripes = tuple(tv.item(iid, "tags") or ())
            semantic = ROW_TAGS.get(tag_key, ROW_TAGS["task"])
            tv.item(iid, tags=stripes + (semantic,))
        sync_action_buttons()

    try:
        if not hasattr(app, "_logs_tab_refreshers"):
            app._logs_tab_refreshers = {}
        app._logs_tab_refreshers[id(client)] = refresh_merged
    except Exception:
        pass

    def _memo_log_index_from_iid(iid: str) -> int | None:
        m = row_meta.get(iid) or {}
        if m.get("kind") != "memo":
            return None
        i = m.get("log_index")
        return i if isinstance(i, int) else None

    def edit_log():
        sel = tv.selection()
        i = _memo_log_index_from_iid(sel[0]) if sel else None
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
        sel = tv.selection()
        i = _memo_log_index_from_iid(sel[0]) if sel else None
        if i is None:
            messagebox.showinfo("Delete log", "Select a memo row (Status N/A) to delete.")
            return
        if not messagebox.askyesno("Delete log", "Delete the selected memo entry?"):
            return
        del client["logs"][i]
        save_clients_cb(app.items)
        refresh_merged()

    def toggle_done_memo():
        sel = tv.selection()
        if not sel:
            messagebox.showinfo("Toggle", "Select a memo row.")
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
            for child in tv.get_children():
                m = row_meta.get(child) or {}
                if m.get("kind") == "memo" and m.get("log_index") == i:
                    tv.selection_set(child)
                    tv.focus(child)
                    break

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

    # ---- Dynamic actions (task vs memo) ----
    controls = ttk.Frame(logs_tab)
    controls.grid(row=1, column=0, sticky="we", pady=(6, 0))

    v_quick = tk.StringVar()
    controls.grid_columnconfigure(0, weight=1)
    ttk.Entry(controls, textvariable=v_quick).grid(row=0, column=0, sticky="ew", padx=(0, 8))

    btn_add = ttk.Button(controls, text="Add log", command=add_note)
    btn_add.grid(row=0, column=1, sticky="w")

    dyn = ttk.Frame(controls)
    dyn.grid(row=0, column=2, sticky="w", padx=(12, 0))

    btn_edit = ttk.Button(dyn, text="Edit", command=edit_log)
    btn_del = ttk.Button(dyn, text="Delete", command=delete_log)
    btn_toggle = ttk.Button(dyn, text="Mark done", command=toggle_done_memo)
    btn_start = ttk.Button(dyn, text="Start", command=lambda: _do_start())
    btn_rejoin = ttk.Button(dyn, text="Resume session", command=lambda: _do_rejoin())
    btn_hold = ttk.Button(dyn, text="Hold", command=lambda: _do_hold())
    btn_finish = ttk.Button(dyn, text="Finished", command=lambda: _do_finish())
    btn_unfinish = ttk.Button(dyn, text="Unfinish", command=lambda: _do_unfinish())

    _all_dyn = (btn_edit, btn_del, btn_toggle, btn_start, btn_rejoin, btn_hold, btn_finish, btn_unfinish)

    def _hide_dynamic():
        for b in _all_dyn:
            b.grid_forget()

    def _do_start():
        if idx is None:
            return
        sel = tv.selection()
        if not sel:
            return
        m = row_meta.get(sel[0]) or {}
        wid = str(m.get("work_item_id", "") or "").strip()
        if m.get("kind") == "work" and wid:
            app._work_resume_held_item(idx, wid)

    def _do_rejoin():
        if idx is None:
            return
        sel = tv.selection()
        if not sel:
            return
        m = row_meta.get(sel[0]) or {}
        wid = str(m.get("work_item_id", "") or "").strip()
        if wid:
            app._work_rehydrate_active_from_item(idx, wid)

    def _do_hold():
        if idx is None:
            return
        app._work_task_hold(idx)

    def _do_finish():
        if idx is None:
            return
        app._work_task_finish(idx)

    def _do_unfinish():
        if idx is None:
            return
        sel = tv.selection()
        if not sel:
            return
        m = row_meta.get(sel[0]) or {}
        wid = str(m.get("work_item_id", "") or "").strip()
        if wid:
            app._work_unfinish_item(idx, wid)

    def sync_action_buttons():
        _hide_dynamic()
        sel = tv.selection()
        if not sel:
            return
        m = row_meta.get(sel[0]) or {}
        kind = m.get("kind")
        col = 0

        def place(btn):
            nonlocal col
            btn.grid(row=0, column=col, padx=(0, 6))
            col += 1

        if kind == "memo":
            li = m.get("log_index")
            done = False
            if isinstance(li, int) and 0 <= li < len(client.get("logs") or []):
                done = bool(client["logs"][li].get("done"))
            btn_toggle.config(text="Mark undone" if done else "Mark done")
            place(btn_edit)
            place(btn_del)
            place(btn_toggle)
            return

        if idx is None or kind not in ("active", "work"):
            return

        wid = str(m.get("work_item_id", "") or "").strip()
        st = _work_item_status(client, wid) if wid else ""

        if kind == "active":
            place(btn_hold)
            place(btn_finish)
            return

        if kind == "work":
            if st == "on_hold" and wid:
                place(btn_start)
            elif st == "active":
                if not client.get("active_work") and wid:
                    place(btn_rejoin)
                if client.get("active_work"):
                    place(btn_hold)
                    place(btn_finish)
            elif st == "completed" and wid:
                place(btn_unfinish)

    def on_tree_double(e):
        row = tv.identify_row(e.y)
        if row and _memo_log_index_from_iid(row) is not None:
            edit_log()

    def on_tree_release(e):
        row = tv.identify_row(e.y)
        col = tv.identify_column(e.x)
        if row:
            tv.selection_set(row)
            tv.focus(row)
            if col == "#1":
                _toggle_done_by_iid(row)
            sync_action_buttons()
        else:
            tv.selection_remove(tv.selection())
            tv.focus("")
            sync_action_buttons()

    tv.bind("<<TreeviewSelect>>", lambda _e: sync_action_buttons())

    tv.unbind("<Button-1>")
    tv.unbind("<ButtonRelease-1>")
    tv.unbind("<Double-1>")
    tv.bind("<Double-1>", on_tree_double)
    tv.bind("<ButtonRelease-1>", on_tree_release)

    def clear_selection(_e=None):
        tv.selection_remove(tv.selection())
        tv.focus("")
        sync_action_buttons()

    logs_tab.bind("<Button-1>", lambda e: clear_selection() if e.widget is logs_tab else None)
    controls.bind("<Button-1>", lambda e: clear_selection())

    def _on_tab_changed(_e=None):
        clear_selection()

    notebook.bind("<<NotebookTabChanged>>", _on_tab_changed)

    menu = tk.Menu(tv, tearoff=False)
    menu.add_command(label="Toggle done", command=toggle_done_memo)
    menu.add_command(label="Edit", command=edit_log)
    menu.add_command(label="Delete", command=delete_log)

    def popup(ev):
        row = tv.identify_row(ev.y)
        if row:
            tv.selection_set(row)
            m = row_meta.get(row) or {}
            st = tk.NORMAL if m.get("kind") == "memo" else tk.DISABLED
            try:
                for i in range(3):
                    menu.entryconfigure(i, state=st)
            except Exception:
                pass
            menu.tk_popup(ev.x_root, ev.y_root)

    tv.bind("<Button-3>", popup)

    refresh_merged()
