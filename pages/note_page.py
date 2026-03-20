from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

try:
    from vertex.utils.app_logging import get_logger
    from vertex.utils.logs_merge import build_all_clients_merged_rows
except ModuleNotFoundError:
    from utils.app_logging import get_logger
    from utils.logs_merge import build_all_clients_merged_rows

try:
    from vertex.ui.dialogs.logdialog import LogDialog
except ModuleNotFoundError:
    from ui.dialogs.logdialog import LogDialog

try:
    from styles.new_ui import NewUI
except Exception:
    class NewUI:
        @staticmethod
        def stripe_tree(tv): pass

LOG = get_logger("note_page")


def tokenize(s: str) -> list[str]:
    s = (s or "").strip().lower()
    parts = re.split(r"[^a-z0-9@._&\-]+", s)
    return [p for p in parts if p]


def norm_text(s: str) -> str:
    return " ".join(tokenize(s))


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


class NotePage:
    """
    Global view: same merged model as the client Logs tab (tasks + memos + history),
    across all clients.
    """

    def __init__(self, app):
        self.app = app
        self.page = None
        self.tree = None
        self._row_meta: dict[str, dict] = {}

        self.q = tk.StringVar()
        self.mgr = tk.StringVar(value="All")
        self.done = tk.StringVar(value="All")
        self.add_client = tk.StringVar(value="")

    def ensure(self, parent: tk.Widget):
        if self.page and self.page.winfo_exists():
            self.refresh()
            return self.page

        self.page = ttk.Frame(parent)
        DARK = _is_dark(self.app)

        header = ttk.Frame(self.page, padding=(10, 8))
        header.pack(fill=tk.X)

        ttk.Label(header, text="Notes", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT)

        ttk.Label(header, text="Manager:").pack(side=tk.LEFT, padx=(14, 4))
        self.cmb_mgr = ttk.Combobox(header, textvariable=self.mgr, width=16, state="readonly")
        self.cmb_mgr.pack(side=tk.LEFT)

        ttk.Label(header, text="Memo status:").pack(side=tk.LEFT, padx=(10, 4))
        self.cmb_done = ttk.Combobox(
            header, textvariable=self.done, width=10, state="readonly",
            values=("All", "Open", "Done"),
        )
        self.cmb_done.pack(side=tk.LEFT)

        ttk.Label(header, text="Add memo to:").pack(side=tk.LEFT, padx=(16, 4))
        self.cmb_add_client = ttk.Combobox(header, textvariable=self.add_client, width=28, state="readonly")
        self.cmb_add_client.pack(side=tk.LEFT)

        ttk.Label(header, text="Search").pack(side=tk.RIGHT, padx=(6, 4))
        ent = ttk.Entry(header, textvariable=self.q, width=36)
        ent.pack(side=tk.RIGHT)
        ent.bind("<KeyRelease>", lambda _e: self.refresh())

        self.cmb_mgr.bind("<<ComboboxSelected>>", lambda _e: self.refresh())
        self.cmb_done.bind("<<ComboboxSelected>>", lambda _e: self.refresh())

        cols = ("status", "client", "task", "time", "note", "mgr")
        tree_frm = ttk.Frame(self.page)
        tree_frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        tree_frm.columnconfigure(0, weight=1)
        tree_frm.rowconfigure(0, weight=1)
        _st = ttk.Style(self.page)
        _st.configure("GlobalNotes.Treeview", rowheight=26)
        self.tree = ttk.Treeview(
            tree_frm, columns=cols, show="headings", height=20, style="GlobalNotes.Treeview"
        )
        vsb = ttk.Scrollbar(tree_frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.heading("status", text="Status")
        self.tree.heading("client", text="Client")
        self.tree.heading("task", text="Task / memo")
        self.tree.heading("time", text="Time")
        self.tree.heading("note", text="Note")
        self.tree.heading("mgr", text="Manager")

        self.tree.column("status", width=92, anchor="w", stretch=False)
        self.tree.column("client", width=200, anchor="w", stretch=True)
        self.tree.column("task", width=240, anchor="w", stretch=True)
        self.tree.column("time", width=160, anchor="w", stretch=False)
        self.tree.column("note", width=220, anchor="w", stretch=True)
        self.tree.column("mgr", width=120, anchor="w", stretch=False)

        self.ROW_TAGS = _ensure_row_tags(self.tree, DARK)

        ctl = ttk.Frame(self.page, padding=(10, 0, 10, 10))
        ctl.pack(fill=tk.X)
        ctl.grid_columnconfigure(0, weight=1)
        self.v_quick = tk.StringVar()
        ttk.Entry(ctl, textvariable=self.v_quick).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.btn_add = ttk.Button(ctl, text="Add log", command=self._add_memo)
        self.btn_add.grid(row=0, column=1, sticky="w")
        self.dyn = ttk.Frame(ctl)
        self.dyn.grid(row=0, column=2, sticky="w", padx=(12, 0))

        self._mk_dyn_buttons()
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._sync_buttons())
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<ButtonRelease-1>", self._on_release)

        self._rebuild_mgr_values()
        self._rebuild_add_client_values()
        self.refresh()
        return self.page

    def _mk_dyn_buttons(self):
        self.btn_edit = ttk.Button(self.dyn, text="Edit", command=self._edit_memo)
        self.btn_del = ttk.Button(self.dyn, text="Delete", command=self._delete_memo)
        self.btn_toggle = ttk.Button(self.dyn, text="Mark done", command=self._toggle_memo)
        self.btn_start = ttk.Button(self.dyn, text="Start", command=self._do_start)
        self.btn_rejoin = ttk.Button(self.dyn, text="Resume session", command=self._do_rejoin)
        self.btn_hold = ttk.Button(self.dyn, text="Hold", command=self._do_hold)
        self.btn_finish = ttk.Button(self.dyn, text="Finished", command=self._do_finish)
        self.btn_unfinish = ttk.Button(self.dyn, text="Unfinish", command=self._do_unfinish)
        self._all_dyn = (
            self.btn_edit, self.btn_del, self.btn_toggle,
            self.btn_start, self.btn_rejoin, self.btn_hold, self.btn_finish, self.btn_unfinish,
        )

    def _hide_dyn(self):
        for b in self._all_dyn:
            b.grid_forget()

    def _client_at(self, idx: int | None):
        if idx is None or not (0 <= idx < len(getattr(self.app, "items", []) or [])):
            return None
        return self.app.items[idx]

    def _mk_meta(self, row: dict) -> dict:
        m = dict(row.get("meta") or {})
        m["client_idx"] = row.get("client_idx")
        return m

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

    def _rebuild_add_client_values(self):
        names = [(c.get("name") or "").strip() for c in getattr(self.app, "items", []) or []]
        names = [n for n in names if n]
        try:
            self.cmb_add_client.configure(values=names)
        except Exception:
            pass
        if names and not self.add_client.get():
            self.add_client.set(names[0])

    def refresh(self):
        if not self.tree:
            return
        self._rebuild_mgr_values()
        self._rebuild_add_client_values()

        q_norm = norm_text(self.q.get())
        mgr_f = self.mgr.get()
        done_f = self.done.get()

        self.tree.delete(*self.tree.get_children())
        self._row_meta.clear()

        items = getattr(self.app, "items", []) or []
        merged = build_all_clients_merged_rows(items)
        tag_apply: list[tuple[str, str]] = []

        for row in merged:
            cidx = row.get("client_idx")
            c = self._client_at(cidx)
            if not c:
                continue
            mgr = (c.get("acct_mgr", "") or "").strip()
            if mgr_f != "All" and mgr != mgr_f:
                continue

            kind = (row.get("meta") or {}).get("kind")
            if kind == "memo":
                li = (row.get("meta") or {}).get("log_index")
                if isinstance(li, int) and 0 <= li < len(c.get("logs") or []):
                    is_done = bool(c["logs"][li].get("done"))
                    if done_f == "Done" and not is_done:
                        continue
                    if done_f == "Open" and is_done:
                        continue

            if q_norm:
                hay = " ".join([
                    norm_text(row.get("client_name", "")),
                    norm_text(row.get("task", "")),
                    norm_text(row.get("note_disp", "")),
                    norm_text(row.get("time_disp", "")),
                    norm_text(mgr),
                ])
                if not all(tok in hay for tok in q_norm.split()):
                    continue

            iid = self.tree.insert(
                "",
                "end",
                values=(
                    row["status"],
                    row.get("client_name", ""),
                    row["task"],
                    row["time_disp"],
                    row["note_disp"],
                    mgr,
                ),
            )
            self._row_meta[iid] = self._mk_meta(row)
            tag_apply.append((iid, row.get("tag") or "task"))

        try:
            NewUI.stripe_tree(self.tree)
        except Exception:
            pass
        for iid, tag_key in tag_apply:
            stripes = tuple(self.tree.item(iid, "tags") or ())
            sem = self.ROW_TAGS.get(tag_key, self.ROW_TAGS["task"])
            self.tree.item(iid, tags=stripes + (sem,))
        self._sync_buttons()

    def _sel_meta(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self._row_meta.get(sel[0])

    def _sync_buttons(self):
        self._hide_dyn()
        m = self._sel_meta()
        if not m:
            return
        kind = m.get("kind")
        cidx = m.get("client_idx")
        c = self._client_at(cidx)
        col = 0

        def place(btn):
            nonlocal col
            btn.grid(row=0, column=col, padx=(0, 6))
            col += 1

        if kind == "memo" and isinstance(cidx, int):
            li = m.get("log_index")
            done = False
            if c and isinstance(li, int) and 0 <= li < len(c.get("logs") or []):
                done = bool(c["logs"][li].get("done"))
            self.btn_toggle.config(text="Mark undone" if done else "Mark done")
            place(self.btn_edit)
            place(self.btn_del)
            place(self.btn_toggle)
            return

        if not isinstance(cidx, int) or not c or kind not in ("active", "work"):
            return

        wid = str(m.get("work_item_id", "") or "").strip()
        st = _work_item_status(c, wid) if wid else ""

        if kind == "active":
            place(self.btn_hold)
            place(self.btn_finish)
            return

        if kind == "work":
            if st == "on_hold" and wid:
                place(self.btn_start)
            elif st == "active":
                if not c.get("active_work") and wid:
                    place(self.btn_rejoin)
                if c.get("active_work"):
                    place(self.btn_hold)
                    place(self.btn_finish)
            elif st == "completed" and wid:
                place(self.btn_unfinish)

    def _memo_index(self, m: dict) -> int | None:
        if m.get("kind") != "memo":
            return None
        i = m.get("log_index")
        return i if isinstance(i, int) else None

    def _add_memo(self):
        name = self.add_client.get().strip()
        if not name:
            messagebox.showinfo("Add log", "Choose a client in 'Add memo to'.")
            return
        t = self.v_quick.get().strip()
        if not t:
            return
        items = getattr(self.app, "items", []) or []
        idx = next((i for i, c in enumerate(items) if (c.get("name") or "").strip() == name), None)
        if idx is None:
            return
        c = items[idx]
        c.setdefault("logs", []).append({
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "text": t,
            "done": False,
            "log_type": "memo",
        })
        self._persist()
        self.v_quick.set("")
        self.refresh()

    def _edit_memo(self):
        m = self._sel_meta()
        if not m:
            return
        cidx = m.get("client_idx")
        li = self._memo_index(m)
        c = self._client_at(cidx)
        if c is None or li is None:
            return
        entry = c["logs"][li]
        d = LogDialog(self.app.winfo_toplevel(), "Edit log entry", initial=entry)
        self.app.wait_window(d)
        if d.result:
            merged = dict(d.result)
            merged["log_type"] = str(entry.get("log_type", "memo") or "memo").strip().lower()
            c["logs"][li] = merged
            self._persist()
            self.refresh()

    def _delete_memo(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        li = self._memo_index(m)
        c = self._client_at(cidx)
        if c is None or li is None:
            return
        if not messagebox.askyesno("Delete log", "Delete the selected memo entry?"):
            return
        del c["logs"][li]
        self._persist()
        self.refresh()

    def _toggle_memo(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        li = self._memo_index(m)
        c = self._client_at(cidx)
        if c is None or li is None:
            return
        c["logs"][li]["done"] = not bool(c["logs"][li].get("done"))
        self._persist()
        self.refresh()

    def _do_start(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        wid = str(m.get("work_item_id", "") or "").strip()
        if isinstance(cidx, int) and wid:
            self.app._work_resume_held_item(cidx, wid)
            self.refresh()

    def _do_rejoin(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        wid = str(m.get("work_item_id", "") or "").strip()
        if isinstance(cidx, int) and wid:
            self.app._work_rehydrate_active_from_item(cidx, wid)
            self.refresh()

    def _do_hold(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        if isinstance(cidx, int):
            self.app._work_task_hold(cidx)
            self.refresh()

    def _do_finish(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        if isinstance(cidx, int):
            self.app._work_task_finish(cidx)
            self.refresh()

    def _do_unfinish(self):
        m = self._sel_meta()
        cidx = m.get("client_idx")
        wid = str(m.get("work_item_id", "") or "").strip()
        if isinstance(cidx, int) and wid:
            self.app._work_unfinish_item(cidx, wid)
            self.refresh()

    def _persist(self):
        if hasattr(self.app, "save_clients_data"):
            self.app.save_clients_data()
        if hasattr(self.app, "refresh_all_logs_tabs"):
            self.app.refresh_all_logs_tabs()

    def _on_double_click(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid:
            return
        m = self._row_meta.get(iid) or {}
        if self._memo_index(m) is not None:
            self.tree.selection_set(iid)
            self._edit_memo()
            return
        cidx = m.get("client_idx")
        if isinstance(cidx, int):
            self.app.navigate("detail", cidx, push=True)
            if hasattr(self.app, "select_detail_tab"):
                try:
                    self.app.select_detail_tab("Logs")
                except Exception:
                    pass

    def _on_release(self, e):
        iid = self.tree.identify_row(e.y)
        if iid:
            self.tree.selection_set(iid)
        self._sync_buttons()
