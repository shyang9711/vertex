from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
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

        

class LogDialog(tk.Toplevel):
    def __init__(self, master, title="Log Entry", initial=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        init = initial or {}

        self.v_ts   = tk.StringVar(value=init.get("ts",""))
        self.v_user = tk.StringVar(value=init.get("user",""))
        self.v_done = tk.BooleanVar(value=bool(init.get("done", False)))

        frm = ttk.Frame(self, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        def row(lbl, var, r, width=40):
            ttk.Label(frm, text=lbl).grid(row=r, column=0, sticky="w", pady=(0,2))
            ent = ttk.Entry(frm, textvariable=var, width=width)
            ent.grid(row=r+1, column=0, sticky="we", pady=(0,6))
            return ent

        e1 = row("Timestamp (ISO, optional)", self.v_ts, 0, 40)
        row("User/Initials", self.v_user, 2, 20)

        done_box = ttk.Checkbutton(frm, text="Done", variable=self.v_done)
        done_box.grid(row=4, column=0, sticky="w", pady=(0,6))

        ttk.Label(frm, text="Text").grid(row=6, column=0, sticky="w", pady=(0,2))
        self.txt = ScrolledText(frm, width=56, height=8, wrap="word")
        self.txt.grid(row=7, column=0, sticky="nsew")
        if "text" in init:
            self.txt.insert("1.0", init["text"])

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, sticky="e", pady=(8,0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda _e: self._save())
        self.bind("<Escape>", lambda _e: self.destroy())

        self.after(50, e1.focus_set)
        self.grab_set()
        self.transient(master)

    def _save(self):
        ts = self.v_ts.get().strip()
        if not ts:
            ts = datetime.datetime.now().isoformat(timespec="seconds")
        self.result = {
            "ts": ts,
            "user": self.v_user.get().strip(),
            "text": self.txt.get("1.0", "end").strip(),
            "done": bool(self.v_done.get()),
        }
        self.destroy()

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
    # Using bold to avoid relying on color only
    # Note: Treeview tags can set 'font' if a named font exists; fallback to foreground/background if you don't manage fonts.
    try:
        import tkinter.font as tkfont
        base = tkfont.nametofont("TkTextFont")
        active_bold = base.copy()
        active_bold.configure(weight="bold")
        tv.tag_configure('active_light', foreground='#111827', background='#EEF2FF', font=active_bold)  # near-black on indigo-50
        tv.tag_configure('active_dark',  foreground='#E5E7EB', background='#312E81', font=active_bold)  # gray-200 on indigo-900
    except Exception:
        # Fallback if fonts aren't available
        tv.tag_configure('active_light', foreground='#111827', background='#EEF2FF')
        tv.tag_configure('active_dark',  foreground='#E5E7EB', background='#312E81')

    return {
        "done":   'done_dark' if dark else 'done_light',
        "active": 'active_dark' if dark else 'active_light',
    }

def init_logs_tab(notebook: ttk.Notebook, app, client: dict, save_clients_cb):
    """Builds the Logs tab for a given client inside the provided notebook."""
    logs_tab = ttk.Frame(notebook, padding=8)
    notebook.add(logs_tab, text="Logs")

    # ---- Treeview with checkbox-ish Done column ----
    cols = ("done", "ts", "user", "text")
    tv = ttk.Treeview(
        logs_tab, columns=cols, show="headings", selectmode="browse", height=10
    )

    DARK = _is_dark(app)
    ROW_TAGS = _ensure_row_tags(tv, DARK)


    tv.heading("done", text="✓")
    tv.heading("ts",   text="When")
    tv.heading("user", text="User")
    tv.heading("text", text="Entry")

    tv.column("done", width=40,  anchor="center", stretch=False)
    tv.column("ts",   width=160, anchor="w",      stretch=False)
    tv.column("user", width=80,  anchor="w",      stretch=False)
    tv.column("text", width=700, anchor="w",      stretch=True)

    yscr = ttk.Scrollbar(logs_tab, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=yscr.set)

    tv.grid(row=0, column=0, sticky="nsew")
    yscr.grid(row=0, column=1, sticky="ns")
    logs_tab.grid_rowconfigure(0, weight=1)
    logs_tab.grid_columnconfigure(0, weight=1)

    def refresh_tv():
        tv.delete(*tv.get_children())
        iids = []
        for entry in (client.get("logs") or []):
            done_mark = "☑" if entry.get("done") else "☐"
            iid = tv.insert("", "end", values=(
                done_mark,
                entry.get("ts", ""),
                entry.get("user", ""),
                entry.get("text", ""),
            ))
            iids.append((iid, bool(entry.get("done"))))
        try:
            NewUI.stripe_tree(tv)
        except Exception:
            pass
        
        for iid, is_done in iids:
            tv.item(iid, tags=(ROW_TAGS["done"] if is_done else ROW_TAGS["active"],))

    def selected_index():
        sel = tv.selection()
        if not sel:
            return None
        # We match by index to keep it simple & robust
        return tv.index(sel[0])

    def add_log():
        LOG.info("Add log clicked")
        d = LogDialog(app.winfo_toplevel(), "Add Log")
        app.wait_window(d)
        if d.result:
            client.setdefault("logs", []).append(d.result)
            save_clients_cb(app.items)
            refresh_tv()

    def edit_log():
        LOG.info("Edit log clicked")
        i = selected_index()
        if i is None:
            messagebox.showinfo("Edit Log", "Select a log row to edit.")
            return
        entry = (client.get("logs") or [])[i]
        d = LogDialog(app.winfo_toplevel(), "Edit Log", initial=entry)
        app.wait_window(d)
        if d.result:
            client["logs"][i] = d.result
            save_clients_cb(app.items)
            refresh_tv()

    def delete_log():
        LOG.info("Delete log clicked")
        i = selected_index()
        if i is None:
            messagebox.showinfo("Delete Log", "Select a log row to delete.")
            return
        if not messagebox.askyesno("Delete Log", "Delete the selected log entry?"):
            return
        del client["logs"][i]
        save_clients_cb(app.items)
        refresh_tv()

    def toggle_done():
        LOG.info("Toggle done clicked")
        sel = tv.selection()
        if not sel:
            messagebox.showinfo("Toggle", "Select a log row to toggle done/undone.")
            return
        _toggle_done_by_iid(sel[0])


    def _toggle_done_by_iid(iid: str):
        # Find the row index from iid, then flip the bool
        try:
            idx = tv.index(iid)
        except Exception:
            return
        client.setdefault("logs", [])
        if 0 <= idx < len(client["logs"]):
            client["logs"][idx]["done"] = not bool(client["logs"][idx].get("done"))
            save_clients_cb(app.items)
            refresh_tv()
            # keep focus/selection on the same row post-refresh
            try:
                new_iid = tv.get_children()[idx]
                tv.selection_set(new_iid)
                tv.focus(new_iid)
            except Exception:
                pass


    # Controls row
    controls = ttk.Frame(logs_tab)
    controls.grid(row=1, column=0, sticky="we", pady=(6, 0))
    controls.grid_columnconfigure(0, weight=1)

    # Quick-add entry
    v_quick = tk.StringVar()
    ttk.Entry(controls, textvariable=v_quick).grid(row=0, column=0, sticky="we")

    def quick_add():
        t = v_quick.get().strip()
        if not t:
            return
        client.setdefault("logs", []).append({
            "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "user": "",
            "text": t,
            "done": False
        })
        save_clients_cb(app.items)
        v_quick.set("")
        refresh_tv()

    ttk.Button(controls, text="Add",    command=add_log).grid(row=0, column=1, padx=(6,0))
    ttk.Button(controls, text="Edit",   command=edit_log).grid(row=0, column=2, padx=(6,0))
    ttk.Button(controls, text="Delete", command=delete_log).grid(row=0, column=3, padx=(6,0))
    ttk.Button(controls, text="Toggle Done", command=toggle_done).grid(row=0, column=4, padx=(6,0))
    ttk.Button(controls, text="Quick Add", command=quick_add).grid(row=0, column=5, padx=(6,0))

    # Double-click to edit, single-click on ✓ column to toggle
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

    # Clear when clicking the tab background or controls area
    logs_tab.bind("<Button-1>", lambda e: clear_selection() if e.widget is logs_tab else None)
    controls.bind("<Button-1>", lambda e: clear_selection())

    # Also clear on notebook/tab switches (optional but nice)
    def _on_tab_changed(_e=None):
        clear_selection()
    notebook.bind("<<NotebookTabChanged>>", _on_tab_changed)


    # Context menu
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

    refresh_tv()
