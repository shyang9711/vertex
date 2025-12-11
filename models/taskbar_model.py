from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))
        
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from functions.utils.app_logging import get_logger

class TaskbarModel:
    """
    Menubar: [File] [Edit] [Company] [Action]
      - Action → Client Actions… (opens in-app Actions page)
    """

    def __init__(
        self,
        root: tk.Tk,
        *,
        on_exit=None,
        on_new_company=None,
        get_account_managers=None,
        set_account_managers=None,
        on_open_preferences=None,
        on_open_actions=None,
        on_open_batch_tasks=None,
        on_save_data=None,
        on_import_data=None,
        on_export_data=None,
        on_check_updates=None,
        on_about=None,
    ):
        self.root = root
        self.on_exit = on_exit or (lambda: root.destroy())
        self.on_new_company = on_new_company or (lambda: None)
        self.get_account_managers = get_account_managers or (lambda: [])
        self.set_account_managers = set_account_managers or (lambda _lst: None)
        self.on_open_preferences = on_open_preferences or (lambda: self._default_prefs_dialog())
        self.on_open_actions = on_open_actions or (lambda: messagebox.showinfo("Action", "Actions page not wired."))
        self.on_open_batch_tasks = on_open_batch_tasks or (
            lambda: messagebox.showinfo(
                "Batch Task Settings",
                "Batch Task Settings is not wired yet.",
            )
        )
        self.on_check_updates = on_check_updates or (lambda: None)
        self.on_about = on_about or (lambda: None)

        self.on_save_data = on_save_data or (lambda: None)
        self.on_import_data = on_import_data or (lambda: None)
        self.on_export_data = on_export_data or (lambda: None)
        
        self.log = get_logger("taskbar")
        self._menu = tk.Menu(root)
        root.config(menu=self._menu)
        self._build_menu()

        is_macos = (root.tk.call("tk", "windowingsystem") == "aqua")
        mod = "Command" if is_macos else "Control"
        root.bind_all(f"<{mod}-q>", lambda e: self._do_exit())
        root.bind_all(f"<{mod}-n>", lambda e: self._do_new_company())
        root.bind_all(f"<{mod}-comma>", lambda e: self._do_preferences())
        root.bind_all(f"<{mod}-s>", lambda e: self._do_save_data())

    def _build_menu(self):
        # File
        m_file = tk.Menu(self._menu, tearoff=False)
        m_file.add_command(
            label="Save Data",
            command=self._do_save_data,
            accelerator="Ctrl+S" if self._accel_ctrl() else "⌘S",
        )
        m_file.add_command(label="Import Data...", command=self._do_import_data)
        m_file.add_command(label="Export Data...", command=self._do_export_data)
        m_file.add_separator()
        m_file.add_command(
            label="Exit",
            command=self._do_exit,
            accelerator="Ctrl+Q" if self._accel_ctrl() else "⌘Q",
        )
        self._menu.add_cascade(label="File", menu=m_file)

        # Edit
        m_edit = tk.Menu(self._menu, tearoff=False)
        m_edit.add_command(label="Account Managers...", command=self._do_account_managers)
        m_edit.add_command(label="Batch Task Settings...", command=self._do_batch_task_settings)
        m_edit.add_separator()
        m_edit.add_command(label="Preferences...", command=self._do_preferences,
                           accelerator="Ctrl+," if self._accel_ctrl() else "⌘,")
        self._menu.add_cascade(label="Edit", menu=m_edit)

        # Company
        m_comp = tk.Menu(self._menu, tearoff=False)
        m_comp.add_command(label="New Company", command=self._do_new_company,
                           accelerator="Ctrl+N" if self._accel_ctrl() else "⌘N")
        self._menu.add_cascade(label="Company", menu=m_comp)

        # Action
        m_action = tk.Menu(self._menu, tearoff=False)
        m_action.add_command(label="Client Actions…", command=self.on_open_actions)
        self._menu.add_cascade(label="Action", menu=m_action)
        
        # Help
        m_help = tk.Menu(self._menu, tearoff=False)
        m_help.add_command(label="Check for Updates...", command=self._do_check_updates)
        m_help.add_separator()
        m_help.add_command(label="About Vertex", command=self._do_about)
        self._menu.add_cascade(label="Help", menu=m_help)

    def _accel_ctrl(self) -> bool:
        return (self.root.tk.call("tk", "windowingsystem") != "aqua")

    # --- handlers
    def _do_exit(self): self.on_exit()
    def _do_new_company(self): self.on_new_company()
    def _do_preferences(self):
        self.log.info("Preferences opened")
        self.on_open_preferences()
    def _do_account_managers(self):
        self.log.info("Account managers dialog opened")
        AccountManagersDialog(self.root, get_fn=self.get_account_managers, set_fn=self.set_account_managers)
    def _do_batch_task_settings(self):
        self.log.info("Batch Task Settings dialog opened")
        self.on_open_batch_tasks()
    # --- data handlers
    def _do_save_data(self):
        self.log.info("Save Data triggered from menu")
        self.on_save_data()

    def _do_import_data(self):
        self.log.info("Import Data triggered from menu")
        self.on_import_data()

    def _do_export_data(self):
        self.log.info("Export Data triggered from menu")
        self.on_export_data()
        
    def _do_check_updates(self):
        self.log.info("Check for updates triggered from menu")
        self.on_check_updates()

    def _do_about(self):
        self.log.info("About dialog triggered from menu")
        self.on_about()

class AccountManagersDialog:
    COLUMNS = ("name", "email", "phone")
    def __init__(self, master, *, get_fn, set_fn):
        self.master = master; self.get_fn = get_fn; self.set_fn = set_fn
        self._data = self._normalize(get_fn() or [])

        dlg = self.dlg = tk.Toplevel(master)
        dlg.title("Account Managers"); dlg.transient(master); dlg.grab_set(); dlg.minsize(520, 320)
        root = ttk.Frame(dlg, padding=10); root.pack(fill="both", expand=True)
        ttk.Label(root, text="Account Managers", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w")

        self.tree = ttk.Treeview(root, columns=self.COLUMNS, show="headings", height=10, selectmode="browse")
        for c in self.COLUMNS:
            self.tree.heading(c, text=c.capitalize()); self.tree.column(c, width=160, anchor="w")
        self.tree.grid(row=1, column=0, sticky="nsew", pady=(8,8))

        btns = ttk.Frame(root); btns.grid(row=2, column=0, sticky="e")
        ttk.Button(btns, text="Add", command=self._add).pack(side="left", padx=(0,6))
        ttk.Button(btns, text="Edit", command=self._edit).pack(side="left", padx=(0,6))
        ttk.Button(btns, text="Delete", command=self._delete).pack(side="left")

        root.columnconfigure(0, weight=1); root.rowconfigure(1, weight=1)
        self._refresh(); dlg.wait_window(dlg)

    def _normalize(self, data):
        out = []
        for x in data or []:
            if isinstance(x, dict):
                out.append({"name": x.get("name",""), "email": x.get("email",""), "phone": x.get("phone","")})
            else:
                out.append({"name": str(x), "email": "", "phone": ""})
        return out

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        for row in self._normalize(self._data):
            self.tree.insert("", "end", values=(row["name"], row["email"], row["phone"]))

    def _add(self):
        d = _EditRowDialog(self.master, title="Add Manager")
        if d.result:
            self._data.append(d.result); self.set_fn(self._data); self._refresh()

    def _edit(self):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values")
        d = _EditRowDialog(self.master, title="Edit Manager", initial={"name":vals[0], "email":vals[1], "phone":vals[2]})
        if d.result:
            idx = self.tree.index(sel[0]); self._data[idx] = d.result; self.set_fn(self._data); self._refresh()

    def _delete(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0]); del self._data[idx]; self.set_fn(self._data); self._refresh()

class _EditRowDialog:
    def __init__(self, master, title="Edit", initial=None):
        self.result = None
        dlg = self.dlg = tk.Toplevel(master); dlg.title(title); dlg.transient(master); dlg.grab_set()
        frm = ttk.Frame(dlg, padding=10); frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Name").grid(row=0, column=0, sticky="w")
        self.e_name = ttk.Entry(frm, width=42); self.e_name.grid(row=1, column=0, sticky="we", pady=(0,8))
        ttk.Label(frm, text="Email").grid(row=2, column=0, sticky="w")
        self.e_email = ttk.Entry(frm, width=42); self.e_email.grid(row=3, column=0, sticky="we", pady=(0,8))
        ttk.Label(frm, text="Phone").grid(row=4, column=0, sticky="w")
        self.e_phone = ttk.Entry(frm, width=42); self.e_phone.grid(row=5, column=0, sticky="we", pady=(0,8))

        btns = ttk.Frame(frm); btns.grid(row=6, column=0, sticky="e")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=(0,6))
        ttk.Button(btns, text="Save", command=self._save).pack(side="right")

        if initial:
            self.e_name.insert(0, initial.get("name","")); self.e_email.insert(0, initial.get("email","")); self.e_phone.insert(0, initial.get("phone",""))
        self.e_name.focus_set(); dlg.wait_window(dlg)

    def _save(self):
        self.result = {"name": self.e_name.get().strip(), "email": self.e_email.get().strip(), "phone": self.e_phone.get().strip()}
        self.dlg.destroy()
