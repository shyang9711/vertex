from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    from vertex.utils.client_tracker import (
        FILE_REQUEST_STATUSES,
        REMINDER_STATUSES,
        ISSUE_STATUSES,
        CATEGORIES,
        ISSUE_TYPES,
        FILE_SOURCES,
        PRIORITIES,
        FILE_TEMPLATES,
        FILE_TEMPLATE_LABELS,
        default_tax_year,
        today_str,
    )
except ModuleNotFoundError:
    from utils.client_tracker import (
        FILE_REQUEST_STATUSES,
        REMINDER_STATUSES,
        ISSUE_STATUSES,
        CATEGORIES,
        ISSUE_TYPES,
        FILE_SOURCES,
        PRIORITIES,
        FILE_TEMPLATES,
        FILE_TEMPLATE_LABELS,
        default_tax_year,
        today_str,
    )

_MIN_SIZE = {
    "file": (560, 460),
    "reminder": (520, 380),
    "issue": (560, 460),
}


class TrackerItemDialog(tk.Toplevel):
    """Modal dialog for add/edit of tracker items (file, reminder, issue)."""

    _active: "TrackerItemDialog | None" = None

    def __init__(self, master, title: str = "Tracker Item", kind: str = "file", initial: dict | None = None):
        super().__init__(master)
        self.title(title)
        self.resizable(True, True)
        self.result = None
        self.kind = (kind or "file").strip().lower()
        init = dict(initial or {})
        is_edit = bool(init.get("id"))
        self._initial_id = init.get("id")

        self._frm = ttk.Frame(self, padding=10)
        self._frm.pack(fill="both", expand=True)
        frm = self._frm
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(3, weight=1)

        self._vars: dict[str, tk.Variable] = {}
        row = 0
        self._first_focus: tk.Widget | None = None

        def _lbl(text: str, r: int, c: int = 0):
            ttk.Label(frm, text=text).grid(row=r, column=c, sticky="w", padx=(0, 4), pady=2)

        def _entry(key: str, r: int, c: int = 1, default: str = "", colspan: int = 1):
            var = tk.StringVar(value=str(init.get(key, default) or ""))
            self._vars[key] = var
            w = ttk.Entry(frm, textvariable=var, width=22)
            w.grid(row=r, column=c, columnspan=colspan, sticky="we", padx=(0, 10), pady=2)
            if self._first_focus is None:
                self._first_focus = w
            return var

        def _combo(key: str, r: int, c: int, values: list[str], default: str = "", colspan: int = 1):
            var = tk.StringVar(value=str(init.get(key, default) or (values[0] if values else "")))
            self._vars[key] = var
            cb = ttk.Combobox(frm, textvariable=var, values=values, state="readonly", width=20)
            cb.grid(row=r, column=c, columnspan=colspan, sticky="we", padx=(0, 10), pady=2)
            if self._first_focus is None:
                self._first_focus = cb
            return var

        def _check(key: str, r: int, c: int, text: str, default: bool = False):
            var = tk.BooleanVar(value=bool(init.get(key, default)))
            self._vars[key] = var
            ttk.Checkbutton(frm, text=text, variable=var).grid(row=r, column=c, columnspan=2, sticky="w", pady=2)
            return var

        if self.kind == "file":
            v_template = tk.StringVar(value="— Custom —")
            _lbl("Template", row)
            cb_tpl = ttk.Combobox(frm, textvariable=v_template, values=FILE_TEMPLATE_LABELS, state="readonly", width=28)
            cb_tpl.grid(row=row, column=1, columnspan=3, sticky="we", pady=2)
            row += 1

            ty_default = init.get("tax_year") or (default_tax_year() if not is_edit else "")
            _lbl("Tax year", row)
            _entry("tax_year", row, default=str(ty_default))
            _lbl("Category", row, 2)
            _combo("category", row, 3, CATEGORIES, default=init.get("category", "Income Tax"))
            row += 1

            _lbl("File / Item", row)
            _entry("name", row)
            _lbl("Status", row, 2)
            _combo("status", row, 3, FILE_REQUEST_STATUSES, default=init.get("status", "Needed"))
            row += 1

            _lbl("Requested", row)
            _entry("requested_date", row)
            _lbl("Received", row, 2)
            _entry("received_date", row, 3)
            row += 1

            _lbl("Due date", row)
            _entry("due_date", row)
            _lbl("Source", row, 2)
            _combo("source", row, 3, FILE_SOURCES, default=init.get("source", "Client"))
            row += 1

            _lbl("Priority", row)
            _combo("priority", row, 1, PRIORITIES, default=init.get("priority", "Normal"))
            _check("repeat_next_year", row, 2, "Repeat next year", default=bool(init.get("repeat_next_year", True)))
            row += 1

            def _on_template(_e=None):
                label = v_template.get()
                for tpl_label, cat in FILE_TEMPLATES:
                    if tpl_label == label and tpl_label != "— Custom —":
                        if "name" in self._vars:
                            self._vars["name"].set(tpl_label)
                        if cat and "category" in self._vars:
                            self._vars["category"].set(cat)
                        break

            cb_tpl.bind("<<ComboboxSelected>>", _on_template)

        elif self.kind == "reminder":
            _lbl("Title", row)
            _entry("title", row, colspan=3)
            row += 1

            _lbl("Category", row)
            _combo("category", row, 1, CATEGORIES, default=init.get("category", "Other"))
            _lbl("Status", row, 2)
            _combo("status", row, 3, REMINDER_STATUSES, default=init.get("status", "Active"))
            row += 1

            _lbl("Priority", row)
            _combo("priority", row, 1, PRIORITIES, default=init.get("priority", "Normal"))
            _check("applies_every_year", row, 2, "Applies every year", default=bool(init.get("applies_every_year", True)))
            row += 1

            _lbl("Tax year", row)
            ty = init.get("tax_year") or (default_tax_year() if not is_edit else "")
            _entry("tax_year", row, default=str(ty), colspan=3)
            row += 1

        elif self.kind == "issue":
            _lbl("Type", row)
            _combo("type", row, 1, ISSUE_TYPES, default=init.get("type", "Other"))
            _lbl("Category", row, 2)
            _combo("category", row, 3, CATEGORIES, default=init.get("category", "Other"))
            row += 1

            ty_default = init.get("tax_year") or (default_tax_year() if not is_edit else "")
            _lbl("Tax year", row)
            _entry("tax_year", row, default=str(ty_default))
            _lbl("Status", row, 2)
            _combo("status", row, 3, ISSUE_STATUSES, default=init.get("status", "Open"))
            row += 1

            _lbl("Title", row)
            _entry("title", row, colspan=3)
            row += 1

            _lbl("Priority", row)
            _combo("priority", row, 1, PRIORITIES, default=init.get("priority", "Normal"))
            _lbl("Opened", row, 2)
            _entry("opened_date", row, 3, default=init.get("opened_date") or today_str())
            row += 1

            _lbl("Closed", row)
            _entry("closed_date", row, colspan=3)
            row += 1

        ttk.Label(frm, text="Note").grid(row=row, column=0, sticky="nw", pady=(6, 0))
        self._note = ScrolledText(frm, width=52, height=3, wrap="word")
        self._note.grid(row=row, column=1, columnspan=3, sticky="we", pady=(6, 0))
        if init.get("note"):
            self._note.insert("1.0", str(init.get("note")))
        row += 1

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=4, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda _e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self._fit_and_center(master)
        self._make_modal(master)

    def _fit_and_center(self, master) -> None:
        min_w, min_h = _MIN_SIZE.get(self.kind, (520, 400))
        self.update_idletasks()
        req_w = self._frm.winfo_reqwidth() + 24
        req_h = self._frm.winfo_reqheight() + 24
        w = max(req_w, min_w)
        h = max(req_h, min_h)
        self.minsize(min_w, min_h)
        try:
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            mw = max(master.winfo_width(), 1)
            mh = max(master.winfo_height(), 1)
            x = mx + max(0, (mw - w) // 2)
            y = my + max(0, (mh - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry(f"{w}x{h}")

    def _make_modal(self, master) -> None:
        self.transient(master)
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self.lift()
        self.focus_force()
        if self._first_focus is not None:
            self.after(50, lambda: self._first_focus.focus_set())

    def _release_modal(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        if TrackerItemDialog._active is self:
            TrackerItemDialog._active = None

    def _cancel(self):
        self.result = None
        self._release_modal()
        self.destroy()

    def _save(self):
        out: dict = {}
        for key, var in self._vars.items():
            if isinstance(var, tk.BooleanVar):
                out[key] = bool(var.get())
            else:
                out[key] = str(var.get() or "").strip()

        out["note"] = self._note.get("1.0", "end").strip()

        if self.kind == "file":
            if not out.get("name"):
                messagebox.showerror("Validation", "File / Item is required.", parent=self)
                return
        else:
            if not out.get("title"):
                messagebox.showerror("Validation", "Title is required.", parent=self)
                return

        if self._initial_id:
            out["id"] = self._initial_id

        self.result = out
        self._release_modal()
        self.destroy()

    @classmethod
    def open(cls, master, title: str, kind: str, initial: dict | None = None) -> dict | None:
        # Only one tracker dialog at a time — focus existing instead of stacking.
        if cls._active is not None:
            try:
                if cls._active.winfo_exists():
                    cls._active.lift()
                    cls._active.focus_force()
                    return None
            except tk.TclError:
                cls._active = None

        dlg = cls(master, title=title, kind=kind, initial=initial)
        cls._active = dlg
        master.wait_window(dlg)
        cls._active = None

        if dlg.result and initial:
            dlg.result.setdefault("id", initial.get("id", ""))
            dlg.result.setdefault("created_ts", initial.get("created_ts", ""))
            dlg.result.setdefault("archived", initial.get("archived", False))
        return dlg.result
