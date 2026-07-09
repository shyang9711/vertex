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
        today_str,
    )


class TrackerItemDialog(tk.Toplevel):
    """Modal dialog for add/edit of tracker items (file, reminder, issue)."""

    def __init__(self, master, title: str = "Tracker Item", kind: str = "file", initial: dict | None = None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.kind = (kind or "file").strip().lower()
        init = dict(initial or {})

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        row = 0
        self._vars: dict[str, tk.Variable] = {}

        def _add_label(text: str, r: int):
            ttk.Label(frm, text=text).grid(row=r, column=0, sticky="w", pady=(4, 0))

        def _add_entry(key: str, r: int, default: str = "", width: int = 32):
            var = tk.StringVar(value=str(init.get(key, default) or ""))
            self._vars[key] = var
            ttk.Entry(frm, textvariable=var, width=width).grid(row=r, column=1, sticky="we", padx=(8, 0), pady=(4, 0))
            return var

        def _add_combo(key: str, r: int, values: list[str], default: str = ""):
            var = tk.StringVar(value=str(init.get(key, default) or (values[0] if values else "")))
            self._vars[key] = var
            cb = ttk.Combobox(frm, textvariable=var, values=values, state="readonly", width=30)
            cb.grid(row=r, column=1, sticky="we", padx=(8, 0), pady=(4, 0))
            return var

        def _add_check(key: str, r: int, text: str, default: bool = False):
            var = tk.BooleanVar(value=bool(init.get(key, default)))
            self._vars[key] = var
            ttk.Checkbutton(frm, text=text, variable=var).grid(row=r, column=1, sticky="w", padx=(8, 0), pady=(4, 0))
            return var

        if self.kind == "file":
            _add_label("Tax year", row)
            _add_entry("tax_year", row, default=str(init.get("tax_year") or ""))
            row += 1

            _add_label("Category", row)
            _add_combo("category", row, CATEGORIES, default=init.get("category", "Income Tax"))
            row += 1

            _add_label("Name", row)
            _add_entry("name", row)
            row += 1

            _add_label("Status", row)
            _add_combo("status", row, FILE_REQUEST_STATUSES, default=init.get("status", "Needed"))
            row += 1

            _add_label("Requested date", row)
            _add_entry("requested_date", row)
            row += 1

            _add_label("Received date", row)
            _add_entry("received_date", row)
            row += 1

            _add_label("Due date", row)
            _add_entry("due_date", row)
            row += 1

            _add_label("Repeat next year", row)
            _add_check("repeat_next_year", row, "Repeat next year", default=bool(init.get("repeat_next_year", True)))
            row += 1

            _add_label("Source", row)
            _add_combo("source", row, FILE_SOURCES, default=init.get("source", "Client"))
            row += 1

            _add_label("Priority", row)
            _add_combo("priority", row, PRIORITIES, default=init.get("priority", "Normal"))
            row += 1

        elif self.kind == "reminder":
            _add_label("Title", row)
            _add_entry("title", row)
            row += 1

            _add_label("Category", row)
            _add_combo("category", row, CATEGORIES, default=init.get("category", "Other"))
            row += 1

            _add_label("Status", row)
            _add_combo("status", row, REMINDER_STATUSES, default=init.get("status", "Active"))
            row += 1

            _add_label("Priority", row)
            _add_combo("priority", row, PRIORITIES, default=init.get("priority", "Normal"))
            row += 1

            _add_label("Applies every year", row)
            _add_check("applies_every_year", row, "Applies every year", default=bool(init.get("applies_every_year", True)))
            row += 1

            _add_label("Tax year (optional)", row)
            _add_entry("tax_year", row)
            row += 1

        elif self.kind == "issue":
            _add_label("Type", row)
            _add_combo("type", row, ISSUE_TYPES, default=init.get("type", "Other"))
            row += 1

            _add_label("Category", row)
            _add_combo("category", row, CATEGORIES, default=init.get("category", "Other"))
            row += 1

            _add_label("Tax year", row)
            _add_entry("tax_year", row)
            row += 1

            _add_label("Title", row)
            _add_entry("title", row)
            row += 1

            _add_label("Status", row)
            _add_combo("status", row, ISSUE_STATUSES, default=init.get("status", "Open"))
            row += 1

            _add_label("Priority", row)
            _add_combo("priority", row, PRIORITIES, default=init.get("priority", "Normal"))
            row += 1

            _add_label("Opened date", row)
            _add_entry("opened_date", row, default=init.get("opened_date") or today_str())
            row += 1

            _add_label("Closed date", row)
            _add_entry("closed_date", row)
            row += 1

        # Note (all kinds)
        ttk.Label(frm, text="Note").grid(row=row, column=0, sticky="nw", pady=(8, 0))
        self._note = ScrolledText(frm, width=40, height=4, wrap="word")
        self._note.grid(row=row, column=1, sticky="we", padx=(8, 0), pady=(8, 0))
        if init.get("note"):
            self._note.insert("1.0", str(init.get("note")))
        row += 1

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Escape>", lambda _e: self.destroy())
        self.grab_set()
        self.transient(master)

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
                messagebox.showerror("Validation", "Name is required.", parent=self)
                return
        else:
            if not out.get("title"):
                messagebox.showerror("Validation", "Title is required.", parent=self)
                return

        if init_id := (getattr(self, "_initial_id", None)):
            out["id"] = init_id

        self.result = out
        self.destroy()

    @classmethod
    def open(cls, master, title: str, kind: str, initial: dict | None = None) -> dict | None:
        dlg = cls(master, title=title, kind=kind, initial=initial)
        if initial and initial.get("id"):
            dlg._initial_id = initial["id"]
        master.wait_window(dlg)
        if dlg.result and initial:
            dlg.result.setdefault("id", initial.get("id", ""))
            dlg.result.setdefault("created_ts", initial.get("created_ts", ""))
            dlg.result.setdefault("archived", initial.get("archived", False))
        return dlg.result
