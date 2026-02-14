"""
Note/Log Dialog for editing note entries (text only; timestamp locked).
Add: no timestamp shown; Edit: timestamp read-only, show "(Edited)" when entry was edited.
"""

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import datetime


class LogDialog(tk.Toplevel):
    def __init__(self, master, title="Note", initial=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        init = initial or {}
        self._initial = init
        frm = ttk.Frame(self, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        # Timestamp: read-only when editing; show "(Edited)" when entry was edited
        ts_val = init.get("ts", "")
        is_edit = bool(ts_val)
        if is_edit:
            if init.get("edited"):
                ts_display = f"{ts_val} (Edited)" if ts_val else "(Edited)"
            else:
                ts_display = ts_val or "—"
            ttk.Label(frm, text="Timestamp").grid(row=0, column=0, sticky="w", pady=(0,2))
            ttk.Label(frm, text=ts_display).grid(row=1, column=0, sticky="w", pady=(0,6))
            row_text = 2
        else:
            row_text = 0

        ttk.Label(frm, text="Text").grid(row=row_text, column=0, sticky="w", pady=(0,2))
        self.txt = ScrolledText(frm, width=56, height=8, wrap="word")
        self.txt.grid(row=row_text + 1, column=0, sticky="nsew", pady=(0,6))
        if "text" in init:
            self.txt.insert("1.0", init["text"])

        btns = ttk.Frame(frm)
        btns.grid(row=row_text + 2, column=0, sticky="e", pady=(8,0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda _e: self._save())
        self.bind("<Escape>", lambda _e: self.destroy())
        self.after(50, self.txt.focus_set)
        self.grab_set()
        self.transient(master)

    def _save(self):
        init = self._initial
        ts_val = init.get("ts", "")
        if ts_val:
            # Edit: keep original timestamp, set edited
            self.result = {
                "ts": ts_val,
                "text": self.txt.get("1.0", "end").strip(),
                "done": bool(init.get("done", False)),
                "edited": True,
            }
        else:
            # Add: set timestamp now, not edited
            self.result = {
                "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                "text": self.txt.get("1.0", "end").strip(),
                "done": False,
                "edited": False,
            }
        self.destroy()
