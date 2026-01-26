"""
Log Dialog for adding and editing log entries.
"""

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import datetime


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
