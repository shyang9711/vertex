from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from typing import Callable, Optional


class OptionalNoteDialog(tk.Toplevel):
    """
    Lightweight modal dialog for optional notes.

    Result semantics:
      - Cancel/Window close => result is None
      - OK => result is a (possibly empty) string (caller decides what empty means)
    """

    def __init__(
        self,
        master,
        *,
        title: str,
        prompt: str,
        initial: str = "",
    ):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result: Optional[str] = None

        frm = ttk.Frame(self, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)

        ttk.Label(frm, text=prompt, style="Subtle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.txt = ScrolledText(frm, width=56, height=6, wrap="word")
        self.txt.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        if initial:
            self.txt.insert("1.0", initial)

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, sticky="e")
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btns, text="OK", command=self._ok).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda _e: self._cancel())
        self.bind("<Return>", lambda _e: self._ok())

        self.transient(master)
        self.grab_set()
        self.after(30, self.txt.focus_set)

        # Center-ish (best-effort). Tkinter doesn't provide perfect screen geometry without platform specifics.
        try:
            self.update_idletasks()
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _cancel(self):
        self.result = None
        self.destroy()

    def _ok(self):
        # Empty string is allowed (optional note).
        self.result = self.txt.get("1.0", "end").strip()
        self.destroy()


class WorkSessionPopup(tk.Toplevel):
    """Always-on-top compact widget for "currently working on" work sessions."""

    def __init__(
        self,
        master,
        *,
        on_hold: Callable[[], None],
        on_finish: Callable[[], None],
    ):
        super().__init__(master)

        self._on_hold = on_hold
        self._on_finish = on_finish
        self._top_keepalive_job: Optional[str] = None

        self.title("Currently working on")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        try:
            self.transient(master)
        except Exception:
            pass

        # Prevent accidental close; the spec wants Hold/Finished to be the end of the workflow.
        self.protocol("WM_DELETE_WINDOW", self._on_close_attempt)

        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Client", style="Subtle.TLabel").grid(row=0, column=0, sticky="w")
        self.lbl_client = ttk.Label(root, text="", style="Header.TLabel", font=("Segoe UI", 10, "bold"))
        self.lbl_client.grid(row=1, column=0, sticky="w", pady=(0, 6))

        ttk.Label(root, text="Task", style="Subtle.TLabel").grid(row=2, column=0, sticky="w")
        self.lbl_task = ttk.Label(root, text="", style="Header.TLabel", wraplength=340)
        self.lbl_task.grid(row=3, column=0, sticky="w", pady=(0, 10))

        btns = ttk.Frame(root)
        btns.grid(row=4, column=0, sticky="e")
        self.btn_hold = ttk.Button(btns, text="Hold", command=self._hold_clicked, width=8)
        self.btn_finish = ttk.Button(btns, text="Finish", command=self._finish_clicked, width=8, style="Accent.TButton")
        self.btn_hold.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_finish.pack(side=tk.LEFT)

        self.after(50, self._keep_top_alive)

        # Reasonable compact size.
        try:
            self.geometry("380x170")
        except Exception:
            pass

    def _keep_top_alive(self):
        # Re-assert topmost periodically (Windows can steal focus from topmost windows).
        if not self.winfo_exists():
            return
        try:
            self.attributes("-topmost", True)
            self.lift()
        except Exception:
            pass
        self._top_keepalive_job = self.after(1000, self._keep_top_alive)

    def destroy(self):
        try:
            if self._top_keepalive_job:
                self.after_cancel(self._top_keepalive_job)
        except Exception:
            pass
        super().destroy()

    def _on_close_attempt(self):
        # Spec wants the session to be ended via Hold/Finished.
        try:
            messagebox.showinfo("Work Session", "Use Hold or Finished to end the current work session.")
        except Exception:
            pass

    def set_content(self, *, client_name: str, task_name: str):
        self.lbl_client.config(text=client_name or "—")
        self.lbl_task.config(text=task_name or "—")

    def _hold_clicked(self):
        self._on_hold()

    def _finish_clicked(self):
        self._on_finish()

