from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[2]  # Scripts/
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))

import tkinter as tk
from tkinter import ttk

try:
    from vertex.utils.app_logging import get_logger
    
except ModuleNotFoundError:
    from utils.app_logging import get_logger
LOG = get_logger("reports")

# (Optional) use NewUI tag names if available
try:
    from styles.new_ui import NewUI
except Exception:
    NewUI = None

class ReportsPage:
    """
    Matches your other pages' pattern: page object with ensure(host) that returns a Frame.
    """
    def __init__(self, app=None, **_kwargs):
        self.app = app
        self.frame: ttk.Frame | None = None
        self.log = get_logger("reports")

    def ensure(self, host) -> ttk.Frame:
        self.log.info("ensure() reports host=%s", getattr(host, "winfo_name", lambda: "?")())
        if self.frame is None or not self.frame.winfo_exists():
            self.frame = ttk.Frame(host, padding=12, style="Card.TFrame")
            self._build(self.frame)
        # (re)attach
        try:
            self.frame.pack_forget()
        except Exception:
            pass
        self.frame.pack(fill="both", expand=True)
        return self.frame

    # --- UI ---
    def _build(self, root: ttk.Frame) -> None:
        # Heading
        hdr = ttk.Frame(root); hdr.pack(fill="x", pady=(0, 8))
        ttk.Label(hdr, text="Reports", style="Header.TLabel").pack(side="left")

        # Body
        card = ttk.Frame(root, padding=16, style="Card.TFrame")
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="🚧 Under Construction 🚧", font=("Segoe UI", 14)).pack(pady=(6, 6))
        ttk.Label(card, text="This section will display client, payroll, and tax reports once implemented.",
                  wraplength=720, justify="center").pack()

        # Back button (like your other navigations)
        btns = ttk.Frame(card); btns.pack(pady=(18, 0))
        ttk.Button(btns, text="← Back to Dashboard",
                   command=lambda: self.app.navigate("dashboard") if self.app else None
                  ).pack()

