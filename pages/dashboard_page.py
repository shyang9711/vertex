from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))
        
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import datetime as _dt
import json, calendar as _cal, uuid
try:
    from functions.models.tasks_model import (
        TasksStore,
        adjust_if_weekend_or_holiday,
        DUE_DATE,
        calc_tags_for_occurrence,
        display_date_for
    )
    from functions.pages.checklist_page import ChecklistPage
    from functions.pages.reports_page import ReportsPage

    from functions.utils.app_logging import get_logger

except ModuleNotFoundError:
    from models.tasks_model import (
        TasksStore,
        adjust_if_weekend_or_holiday,
        DUE_DATE,
        calc_tags_for_occurrence,
        display_date_for
    )
    from pages.checklist_page import ChecklistPage
    from pages.reports_page import ReportsPage

    from utils.app_logging import get_logger
LOG = get_logger("dashboard")

# --- Theme ---
try:
    from ..styles.new_ui import NewUI
except Exception:
    NewUI = None

class AutoCompleteCombobox(ttk.Combobox):
    """Autocomplete that filters dropdown as you type without overwriting text."""
    def __init__(self, master=None, **kwargs):
        self._all_values = list(kwargs.get("values", []))
        super().__init__(master, **kwargs)
        self.bind("<KeyRelease>", self._on_keyrelease)
        self.bind("<FocusIn>", lambda e: self._refresh_values())

    def set_values(self, values):
        self._all_values = list(values)
        self["values"] = values

    def _refresh_values(self):
        typed = self.get()
        low = typed.lower()
        pool = self._all_values
        matches = [v for v in pool if low in v.lower()] if low else pool
        self["values"] = matches if matches else pool

    def _on_keyrelease(self, event):
        if event.keysym in (
            "Shift_L","Shift_R","Control_L","Control_R","Alt_L","Alt_R",
            "Left","Right","Up","Down","Home","End","Escape","Tab"
        ):
            return
        self._refresh_values()
        try:
            self.focus_set()
        except Exception:
            pass

def _today():
    return _dt.date.today()

# ---------- Weekend/Holiday helpers (unchanged) ----------
WEEKDAY_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
WD_NAME_TO_INT = {name: i for i, name in enumerate(WEEKDAY_NAMES)}

class DashboardPage:
    def __init__(self, app):
        self.app = app
        self.frame = None
        self.log = get_logger("dashboard")
        data_dir = None
        try:
            data_dir = getattr(app, "TASKS_DIR", None)
            data_dir = Path(data_dir) if data_dir else None
        except Exception:
            data_dir = None

        if not data_dir:
            # Fallback for dev mode
            pages_dir     = Path(__file__).resolve().parent
            functions_dir = pages_dir.parent
            data_dir      = functions_dir / "data" / "tasks"

        data_dir.mkdir(parents=True, exist_ok=True)

        self._data_dir = data_dir
        self.store = TasksStore(self._data_dir, app=self.app)
        self._row_tags = {}

        self._cal_year = None
        self._cal_month = None
        self._cal_wrap = None
        self._todo_rows = {}
        self._show_all_past = False
        self._CAL_MAX_DOTS = 4

        self._todo_ctx = None
        try:
            import tkinter.font as tkfont
            self._base_font = tkfont.nametofont("TkDefaultFont")
            self._strike_font = tkfont.Font(**self._base_font.configure())
            self._strike_font.configure(overstrike=1)
        except Exception:
            self._base_font = None
            self._strike_font = None

        self._grey_text = "#6B7280"  # same grey as 'done'

        self._CAL_CELL_MINHEIGHT = 60


    # -------- lifecycle --------
    def ensure(self, host):
        self.log.info("ensure() dashboard")
        if self.frame is None or not self.frame.winfo_exists():
            self.frame = ttk.Frame(host, padding=14, style="TFrame")
            self._build_ui(self.frame)
        return self.frame

    def show(self, host):
        self.log.info("show()")
        frm = self.ensure(host)
        frm.pack(fill=tk.BOTH, expand=True)
        self._refresh_todo_feed(); self._draw_calendar()
        try:
            frm.after(0, lambda: (self._refresh_todo_feed(), self._draw_calendar()))
        except Exception:
            # Fallback if after() isn't available for some reason
            self._refresh_todo_feed(); self._draw_calendar()

    # --- Import data ---
    def reload_from_disk(self):
        # Re-read tasks.json from disk and refresh visible widgets
        self.store = TasksStore(self._data_dir)
        try:
            self._refresh_todo_feed()
            self._draw_calendar()
        except Exception:
            pass

    # -------- build UI --------
    def _build_ui(self, root):
        # Apply modern theme if available
        if NewUI:
            try:
                NewUI.install(root.winfo_toplevel())
            except Exception:
                pass

        for i, w in enumerate((1, 2, 1)):
            root.columnconfigure(i, weight=w)
        root.rowconfigure(1, weight=1)

        # ---- Header bar
        header = ttk.Frame(root, style="TFrame")
        header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0,10))
        header.columnconfigure(0, weight=1)   # title stretches left
        header.columnconfigure(1, weight=0)   # nav buttons to the right

        # Title on the left
        title_lbl = ttk.Label(header, text="Dashboard", style="Header.TLabel")
        title_lbl.grid(row=0, column=0, sticky="w")

        # Navigation buttons container (right side)
        nav = ttk.Frame(header)
        nav.grid(row=0, column=1, sticky="e")

        ttk.Button(
            nav,
            text="Notes",
            command=lambda: self.app.navigate("notes", push=True),
            width=8,
            style="Accent.TButton"
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Button(
            nav,
            text="Company List",
            command=self._open_company_list,
            style="Accent.TButton"
        ).pack(side="left", padx=(0, 8))


        # Buttons to other pages
        ttk.Button(
            nav,
            text="Checklist",
            style="Accent.TButton",
            command=self._open_taxes_page
        ).pack(side="left", padx=(0, 8))

        # add more nav buttons here as needed
        ttk.Button(
            nav,
            text="Reports",
            command=self._open_reports, 
        ).pack(side="left", padx=(0, 8))

        # ---- Left card: Today
        left = ttk.Frame(root, padding=14, style="Card.TFrame")
        left.grid(row=1, column=0, sticky="nsew", padx=(0,10))
        today = _today()
        ttk.Label(left, text=today.strftime("%A, %B %d, %Y"), style="Header.TLabel").pack(anchor="w")
        total = len(getattr(self.app, "items", []))
        has_email = sum(1 for c in getattr(self.app, "items", []) if c.get("emails"))
        ttk.Label(left, text=f"Companies: {total}").pack(anchor="w", pady=(10,0))
        ttk.Label(left, text=f"With Emails: {has_email}", style="Subtle.TLabel").pack(anchor="w")

        # ---- Middle card: To‑Do
        mid = ttk.Frame(root, padding=14, style="Card.TFrame")
        mid.grid(row=1, column=1, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(1, weight=1)
        # Right-side toolbar
        top_row = ttk.Frame(mid)
        top_row.grid(row=0, column=0, sticky="ew", pady=(0,4))

        # 3-column layout: left/right flex, center fixed
        top_row.columnconfigure(0, weight=1)   # spacer (left)
        top_row.columnconfigure(1, weight=0)   # center (buttons)
        top_row.columnconfigure(2, weight=1)   # spacer (right)
                
        btns = ttk.Frame(top_row)
        btns.grid(row=0, column=1)  # center column
        ttk.Button(btns, text="Add Task",   style="Accent.TButton",  command=self._add_task).pack(side=tk.LEFT)
        ttk.Button(btns, text="Edit",       style="Outline.TButton", command=self._edit_task).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(btns, text="Delete",                              command=self._delete_task).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(btns, text="Stop Recurrence",                     command=self._stop_recurring).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(btns, text="Pause Recurrence",           command=self._pause_recurring).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(btns, text="Resume Recurrence",          command=self._resume_recurring).pack(side=tk.LEFT, padx=(8,0))


        # Show past toggle aligned to the right
        self._show_past_var = tk.BooleanVar(value=self._show_all_past)
        def _flip_past():
            self._show_all_past = bool(self._show_past_var.get())
            self._refresh_todo_feed()
        ttk.Checkbutton(top_row, text="Show past", variable=self._show_past_var, command=_flip_past).grid(row=1, column=1, sticky="e")

        cols = ("mark", "kind", "company", "date")
        self.todo_tv = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse", height=14, style="Modern.Treeview")
        self.todo_tv.column("mark", width=16, minwidth=14, stretch=False, anchor="center")
        self.todo_tv.column("kind", width=48, minwidth=20, stretch=True, anchor="center")
        self.todo_tv.column("company", width=48, minwidth=20, stretch=True, anchor="center")
        self.todo_tv.column("date", width=48, minwidth=20, stretch=True, anchor="center")
        self.todo_tv.heading("mark", text="", anchor="center")
        self.todo_tv.heading("kind",    text="Type",    command=lambda: self._on_sort_click("kind"))
        self.todo_tv.heading("company", text="Company", command=lambda: self._on_sort_click("company"))
        self.todo_tv.heading("date",    text="Date",    command=lambda: self._on_sort_click("date"))
        self.todo_tv.grid(row=1, column=0, sticky="nsew")
        ysb = ttk.Scrollbar(mid, orient="vertical", command=self.todo_tv.yview)
        self.todo_tv.configure(yscrollcommand=ysb.set)
        ysb.grid(row=1, column=1, sticky="ns")

        # interactions
        self.todo_tv.bind("<Button-1>", self._todo_click)
        self.todo_tv.bind("<Double-1>", self._todo_open_company)
        self.todo_tv.bind("<Button-3>", self._todo_show_context_menu)

        # ---- Right card: Calendar
        right = ttk.Frame(root, padding=14, style="Card.TFrame")
        right.grid(row=1, column=2, sticky="nsew", padx=(10,0))
        self._cal_year, self._cal_month = today.year, today.month

        ctrl = ttk.Frame(right)
        ctrl.pack(anchor="n")
        ttk.Button(ctrl, text="◀", width=3, command=lambda: self._shift_month(-1)).pack(side=tk.LEFT)
        self._cal_label_btn = ttk.Button(ctrl, text="", width=20, command=self._choose_month_year)
        self._cal_label_btn.pack(side=tk.LEFT, padx=6)
        ttk.Button(ctrl, text="▶", width=3, command=lambda: self._shift_month(+1)).pack(side=tk.LEFT)

        self._cal_wrap = ttk.Frame(right)
        self._cal_wrap.pack(pady=(8,0))

        
        legend = ttk.Frame(right)
        legend.pack(anchor="n", pady=(6, 0))
        def _legend_badge(parent, color, text):
            cv = tk.Canvas(parent, width=16, height=16, highlightthickness=0, bd=0)
            cv.create_oval(2, 2, 14, 14, fill=color, outline="")
            cv.pack(side="left", padx=(0,4))
            ttk.Label(parent, text=text).pack(side="left", padx=(0,12))

        _legend_badge(legend, "#22C55E", "to-do")
        _legend_badge(legend, "#9CA3AF", "done")


        try:
            root.after(0, lambda: (self._refresh_todo_feed(), self._draw_calendar()))
        except Exception:
            self._refresh_todo_feed(); self._draw_calendar()
        
        self._sort_col = "date"
        self._sort_asc = False
        self._update_sort_headers()

    # -------------- To‑Do feed --------------
    def _refresh_todo_feed(self):
        tv = getattr(self, "todo_tv", None)
        if not tv or not tv.winfo_exists():
            return

        import datetime as _dt
        tv.delete(*tv.get_children())
        self._todo_rows = {}
        
        tv.tag_configure("due", background="#FEF3C7")
        tv.tag_configure("submission", background="#FFEDD5")

        today = _dt.date.today()
        past_days = 999 if getattr(self, "_show_all_past", False) else 7
        future_days = 7
        window_start = today - _dt.timedelta(days=past_days)
        window_end   = today + _dt.timedelta(days=future_days)

        rows = []
        def occurs_between(t, start_d, end_d):
            comp = set(t.get("completed", []) or [])
            # was: 3 — too small for (lead 2 + weekend 2)
            SHIFT_BUFFER = 10
            scan_start = start_d - _dt.timedelta(days=SHIFT_BUFFER)
            scan_end   = end_d   + _dt.timedelta(days=SHIFT_BUFFER)

            d = scan_start
            while d <= scan_end:
                if self.store.occurs_on(t, d):
                    orig = d
                    disp = display_date_for(t, orig)
                    # mark done if either actual OR display date was stored
                    is_done = (orig.isoformat() in comp) or (disp.isoformat() in comp)
                    yield orig, disp, is_done
                d += _dt.timedelta(days=1)


        rows = []
        seen = set()

        # ---- PAST (up to yesterday)
        for i, t in enumerate(self.store.tasks):
            if not t.get("is_enabled", True):
                continue
            kind = t.get("kind", "")
            company = t.get("company_name") or self._company_name(t.get("company_idx")) or ""
            for orig, disp, is_done in occurs_between(t, window_start, today - _dt.timedelta(days=1)):
                if window_start <= disp <= window_end:
                    key = (i, orig.isoformat())
                    if key not in seen:
                        rows.append((disp, is_done, i, kind, company, orig))
                        seen.add(key)

        # ---- FUTURE (today .. window_end)
        for i, t in enumerate(self.store.tasks):
            if not t.get("is_enabled", True):
                continue
            kind = t.get("kind", "")
            company = t.get("company_name") or self._company_name(t.get("company_idx")) or ""
            futures = []
            for orig, disp, is_done in occurs_between(t, today, window_end):
                if today <= disp <= window_end:
                    futures.append((disp, is_done, i, kind, company, orig))
            if not futures:
                continue
            futures.sort(key=lambda x: x[0])
            first = futures[0]
            
            for item in (first, next((f for f in futures[1:] if f[1] is False), None)):
                if item is None:
                    continue
                disp, is_done, i_task, kind, company, orig = item
                key = (i_task, orig.isoformat())
                if key not in seen:
                    rows.append(item)
                    seen.add(key)
        
        def _row_sort_key_date(item):
            disp, is_done, i_task, kind, company, _orig = item
            task = self.store.tasks[i_task]
            method = (task.get("method") or "").lower()
            lead_flag = int(task.get("action_lead_days", 0) or 0) > 0
        
            done_rank = 1 if is_done else 0
        
            submission_first = 0 if (not is_done and (lead_flag or method in ("mail","direct_deposit"))) else 1
        
            date_ord = disp.toordinal() * (-1 if not getattr(self, "_sort_asc", True) else 1)
        
            # Then A→Z for ties
            return (
                date_ord,
                done_rank,
                submission_first,
                (company or "").casefold(),
                (kind or "").casefold(),
            )

        
        key  = getattr(self, "_sort_col", "date")
        asc  = getattr(self, "_sort_asc", False)
        
        if key == "date":
            rows.sort(key=_row_sort_key_date)
        elif key == "kind":
            rows.sort(key=lambda r: ((r[3] or "").casefold(),
                                     (r[4] or "").casefold(),
                                     r[0].toordinal()))
            if not asc: rows.reverse()
        elif key == "company":
            rows.sort(key=lambda r: ((r[4] or "").casefold(),
                                     (r[3] or "").casefold(),
                                     r[0].toordinal()))
            if not asc: rows.reverse()


        for disp, is_done, i_task, kind, company, orig in rows:
            task = self.store.tasks[i_task]
            canc = set(task.get("cancelled", []) or [])
            is_cancelled = (orig.isoformat() in canc) or (disp.isoformat() in canc)

            # Mark symbol: to-do "ㅁ", done "v", cancelled "×"
            if is_cancelled:
                mark = "×"
            else:
                mark = "v" if is_done else "ㅁ"

            iid = tv.insert("", "end", values=(mark, kind, company, disp.isoformat()))
            self._todo_rows[iid] = (i_task, orig)

            # base semantic tags
            tags = list(calc_tags_for_occurrence(task, disp, (is_done or is_cancelled), today))
            # keep submission highlight rules for pending submissions
            if (not is_done) and (not is_cancelled) and int(task.get("action_lead_days", 0) or 0) > 0:
                if "submission" not in tags:
                    tags.append("submission")
            # Show cancelled visually as strikethrough + gray (same as done)
            if is_cancelled and "cancelled" not in tags:
                # force gray base and add cancelled for strikethrough
                tags = [t for t in tags if t not in ("todo",)]  # drop 'todo' if present
                if "done" not in tags:
                    tags.insert(0, "done")
                tags.append("cancelled")

            tv.item(iid, tags=tuple(tags))
            self._row_tags[iid] = tuple(tags)

        tv.tag_configure("done", foreground="#6B7280")
        tv.tag_configure("cancelled", foreground="#6B7280")
        tv.tag_configure("todo", foreground="")
        tv.tag_configure("due", background="#FEF3C7") 
        tv.tag_configure("submission", background="#FFEDD5")

        # zebra striping if theme helper exists
        if NewUI:
            try:
                NewUI.stripe_tree(tv)
            except Exception:
                pass
            self._stripe_and_merge(tv)

            
        try:
            import tkinter.font as tkfont
            _base = tkfont.nametofont("TkDefaultFont")
            self._cancel_font = tkfont.Font(**_base.configure())
            self._cancel_font.configure(overstrike=1)
        except Exception:
            self._cancel_font = None

        # Apply font to cancelled rows (later at paint time)
        if self._cancel_font:
            tv.tag_configure("cancelled", font=self._cancel_font)
            
        for iid in tv.get_children(""):
            tags = list(tv.item(iid, "tags") or [])
            if "due" in tags and tags[-1] != "due":
                tags = [t for t in tags if t != "due"] + ["due"]
                tv.item(iid, tags=tuple(tags))
        
        self._update_sort_headers()

    def _todo_click(self, event):
        tv = self.todo_tv
        row = tv.identify_row(event.y)
        col = tv.identify_column(event.x)
        if not row:
            tv.selection_remove(tv.selection())
            return "break"
        if col == "#1":
            i_task, d = self._todo_rows[row]
            self._toggle_done_for_date(self.store.tasks[i_task], d)
            return "break"
        else:
            tv.selection_set(row)
            tv.focus(row)
            return

    def _todo_open_company(self, _e=None):
        tv = self.todo_tv
        sel = tv.selection()
        if not sel: return
        i_task, _d = self._todo_rows[sel[0]]
        self._navigate_company(self.store.tasks[i_task])

    def _toggle_done_for_date(self, task, date_obj: _dt.date):
        self.store.toggle_done_for_date(task, date_obj)
        self._refresh_todo_feed()
        self._draw_calendar()

    def _on_sort_click(self, col: str):
        # Toggle if clicking same column; otherwise switch to new col ascending
        if col == getattr(self, "_sort_col", "date"):
            self._sort_asc = not getattr(self, "_sort_asc", True)
        else:
            self._sort_col = col
            self._sort_asc = True
        self._update_sort_headers()
        self._refresh_todo_feed()

    def _update_sort_headers(self):
        tv = getattr(self, "todo_tv", None)
        if not tv or not tv.winfo_exists():
            return
        arrow = "▲" if getattr(self, "_sort_asc", True) else "▼"
        labels = {"mark":"", "kind":"Type", "company":"Company", "date":"Date"}
        for c in ("mark","kind","company","date"):
            text = labels[c]
            if c == getattr(self, "_sort_col", "date") and text:
                text = f"{text} {arrow}"
            # keep commands wired for the sortable ones
            if c == "kind":
                tv.heading(c, text=text, command=lambda cc="kind": self._on_sort_click(cc))
            elif c == "company":
                tv.heading(c, text=text, command=lambda cc="company": self._on_sort_click(cc))
            elif c == "date":
                tv.heading(c, text=text, command=lambda cc="date": self._on_sort_click(cc))
            else:
                tv.heading(c, text=text)


    def _navigate_company(self, task):
        items = getattr(self.app, "items", [])
        idx = task.get("company_idx")
        if idx is None:
            name = task.get("company_name")
            if name:
                try:
                    idx = next(i for i, c in enumerate(items) if c.get("name") == name)
                except StopIteration:
                    idx = None
        if idx is not None and 0 <= idx < len(items):
            try:
                self.app.navigate("detail", idx, push=True)
            except Exception as e:
                LOG.exception("Failed navigating to company for task %s: %s", task.get("title"), e)
                self.app.navigate("detail", idx)
        else:
            messagebox.showinfo("Open Company", "No linked company found for this task.")

            # still in class DashboardPage, add this helper anywhere among the other methods

    def _open_company_list(self):
        # Clear any prior query so the full list shows (harmless if q doesn’t exist)
        try:
            if hasattr(self.app, "q"):
                self.app.q.set("")
        except Exception:
            pass
        # Navigate to the list/table page
        self.app.navigate("search", None, push=True)

    def _open_taxes_page(self):
        self.app.navigate("taxes", None, push=True)

    def _open_reports(self):
        try:
            self.app.navigate("reports", None, push=True)
        except Exception as e:
            LOG.exception("Failed to open Reports page: %s", e)
            
    def open_batch_task_settings(self):
        """
        UI dialog to:
          1) Resume all paused recurring tasks for companies of the selected Account Manager.
          2) Pause all active recurring tasks for companies of the selected Account Manager.
        Stopped tasks (with end_on set) are NOT touched.
        """
        items = getattr(self.app, "items", []) or []

        # --- get account managers from app, if wired ---
        mgr_getter = getattr(self.app, "get_account_managers", None)
        if callable(mgr_getter):
            raw_mgrs = mgr_getter() or []
        else:
            raw_mgrs = []

        mgr_names: list[str] = []
        for x in raw_mgrs:
            if isinstance(x, dict):
                name = (x.get("name") or "").strip()
            else:
                name = str(x).strip()
            if name:
                mgr_names.append(name)

        if not mgr_names:
            messagebox.showinfo(
                "Batch Task Settings",
                "No account managers found. Use Edit → Account Managers… first.",
            )
            return

        d = tk.Toplevel(self.app)
        d.title("Batch Task Settings")
        d.resizable(False, False)

        frm = ttk.Frame(d, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Account Manager", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )

        v_mgr = tk.StringVar()
        cb_mgr = AutoCompleteCombobox(
            frm,
            textvariable=v_mgr,
            values=mgr_names,
            state="readonly",
            width=32,
        )
        cb_mgr.grid(row=1, column=0, columnspan=2, sticky="we", pady=(0, 8))

        ttk.Label(
            frm,
            text=(
                "These actions only affect recurring tasks:\n"
                "• One-off tasks are ignored\n"
                "• Stopped tasks (with an end date) are NOT changed"
            ),
            wraplength=360,
            style="Subtle.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))

        def _selected_manager_name() -> str:
            name = (v_mgr.get() or "").strip()
            if not name:
                messagebox.showwarning(
                    "Batch Task Settings", "Please choose an Account Manager first."
                )
            return name

        def _companies_for_manager(name: str):
            idxs: list[int] = []
            names: list[str] = []

            target = (name or "").strip().casefold()

            for i, c in enumerate(items):
                # Include your real key: acct_mgr
                mgr_val = (
                    c.get("account_manager")
                    or c.get("account_manager_name")
                    or c.get("acct_mgr")      # <--- IMPORTANT
                    or ""
                )
                cur = (mgr_val or "").strip().casefold()
                if cur == target and target:
                    idxs.append(i)
                    cname = (c.get("name") or "").strip()
                    if cname:
                        names.append(cname)

            return idxs, names

        def _do_resume():
            name = _selected_manager_name()
            if not name:
                return
            idxs, names = _companies_for_manager(name)
            if not idxs and not names:
                messagebox.showinfo(
                    "Batch Task Settings",
                    f"No companies found for account manager '{name}'.",
                )
                return

            if not messagebox.askyesno(
                "Confirm",
                (
                    f"Resume all paused recurring tasks for companies\n"
                    f"assigned to '{name}'?\n\n"
                    "Stopped tasks will not be changed."
                ),
            ):
                return

            changed = self.store.batch_resume_for_companies(idxs, names)
            messagebox.showinfo(
                "Batch Task Settings",
                f"Resumed {changed} task(s) for '{name}'."
                if changed
                else f"No paused recurring tasks to resume for '{name}'.",
            )
            self._refresh_todo_feed()
            self._draw_calendar()

        def _do_pause():
            name = _selected_manager_name()
            if not name:
                return
            idxs, names = _companies_for_manager(name)
            if not idxs and not names:
                messagebox.showinfo(
                    "Batch Task Settings",
                    f"No companies found for account manager '{name}'.",
                )
                return

            if not messagebox.askyesno(
                "Confirm",
                (
                    f"Pause all active recurring tasks for companies\n"
                    f"assigned to '{name}'?\n\n"
                    "Stopped tasks will not be changed."
                ),
            ):
                return

            changed = self.store.batch_pause_for_companies(idxs, names)
            messagebox.showinfo(
                "Batch Task Settings",
                f"Paused {changed} task(s) for '{name}'."
                if changed
                else f"No active recurring tasks to pause for '{name}'.",
            )
            self._refresh_todo_feed()
            self._draw_calendar()

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e")

        ttk.Button(
            btns,
            text="Resume paused tasks",
            style="Accent.TButton",
            command=_do_resume,
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            btns,
            text="Pause active tasks",
            command=_do_pause,
        ).pack(side="left", padx=(0, 8))

        ttk.Button(btns, text="Close", command=d.destroy).pack(side="left")

        d.grab_set()
        self.app.wait_window(d)


    # -------------- Calendar --------------
    def _draw_calendar(self):
        self.log.debug("Draw calendar %04d-%02d", self._cal_year or 0, self._cal_month or 0)
        if not self._cal_wrap: return
        for w in self._cal_wrap.winfo_children(): w.destroy()
        self._cal_label_btn.config(text=f"{_cal.month_name[self._cal_month]} {self._cal_year}")

        # Header
        header = ttk.Frame(self._cal_wrap)
        header.grid(row=0, column=0, columnspan=7, sticky="ew")
        for i, wd in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            ttk.Label(self._cal_wrap, text=wd, width=6, anchor="center").grid(row=1, column=i, padx=2, pady=(0,4))

        # Build dot map
        from collections import defaultdict
        dot_map = defaultdict(list)
        y, m = self._cal_year, self._cal_month
        month_first = _dt.date(y, m, 1)
        last_day = _cal.monthrange(y, m)[1]
        month_last = _dt.date(y, m, last_day)
        scan_end = month_last + _dt.timedelta(days=3)

        for t in self.store.tasks:
            comp = set(t.get("completed", []) or [])
            canc = set(t.get("cancelled", []) or [])
            d = month_first
            while d <= scan_end:
                if self.store.occurs_on(t, d):
                    disp = display_date_for(t, d)
                    if disp.year == y and disp.month == m:
                        d_iso   = d.isoformat()
                        disp_iso = disp.isoformat()
                        is_done = (d_iso in comp) or (disp_iso in comp)
                        is_cancelled = (d_iso in canc) or (disp_iso in canc) 
                        
                        if is_cancelled:
                            state = "cancelled"
                        elif is_done:
                            state = "done"
                        else:
                            state = "todo"

                        dot_map[disp.day].append((state, t, d)) 
                d += _dt.timedelta(days=1)

        cal = _cal.Calendar(firstweekday=0)
        for r, week in enumerate(cal.monthdayscalendar(y, m), start=2):
            for c, day in enumerate(week):
                cell = ttk.Frame(self._cal_wrap, padding=6, style="Card.TFrame")
                cell.grid(row=r, column=c, padx=2, pady=2, sticky="nsew")
                self._cal_wrap.grid_columnconfigure(c, weight=1, minsize=70)
                self._cal_wrap.grid_rowconfigure(r, weight=1, uniform="calrow", minsize=self._CAL_CELL_MINHEIGHT)


                if day == 0:
                    continue

                ddate = _dt.date(y, m, day)
                ttk.Label(cell, text=str(day)).pack(anchor="w")

                items = dot_map.get(day, [])
                if items:
                    pending_count = sum(1 for st, *_ in items if st == "todo")
                    gray_count    = sum(1 for st, *_ in items if st in ("done", "cancelled"))

                    # Row to hold our number-in-dot badges
                    badges_row = tk.Frame(cell)
                    badges_row.pack(anchor="w", pady=(2, 0))

                    def _make_badge(parent, count, fill, fg="#FFFFFF"):
                        cv = tk.Canvas(parent, width=20, height=20, highlightthickness=0, bd=0)
                        cv.create_oval(2, 2, 18, 18, fill=fill, outline="")
                        cv.create_text(10, 10, text=str(count), fill=fg, font=("TkDefaultFont", 9, "bold"))
                        cv.pack(side="left", padx=3)
                        return cv

                    # Green = to-do, Gray = done
                    if pending_count > 0:
                        _make_badge(badges_row, pending_count, "#22C55E")   # green
                    if gray_count > 0:
                        _make_badge(badges_row, gray_count, "#9CA3AF")      # gray

                    # Open the day dialog when the cell or badges are clicked
                    def _open_day(_e=None, date=ddate, pairs=items):
                        self._open_day_dialog(date, [(t, orig) for (_state, t, orig) in pairs])

                    cell.bind("<Button-1>", _open_day)
                    for w in badges_row.winfo_children():
                        w.bind("<Button-1>", _open_day)


    def _open_day_dialog(self, display_date: _dt.date, items):
        d = tk.Toplevel(self.app); d.title(display_date.isoformat()); d.resizable(False, False)
        frm = ttk.Frame(d, padding=12)
        frm.pack(fill="both", expand=True)
    
        if not items:
            ttk.Label(frm, text="No tasks on this date.").pack()
            ttk.Button(frm, text="Close", command=d.destroy).pack(pady=(8,0), anchor="e")
            d.grab_set(); self.app.wait_window(d); return
    
        # ✅ sort: submission first, then A→Z (company → type → title)
        today = _dt.date.today()
    
        def _company_of(t):
            return t.get("company_name") or self._company_name(t.get("company_idx")) or ""
    
        def _is_done(t, orig_date):
            return (orig_date.isoformat() in (t.get("completed", []) or []))
    
        def _day_sort_key(pair):
            t, orig_date = pair
            company_name = _company_of(t)
            is_done_now  = _is_done(t, orig_date)

            tags = calc_tags_for_occurrence(t, display_date, is_done_now, today)
            if not is_done_now and int(t.get("action_lead_days", 0) or 0) > 0:
                if "submission" not in tags:
                    tags = list(tags) + ["submission"]
            submission_first = 0 if (not is_done_now and "submission" in tags) else 1

            # Use negative ordinal so newer/future dates come first
            neg_date = -orig_date.toordinal()

            return (
                0 if not is_done_now else 1,  # pending before done
                submission_first,             # among pending: submission first
                neg_date,                     # NEW: future → past within group
                company_name.casefold(),
                (t.get("kind") or "").casefold(),
                (t.get("title") or "").casefold(),
            )
    
        items = sorted(items, key=_day_sort_key)
    
        for (t, orig_date) in items:
            company = t.get("company_name") or self._company_name(t.get("company_idx")) or ""
            row = ttk.Frame(frm); row.pack(fill="x", pady=2)

            comp = set(t.get("completed", []) or [])
            canc = set(t.get("cancelled", []) or [])
            disp_s = display_date_for(t, orig_date).isoformat()
            is_done = (orig_date.isoformat() in comp) or (disp_s in comp)
            is_cancelled = (orig_date.isoformat() in canc) or (disp_s in canc)

            v_done = tk.BooleanVar(value=is_done and not is_cancelled)
            v_cancel = tk.BooleanVar(value=is_cancelled)

            # Label with strikethrough when cancelled
            lbl = ttk.Label(row, text=f"{t.get('title','')} ({t.get('kind','')}) — {company}  [actual: {orig_date.isoformat()}]")

            def _refresh_label():
                strike = getattr(self, "_strike_font", None)
                base   = getattr(self, "_base_font", None)
                grey   = getattr(self, "_grey_text", "#6B7280")
                if strike and base:
                    if v_cancel.get():
                        lbl.configure(font=strike, foreground=grey)
                    else:
                        lbl.configure(font=base, foreground="")
                else:
                    lbl.configure(foreground=(grey if v_cancel.get() else ""))

            def _flip_done(task=t, date_=orig_date, var=v_done, other=v_cancel):
                # If marking done, clear cancellation
                if var.get():
                    if other.get():
                        other.set(False)
                        self.store.toggle_cancel_for_date(task, date_)
                    # mark done
                    self.store.toggle_done_for_date(task, date_)
                else:
                    # unmark done
                    self.store.toggle_done_for_date(task, date_)
                _refresh_label()
                self._refresh_todo_feed(); self._draw_calendar()

            def _flip_cancel(task=t, date_=orig_date, var=v_cancel, other=v_done):
                # If cancelling, clear done
                if var.get():
                    if other.get():
                        other.set(False)
                        self.store.toggle_done_for_date(task, date_)
                    self.store.toggle_cancel_for_date(task, date_)
                else:
                    self.store.toggle_cancel_for_date(task, date_)
                _refresh_label()
                self._refresh_todo_feed(); self._draw_calendar()

            ttk.Checkbutton(row, text="Done", variable=v_done, command=_flip_done).pack(side=tk.LEFT, padx=(0,6))
            ttk.Checkbutton(row, text="Cancelled", variable=v_cancel, command=_flip_cancel).pack(side=tk.LEFT, padx=(0,10))
            lbl.pack(side=tk.LEFT)

            # initial paint + live updates
            _refresh_label()
            v_cancel.trace_add("write", lambda *_: _refresh_label())
            v_done.trace_add("write",   lambda *_: _refresh_label())
    
        ttk.Separator(frm).pack(fill="x", pady=8)
        ttk.Button(frm, text="Close", command=d.destroy).pack(anchor="e")
        d.grab_set(); self.app.wait_window(d)


    def _choose_month_year(self):
        d = tk.Toplevel(self.app)
        d.title("Pick Month / Year")
        d.resizable(False, False)
        frm = ttk.Frame(d, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Month").grid(row=0, column=0, sticky="w")
        months = list(_cal.month_name)[1:]
        v_month = tk.StringVar(value=_cal.month_name[self._cal_month])
        cb_month = ttk.Combobox(frm, textvariable=v_month, values=months, state="readonly", width=14)
        cb_month.grid(row=1, column=0, sticky="w", pady=(0,6))

        ttk.Label(frm, text="Year").grid(row=0, column=1, sticky="w", padx=(12,0))
        v_year = tk.IntVar(value=self._cal_year)
        sp_year = tk.Spinbox(frm, from_=1980, to=2100, textvariable=v_year, width=8)
        sp_year.grid(row=1, column=1, sticky="w", padx=(12,0), pady=(0,6))

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="Cancel", command=d.destroy).pack(side=tk.RIGHT, padx=(6,0))
        def _apply():
            try:
                month_idx = months.index(v_month.get()) + 1
                year_val = int(v_year.get())
            except Exception:
                d.destroy(); return
            self._cal_year, self._cal_month = year_val, month_idx
            self._draw_calendar(); d.destroy()
        ttk.Button(btns, text="Go", style="Accent.TButton", command=_apply).pack(side=tk.RIGHT)
        d.grab_set(); self.app.wait_window(d)

    # -------------- CRUD --------------
    def _add_task(self):
        r = self._task_dialog(title="Add Task", init={
            "recurrence": {"freq": "one-off"},
            "notify_days": 4,
            "method": "none",
            "action_lead_days": 0
        })
        if not r: return
        r.setdefault("id", str(uuid.uuid4()))
        self.store.tasks.append(r)
        self.store.save()
        self._refresh_todo_feed(); self._draw_calendar()

    def _delete_task(self):
        i_task = self._selected_task_index()
        if i_task is None: return
        if not messagebox.askyesno("Confirm", "Delete this task?"): return
        del self.store.tasks[i_task]
        self.store.save()
        self._refresh_todo_feed(); self._draw_calendar()


    def _edit_task(self):
        i_task = self._selected_task_index()
        if i_task is None: return
        cur = json.loads(json.dumps(self.store.tasks[i_task]))
        r = self._task_dialog(title="Edit Task", init=cur)
        if not r: return
        self.store.tasks[i_task] = r
        self.store.save()
        self._refresh_todo_feed(); self._draw_calendar()

    def _selected_task_index(self):
        tv = getattr(self, "todo_tv", None)
        if tv and tv.winfo_exists():
            sel = tv.selection()
            if sel:
                iid = sel[0]
                i_task, _ = self._todo_rows.get(iid, (None, None))
                return i_task
            focus = tv.focus()
            if focus:
                i_task, _ = self._todo_rows.get(focus, (None, None))
                return i_task
        return None

    def _stop_recurring(self):
        i_task = self._selected_task_index()
        if i_task is None:
            return
        t = self.store.tasks[i_task]
        rec = (t.get("recurrence") or {})
        if (rec.get("freq") or "one-off") == "one-off":
            messagebox.showinfo("Stop Recurrence", "This task is not recurring.")
            return
        today = _dt.date.today().isoformat()
        t["end_on"] = today  # cap future occurrences; past remains
        # keep is_enabled = True so past still shows
        self.store.save()
        self._refresh_todo_feed()
        self._draw_calendar()

    def _pause_recurring(self):
        i_task = self._selected_task_index()
        if i_task is None:
            return
        t = self.store.tasks[i_task]
        rec = (t.get("recurrence") or {})
        if (rec.get("freq") or "one-off") == "one-off":
            messagebox.showinfo("Pause Recurrence", "This task is not recurring.")
            return
        if t.get("is_paused"):
            messagebox.showinfo("Pause Recurrence", "This recurring task is already paused.")
            return

        today = _dt.date.today()
        # Use the same logic as batch pause: start from the first to-do
        # after the most recent done/cancelled occurrence.
        pause_start = self.store.compute_pause_start_date(t, today=today) or today

        t["pause_from"] = pause_start.isoformat()
        # Clear any previous resume marker; this new pause period supersedes it
        t["resume_from"] = ""
        t["is_paused"] = True
        self.store.save()
        self._refresh_todo_feed()
        self._draw_calendar()


    def _resume_recurring(self):
        i_task = self._selected_task_index()
        if i_task is None:
            return
        t = self.store.tasks[i_task]
        rec = (t.get("recurrence") or {})
        if (rec.get("freq") or "one-off") == "one-off":
            messagebox.showinfo("Resume Recurrence", "This task is not recurring.")
            return
        if not t.get("is_paused"):
            messagebox.showinfo("Resume Recurrence", "This recurring task is not paused.")
            return
        today = _dt.date.today().isoformat()
        t["resume_from"] = today
        t["is_paused"] = False
        self.store.save()
        self._refresh_todo_feed()
        self._draw_calendar()


    # -------------- Task dialog (modernized layout) --------------
    def _task_dialog(self, title, init=None):
        init = init or {}
        d = tk.Toplevel(self.app); d.title(title); d.resizable(False, False)
        frm = ttk.Frame(d, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        _title_locked = {"value": False}

        v_kind = tk.StringVar(value=(init.get("kind","payroll") or "other").upper())
        v_company_name = tk.StringVar(value= init.get("company_name") or (self._company_name(init.get("company_idx")) or ""))
        rec = init.get("recurrence", {"freq":"one-off"})

        _original_task_id   = init.get("id", "")
        _original_start_on  = init.get("start_on", "")
        _original_anchor    = rec.get("anchor_date", "")

        v_mode = tk.StringVar(value=("recurring" if rec.get("freq","one-off")!="one-off" else "one-off"))
        v_freq = tk.StringVar(value=rec.get("freq","one-off"))
        v_due  = tk.StringVar(value=init.get("due",""))
        v_title = tk.StringVar()
        v_enabled = tk.BooleanVar(value=init.get("is_enabled", True))
        v_notify = tk.IntVar(value=int(init.get("notify_days", 4)))
        v_method = tk.StringVar(value=init.get("method", "none"))
        v_action_lead = tk.IntVar(value=int(init.get("action_lead_days", 0)))

        default_dom  = 5 if rec.get("freq") == "semi-monthly" else 5
        default_dom2 = 20 if rec.get("freq") == "semi-monthly" else 20
        v_dom  = tk.IntVar(value=int(rec.get("dom", default_dom)))
        v_dom2 = tk.IntVar(value=int(rec.get("dom2", default_dom2)))
        v_anchor = tk.StringVar(value=rec.get("anchor_date",""))
        wk_val = rec.get("weekday", "")
        try:
            wk_display = WEEKDAY_NAMES[int(wk_val)] if wk_val != "" else ""
        except Exception:
            wk_display = ""
        v_weekday = tk.StringVar(value=wk_display)

        v_months = tk.StringVar(value=",".join(str(m) for m in (rec.get("months") or [])))

        def _current_auto_title():
            kind = (v_kind.get() or "").strip().upper()
            comp = (v_company_name.get() or "").strip().upper()
            if (v_mode.get() or "one-off") == "one-off":
                mmdd = ""
                s = (v_due.get() or "").strip()
                try:
                    d_ = _dt.date.fromisoformat(s)
                    mmdd = d_.strftime("%m-%d")
                except Exception:
                    pass
                suffix = f"({mmdd})" if mmdd else "(ONE-OFF)"
            else:
                suffix = f"({(v_freq.get() or '').strip().upper()})"
            base = " ".join(x for x in [kind, comp] if x)
            return f"{base} {suffix}".strip()

        def _maybe_set_auto_title(*_):
            if _title_locked["value"]:
                return
            v_title.set(_current_auto_title())

        def _ensure_freq_when_recurring():
            if (v_mode.get() == "recurring") and (v_freq.get() in ("", "one-off")):
                v_freq.set("monthly")

        def _trace_mode(*_):
            _ensure_due_if_one_off()
            _ensure_freq_when_recurring()
            _maybe_set_auto_title()

        def _ensure_due_if_one_off():
            if (v_mode.get() or "one-off") == "one-off":
                s = (v_due.get() or "").strip()
                if not s:
                    v_due.set(_dt.date.today().isoformat())

        _ensure_due_if_one_off()

        if init.get("title"):
            v_title.set(init["title"]); _title_locked["value"] = True
        else:
            v_title.set(_current_auto_title())

        for var in (v_kind, v_freq, v_company_name):
            var.trace_add("write", _maybe_set_auto_title)
        v_mode.trace_add("write", _trace_mode)
        v_due.trace_add("write", _trace_mode)

        # Title row (now spans and resizes nicely)
        ttk.Label(frm, text="Title").grid(row=0, column=0, sticky="w")
        title_entry = ttk.Entry(frm, textvariable=v_title, width=48)
        title_entry.grid(row=0, column=1, columnspan=3, sticky="we", padx=(6,0), pady=(0,6))
        def _on_title_key(_evt=None):
            if v_title.get() != _current_auto_title():
                _title_locked["value"] = True
        title_entry.bind("<KeyRelease>", _on_title_key)

        # Row: Type / Enabled / Notify days
        ttk.Label(frm, text="Type").grid(row=1, column=0, sticky="w", pady=(4,0))
        ttk.Combobox(frm, textvariable=v_kind, values=["PAYROLL","BOOKKEEPING","PAYROLL TAX","SALES TAX","CALSAVERS","WORKERS COMP", "OTHER"], width=18, state="readonly").grid(row=1, column=1, sticky="w", padx=(6,0), pady=(4,0))
        ttk.Checkbutton(frm, text="Enabled", variable=v_enabled).grid(row=1, column=2, sticky="w", padx=(12,0), pady=(4,0))
        ttk.Label(frm, text="Notify days").grid(row=1, column=3, sticky="e", pady=(4,0))
        ttk.Entry(frm, textvariable=v_notify, width=6).grid(row=1, column=4, sticky="w", padx=(6,0), pady=(4,0))

        # Company (autocomplete)
        ttk.Label(frm, text="Company").grid(row=2, column=0, sticky="w", pady=(6,0))
        names = [c.get("name","") for c in getattr(self.app, "items", [])]
        AutoCompleteCombobox(frm, textvariable=v_company_name, values=names, width=48).grid(row=2, column=1, columnspan=4, sticky="we", padx=(6,0), pady=(6,0))

        ttk.Separator(frm).grid(row=3, column=0, columnspan=5, sticky="we", pady=10)

        # Mode + Method
        ttk.Label(frm, text="Task kind").grid(row=4, column=0, sticky="w")
        ttk.Combobox(frm, textvariable=v_mode, values=["one-off","recurring"], width=14, state="readonly").grid(row=4, column=1, sticky="w", padx=(6,12))
        ttk.Label(frm, text="Submission").grid(row=4, column=2, sticky="e")
        ttk.Combobox(frm, textvariable=v_method, values=["none","mail","direct_deposit"], width=18, state="readonly").grid(row=4, column=3, sticky="w", padx=(6,0))
        ttk.Label(frm, text="Lead days").grid(row=4, column=4, sticky="e")
        ttk.Entry(frm, textvariable=v_action_lead, width=6).grid(row=4, column=5, sticky="w", padx=(6,0))

        # Frequency (only for recurring)
        freq_row = ttk.Frame(frm)
        freq_row.grid(row=5, column=0, columnspan=6, sticky="w", pady=(6,0))
        ttk.Label(freq_row, text="Frequency").pack(side=tk.LEFT)
        ttk.Combobox(
            freq_row,
            textvariable=v_freq,
            values=["monthly","semi-monthly","weekly","biweekly","quarterly"],
            width=16,
            state="readonly"
        ).pack(side=tk.LEFT, padx=(6,0))

        # Dynamic area
        dyn = ttk.Frame(frm)
        dyn.grid(row=6, column=0, columnspan=6, sticky="we", pady=(8,0))
        for i in range(6): dyn.columnconfigure(i, weight=1)

        # One-off holder
        holder_one = ttk.Frame(dyn)
        ttk.Label(holder_one, text="Due date (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        ttk.Entry(holder_one, textvariable=v_due, width=20).grid(row=0, column=1, sticky="w", padx=(6,0))

        # Recurring holders (only one visible at a time)
        holder_monthly = ttk.Frame(dyn)
        ttk.Label(holder_monthly, text="Day of month").grid(row=0, column=0, sticky="w")
        ttk.Entry(holder_monthly, textvariable=v_dom, width=6).grid(row=0, column=1, sticky="w", padx=(6,0))

        holder_semi = ttk.Frame(dyn)
        ttk.Label(holder_semi, text="Days of month (e.g., 5 & 20)").grid(row=0, column=0, sticky="w")
        ttk.Entry(holder_semi, textvariable=v_dom, width=6).grid(row=0, column=1, sticky="w", padx=(6,0))
        ttk.Entry(holder_semi, textvariable=v_dom2, width=6).grid(row=0, column=2, sticky="w", padx=(6,0))


        _wk_values = WEEKDAY_NAMES
        holder_weekly = ttk.Frame(dyn)
        ttk.Label(holder_weekly, text="Weekday").grid(row=0, column=0, sticky="w")
        ttk.Combobox(holder_weekly, textvariable=v_weekday, values=_wk_values, width=12, state="readonly").grid(row=0, column=1, sticky="w", padx=(6,12))
        ttk.Label(holder_weekly, text="Anchor date (optional, YYYY-MM-DD or YYYYMMDD)").grid(row=1, column=0, sticky="w")
        ttk.Entry(holder_weekly, textvariable=v_anchor, width=20).grid(row=1, column=1, sticky="w", padx=(6,0))

        holder_biweekly = ttk.Frame(dyn)
        ttk.Label(holder_biweekly, text="Weekday").grid(row=0, column=0, sticky="w")
        ttk.Combobox(holder_biweekly, textvariable=v_weekday, values=_wk_values, width=12, state="readonly").grid(row=0, column=1, sticky="w", padx=(6,12))
        ttk.Label(holder_biweekly, text="Anchor date (optional, YYYY-MM-DD or YYYYMMDD)").grid(row=1, column=0, sticky="w")
        ttk.Entry(holder_biweekly, textvariable=v_anchor, width=20).grid(row=1, column=1, sticky="w", padx=(6,0))

        holder_quarterly = ttk.Frame(dyn)
        ttk.Label(holder_quarterly, text="Day of month").grid(row=0, column=0, sticky="w")
        ttk.Entry(holder_quarterly, textvariable=v_dom, width=6).grid(row=0, column=1, sticky="w", padx=(6,0))
        ttk.Label(holder_quarterly, text="Quarter months (e.g., 1,4,7,10)").grid(row=1, column=0, sticky="w")
        ttk.Entry(holder_quarterly, textvariable=v_months, width=20).grid(row=1, column=1, sticky="w", padx=(6,0))

        def _render_dyn(*_):
            # Hide all first
            for w in (holder_one, holder_monthly, holder_semi, holder_weekly, holder_biweekly, holder_quarterly):
                w.grid_forget()

            if v_mode.get() == "one-off":
                holder_one.grid(row=0, column=0, sticky="w")
                try: freq_row.grid_remove()
                except Exception: pass
            else:
                try: freq_row.grid()  # show frequency when recurring
                except Exception: pass
                _ensure_freq_when_recurring()
                f = (v_freq.get() or "").lower()
                if f == "monthly":
                    holder_monthly.grid(row=0, column=0, sticky="w")
                elif f == "semi-monthly":
                    holder_semi.grid(row=0, column=0, sticky="w")
                elif f == "weekly":
                    holder_weekly.grid(row=0, column=0, sticky="w")
                elif f == "biweekly":
                    holder_biweekly.grid(row=0, column=0, sticky="w")
                elif f == "quarterly":
                    holder_quarterly.grid(row=0, column=0, sticky="w")

            _ensure_due_if_one_off()
            _maybe_set_auto_title()


        for var in (v_mode, v_freq):
            var.trace_add("write", lambda *_: _render_dyn())
        _render_dyn()


        out = {}

        def _parse_date_local(s):
            if not s: return None
            s = str(s).strip()
            try: return _dt.date.fromisoformat(s)
            except Exception: pass
            if len(s) == 8 and s.isdigit():
                try: return _dt.datetime.strptime(s, "%Y%m%d").date()
                except Exception: return None
            return None

        def _normalize_date_str(s: str) -> str:
            d = _parse_date_local(s)
            return d.isoformat() if d else (s or "").strip()

        def _month_last_day_local(y, m):
            import calendar as _cal_
            return _cal_.monthrange(y, m)[1]

        def _ok():
            title_txt = v_title.get().strip()
            if not title_txt:
                messagebox.showerror("Validation", "Title is required"); return
            kind = (v_kind.get() or "other").strip().upper()
            company_name = v_company_name.get().strip() or None
            company_idx = next((i for i,c in enumerate(getattr(self.app, "items", [])) if c.get("name") == company_name), None)
            method = v_method.get(); action_lead = int(v_action_lead.get())
            if method in ("mail","direct_deposit"): action_lead = max(action_lead, 2)
            if v_mode.get() == "one-off":
                rec_ = {"freq": "one-off"}; due = v_due.get().strip()
            else:
                f = v_freq.get(); rec_ = {"freq": f}; due = ""
                if f == "monthly": rec_.update(dom=int(v_dom.get()))
                elif f == "semi-monthly": rec_.update(dom=int(v_dom.get()), dom2=int(v_dom2.get()))
                elif f in ("biweekly","weekly"):
                    wds = (v_weekday.get() or "").strip()
                    anc = _normalize_date_str((v_anchor.get() or "").strip())
                    if anc:
                        rec_.update(anchor_date=anc)              # user explicitly set it
                    elif _original_anchor:
                        rec_.update(anchor_date=_original_anchor) # keep original on edit
                    # else: leave it absent for brand-new tasks with no anchor
                    
                    if wds:
                        idx = WD_NAME_TO_INT.get(wds)
                        if idx is not None:
                            rec_.update(weekday=idx)

                elif f == "quarterly":
                    months = [int(x) for x in v_months.get().replace(" ", "").split(",") if x]
                    rec_.update(dom=int(v_dom.get()), months=months)

            start_on = _original_start_on or (init.get("start_on") or None)
            try:
                today = _dt.date.today()
                if (v_mode.get() or "one-off") != "one-off" and not _original_task_id and not start_on:
                    f = (rec_.get("freq") or "").lower()
                    if f == "monthly":
                        dom = int(v_dom.get())
                        y, m = today.year, today.month
                        dom = min(dom, _month_last_day_local(y, m))
                        if today.day > dom:
                            m = 1 if m == 12 else m + 1
                            y = y + 1 if m == 1 else y
                            dom = min(dom, _month_last_day_local(y, m))
                        start_on = _dt.date(y, m, dom).isoformat()
                    elif f == "semi-monthly":
                        d1, d2 = sorted((int(v_dom.get()), int(v_dom2.get())))
                        cands = []
                        for dd in (d1, d2):
                            dd = min(dd, _month_last_day_local(today.year, today.month))
                            if today.day <= dd:
                                cands.append(_dt.date(today.year, today.month, dd))
                        if not cands:
                            y = today.year + (today.month == 12)
                            m = 1 if today.month == 12 else today.month + 1
                            for dd in (d1, d2):
                                dd = min(dd, _month_last_day_local(y, m))
                                cands.append(_dt.date(y, m, dd))
                        start_on = min(cands).isoformat()
                    elif f in ("weekly","biweekly"):
                        wd_s = (v_weekday.get() or "").strip()
                        anc = _normalize_date_str((v_anchor.get() or "").strip())
                        if anc:
                            start_on = anc
                        elif wd_s:
                            idx = WD_NAME_TO_INT.get(wd_s)
                            if idx is not None:
                                offset = (idx - today.weekday()) % 7
                                start_on = (today + _dt.timedelta(days=offset)).isoformat()

                    elif f == "quarterly":
                        months = [int(x) for x in (v_months.get() or "").replace(" ", "").split(",") if x] or [1,4,7,10]
                        dom = int(v_dom.get())
                        y0, m0 = today.year, today.month
                        for step in range(0, 24):
                            y = y0 + (m0 - 1 + step) // 12
                            m = (m0 - 1 + step) % 12 + 1
                            if m in months:
                                dd = min(dom, _month_last_day_local(y, m))
                                cand = _dt.date(y, m, dd)
                                if cand >= today:
                                    start_on = cand.isoformat(); break
            except Exception:
                pass

            out.update({
                "id": init.get("id", str(uuid.uuid4())),
                "title": title_txt,
                "kind": kind,
                "category": init.get("category", "other"),
                "company_idx": company_idx,
                "company_name": company_name,
                "start_on": start_on or "",
                "recurrence": rec_,
                "due": due,
                "due_time": init.get("due_time", ""),
                "notify_days": int(v_notify.get()),
                "completed": init.get("completed", []),
                "is_enabled": bool(v_enabled.get()),
                # keep cancellation & pause related flags when editing
                "cancelled": init.get("cancelled", []),
                "end_on": init.get("end_on", ""),
                "is_paused": init.get("is_paused", False),
                "pause_from": init.get("pause_from", ""),
                "resume_from": init.get("resume_from", ""),
                "notes": init.get("notes", ""),
                "method": method,
                "action_lead_days": action_lead,
            })

            d.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=99, column=0, columnspan=6, sticky="e", pady=(12,0))
        ttk.Button(btns, text="Cancel", command=d.destroy).pack(side=tk.RIGHT, padx=(8,0))
        ttk.Button(btns, text="Save", style="Accent.TButton", command=_ok).pack(side=tk.RIGHT)
        d.grab_set(); self.app.wait_window(d)
        return out if out else None

    # -------------- helpers --------------
    def _company_name(self, idx):
        items = getattr(self.app, "items", [])
        if isinstance(idx, int) and 0 <= idx < len(items):
            return items[idx].get("name")
        return None

    def _shift_month(self, delta):
        y, m = self._cal_year, self._cal_month
        m = m + delta
        y = y + (m-1)//12
        m = (m-1)%12 + 1
        self._cal_year, self._cal_month = y, m
        self._draw_calendar()

            
    def _stripe_and_merge(self, tv: ttk.Treeview) -> None:
        # Merge zebra striping tags with our semantic tags so colors persist
        for iid in tv.get_children(""):
            semantic = list(self._row_tags.get(iid, ()))
            current = list(tv.item(iid, "tags") or ())
            # keep any stripe tags but ensure semantic tags remain
            merged = [t for t in current if t not in semantic] + semantic
            # priority: submission > due (these go last so their backgrounds win)
            for prio in ("submission", "due"):
                if prio in merged:
                    merged = [t for t in merged if t != prio] + [prio]
            tv.item(iid, tags=tuple(merged))

    def _todo_show_context_menu(self, event):
        tv = self.todo_tv
        row = tv.identify_row(event.y)
        if not row:
            return "break"

        # ensure selection follows cursor
        tv.selection_set(row)

        i_task, date_obj = self._todo_rows.get(row, (None, None))
        if i_task is None or date_obj is None:
            return "break"

        # (Re)build menu each time to keep it simple
        import tkinter as tk
        if self._todo_ctx is None:
            self._todo_ctx = tk.Menu(tv, tearoff=0)
        else:
            self._todo_ctx.delete(0, "end")

        self._todo_ctx.add_command(
            label="Done",
            command=lambda: self._set_row_state(i_task, date_obj, "done"),
        )
        self._todo_ctx.add_command(
            label="To-do",
            command=lambda: self._set_row_state(i_task, date_obj, "todo"),
        )
        self._todo_ctx.add_command(
            label="Cancel",
            command=lambda: self._set_row_state(i_task, date_obj, "cancel"),
        )

        try:
            self._todo_ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self._todo_ctx.grab_release()

        return "break"

    def _set_row_state(self, i_task, date_obj, state):
        task = self.store.tasks[i_task]
        self.store.set_state_for_date(task, date_obj, state)
        # redraw both list + calendar so strikethrough/markers stay in sync
        self._refresh_todo_feed()
        self._draw_calendar()