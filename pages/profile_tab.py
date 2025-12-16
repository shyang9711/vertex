from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))
        
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import json
import calendar as _cal
import datetime as _dt
try:
    from functions.models.tasks_model import (
        adjust_if_weekend_or_holiday,
        next_monthly_on_or_after,
        next_semi_monthly_on_or_after,
        next_quarterly_on_or_after,
        DUE_DATE,
        calc_tags_for_occurrence,
        display_date_for,
    )
    from functions.utils.app_logging import get_logger
    
except ModuleNotFoundError:
    from models.tasks_model import (
        adjust_if_weekend_or_holiday,
        next_monthly_on_or_after,
        next_semi_monthly_on_or_after,
        next_quarterly_on_or_after,
        DUE_DATE,
        calc_tags_for_occurrence,
        display_date_for,
    )
    from utils.app_logging import get_logger
LOG = get_logger("profile")


def _ensure_officer_dict(x: dict) -> dict:
    if not isinstance(x, dict):
        x = {"name": str(x).strip()}
    o = {
        "name":        str(x.get("name","")).strip(),
        "first_name":  str(x.get("first_name","")).strip(),
        "middle_name": str(x.get("middle_name","")).strip(),
        "last_name":   str(x.get("last_name","")).strip(),
        "nickname":    str(x.get("nickname","")).strip(),
        "email":       str(x.get("email","")).strip(),
        "phone":       str(x.get("phone","")).strip(),
        "addr1":       str(x.get("addr1","")).strip(),
        "addr2":       str(x.get("addr2","")).strip(),
        "city":        str(x.get("city","")).strip(),
        "state":       str(x.get("state","")).strip(),
        "zip":         str(x.get("zip","")).strip(),
        "dob":         str(x.get("dob","")).strip(),
    }
    
    if not o["name"]:
        parts = [o["first_name"], o["middle_name"], o["last_name"]]
        base = " ".join([p for p in parts if p]).strip()
        if o["nickname"]:
            o["name"] = f'{base} ("{o["nickname"]}")' if base else o["nickname"]
        else:
            o["name"] = base
    return o

def _parse_date_flex_local(s):
    s = (s or "").strip()
    if not s:
        return None
    # ISO first
    try:
        return _dt.date.fromisoformat(s)
    except Exception:
        pass
    # Compact YYYYMMDD
    if len(s) == 8 and s.isdigit():
        try:
            return _dt.datetime.strptime(s, "%Y%m%d").date()
        except Exception:
            return None
    return None


def _month_last_day(y: int, m: int) -> int:
    return _cal.monthrange(y, m)[1]

WEEKDAY_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
WD_NAME_TO_INT = {name: i for i, name in enumerate(WEEKDAY_NAMES)}

def init_profile_tab(
    nb: ttk.Notebook,
    app,                    # the App instance (needs .navigate, ._edit_rates, ._refresh_sales_tax_for)
    company: dict,
    edit_rates_cb=None,     # optional callback for "Edit Rates"
    refresh_sales_cb=None   # optional callback for "Refresh Sales Tax"
):
    """
    Builds the Profile tab into the provided Notebook.
    Returns the created Frame.
    """
    prof = ttk.Frame(nb, padding=8)
    nb.add(prof, text="Profile")
    prof.columnconfigure(0, weight=1, uniform="cols")
    prof.columnconfigure(1, weight=1, uniform="cols")

    # Left & Right columns
    left = ttk.Frame(prof); left.grid(row=0, column=0, sticky="nsew", padx=(0,8))
    right = ttk.Frame(prof); right.grid(row=0, column=1, sticky="nsew")

    # ---------- RIGHT: Personnel ----------
    ttk.Label(right, text="Personnel", font=("Segoe UI", 11, "bold")).pack(anchor="w")
    people_cols = ("role","first","last","email","phone")
    people_wrap = ttk.Frame(right)
    people_wrap.pack(fill=tk.BOTH, expand=True, pady=(4,8))

    ysb = ttk.Scrollbar(people_wrap, orient="vertical")
    _profile_style = ttk.Style(prof)
    _profile_style.configure("Profile.Treeview", rowheight=40)
    people_tree = ttk.Treeview(people_wrap, columns=people_cols, show="headings", height=5, selectmode="browse", style="Profile.Treeview")
    for col, label, w in (("role","Role",80),("first","First",100),("last","Last",100),("email","Email",200),("phone","Phone",100)):
        people_tree.heading(col, text=label)
        people_tree.column(col, width=w, anchor="w", stretch=True)
    ysb.config(command=people_tree.yview)

    people_tree.pack(side="left", fill=tk.BOTH, expand=True)
    ysb.pack(side="right", fill="y")

    # Map tree item -> (role_key, index)
    person_index_map = {}

    def _add_people(role_key: str, items_list):
        for i, o in enumerate(items_list or []):
            o2 = _ensure_officer_dict(o)
            iid = people_tree.insert(
                "", "end",
                values=(
                    "Officer" if role_key == "officers" else "Employee",
                    o2.get("first_name",""),
                    o2.get("last_name",""),
                    o2.get("email",""),
                    o2.get("phone",""),
                )
            )
            person_index_map[iid] = (role_key, i)

    _add_people("officers",  company.get("officers", []))
    _add_people("employees", company.get("employees", []))
    people_tree.configure(height=min(5, len(person_index_map)))

    def _open_person_page(_e=None):
        LOG.info("Open person detail from Profile")
        sel = people_tree.selection()
        if not sel: return
        iid = sel[0]
        role_key, pidx = person_index_map.get(iid, ("officers", 0))
        # delegate navigation to app
        # NOTE: app.navigate("person", (idx, role_key, pidx), push=True) requires current company idx.
        # The caller (client_manager) attaches this tab inside a detail page where app._current_detail_idx is set.
        company_idx = _resolve_company_idx_from_company()
        if company_idx is not None:
            app.navigate("person", payload=(company_idx, role_key, pidx), push=True)

    people_tree.bind("<Double-1>", _open_person_page)

    # ---------- RIGHT: Company Tasks (Dashboard-like) ----------
    ttk.Label(right, text="Tasks", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8,2))

    toolbar = ttk.Frame(right)
    toolbar.pack(fill="x", pady=(0,4))
    # Sync "Show all past" with dashboard if available
    right._show_all_past_var = tk.BooleanVar(value=getattr(getattr(app, "dashboard", None), "_show_all_past", False))
    _show_all_past_var = right._show_all_past_var

    def _flip_show_all_past():
        val = bool(_show_all_past_var.get())
        if hasattr(app, "dashboard") and getattr(app.dashboard, "_show_all_past", None) is not None:
            app.dashboard._show_all_past = val
        _refresh_company_tasks_tv()

    ttk.Checkbutton(toolbar, text="Show all past", variable=_show_all_past_var,
                    command=_flip_show_all_past).pack(side="left")
    ttk.Button(toolbar, text="Add Task", command=lambda: _open_add_task_dialog()).pack(side="left", padx=(8,0))
    ttk.Button(toolbar, text="Edit",   command=lambda: _edit_company_task()).pack(side="left", padx=(6,0))
    ttk.Button(toolbar, text="Delete", command=lambda: _delete_company_task()).pack(side="left", padx=(6,0))
    ttk.Button(toolbar, text="Stop Recurrence", command=lambda: _stop_company_recurring()).pack(side="left", padx=(6,0))
    ttk.Button(toolbar, text="Pause Recurrence",  command=lambda: _pause_company_recurring()).pack(side="left", padx=(6,0))
    ttk.Button(toolbar, text="Resume Recurrence", command=lambda: _resume_company_recurring()).pack(side="left", padx=(6,0))

    cols = ("mark","title","kind","due")
    company_tasks_tv = ttk.Treeview(right, columns=cols, show="headings", height=10, selectmode="browse", style="Profile.Treeview")
    company_tasks_tv.heading("mark", text="")
    company_tasks_tv.heading("title", text="Task")
    company_tasks_tv.heading("kind", text="Type")
    company_tasks_tv.heading("due",  text="Deadline")

    company_tasks_tv.column("mark", width=28, anchor="center", stretch=False)
    company_tasks_tv.column("title", width=280, anchor="w")
    company_tasks_tv.column("kind",  width=120, anchor="w")
    company_tasks_tv.column("due",   width=120, anchor="w", stretch=False)

    company_tasks_tv.pack(fill="both", expand=True)
    company_tasks_tv.tag_configure("done", foreground="#6B7280")
    company_tasks_tv.tag_configure("cancelled", foreground="#6B7280")
    company_tasks_tv.tag_configure("due",  background="#FEF3C7")
    company_tasks_tv.tag_configure("todo", foreground="")
    company_tasks_tv.tag_configure("submission", background="#FFEDD5")

    try:
        import tkinter.font as tkfont
        _base = tkfont.nametofont("TkDefaultFont")
        _cancel_font = tkfont.Font(**_base.configure()); _cancel_font.configure(overstrike=1)
        company_tasks_tv.tag_configure("cancelled", font=_cancel_font)
    except Exception:
        pass
    
    _company_todo_rows = {}

    def _safe_redraw_dashboard():
        try:
            d = getattr(app, "dashboard", None)
            if d and getattr(d, "_cal_wrap", None) and d._cal_wrap.winfo_exists():
                d._draw_calendar()
        except Exception:
            pass

    def _company_toggle_row(iid):
        i_task, orig_date = _company_todo_rows.get(iid, (None, None))
        if i_task is None or orig_date is None:
            return
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return
        # Flip done on ORIGINAL due date (same rule as dashboard)
        t = dash.store.tasks[i_task]
        s = orig_date.isoformat()
        comp = t.setdefault("completed", [])
        if s in comp:
            comp.remove(s)
        else:
            comp.append(s)
        # Persist + refresh everywhere
        dash.store.save()
        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()


    def _on_company_tasks_click(e):
        # Only toggle when clicking the first column (checkbox column)
        if company_tasks_tv.identify_region(e.x, e.y) != "cell":
            return
        col = company_tasks_tv.identify_column(e.x)  # '#1' = first column
        if col != "#1":
            return
        row_id = company_tasks_tv.identify_row(e.y)
        if not row_id:
            return
        _company_toggle_row(row_id)
        return "break"  # stop default selection flicker
    
    def _on_company_tasks_dbl(e):
        # Don’t conflict with the checkbox column
        if company_tasks_tv.identify_region(e.x, e.y) != "cell":
            return
        if company_tasks_tv.identify_column(e.x) == "#1":
            return  # first column = checkbox
        _edit_company_task()

    def _on_company_tasks_delete(e):
        _delete_company_task()
        return "break"

    def _on_company_tasks_enter(e):
        _edit_company_task()
        return "break"

    company_tasks_tv.bind("<Button-1>", _on_company_tasks_click)
    company_tasks_tv.bind("<Delete>", _on_company_tasks_delete)
    company_tasks_tv.bind("<Return>", _on_company_tasks_enter)
    company_tasks_tv.bind("<Double-1>", _on_company_tasks_dbl)

    _company_ctx = None

    def _company_show_context_menu(e):
        nonlocal _company_ctx
        row_id = company_tasks_tv.identify_row(e.y)
        if not row_id:
            return
        company_tasks_tv.selection_set(row_id)

        i_task, orig_date = _company_todo_rows.get(row_id, (None, None))
        if i_task is None:
            return
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return
        t = dash.store.tasks[i_task]

        import tkinter as tk
        # ✅ Build fresh each time (or delete all items if you prefer to reuse the instance)
        if _company_ctx is None:
            _company_ctx = tk.Menu(company_tasks_tv, tearoff=False)
        else:
            _company_ctx.delete(0, "end")  # <<< IMPORTANT: clear old commands

        def _set_state(state):
            dash.store.set_state_for_date(t, orig_date, state)
            dash._refresh_todo_feed()
            _safe_redraw_dashboard()
            _refresh_company_tasks_tv()

        _company_ctx.add_command(label="Done",    command=lambda: _set_state("done"))
        _company_ctx.add_command(label="To-do",   command=lambda: _set_state("todo"))
        _company_ctx.add_command(label="Cancel",  command=lambda: _set_state("cancel"))

        _company_ctx.tk_popup(e.x_root, e.y_root)


    company_tasks_tv.bind("<Button-3>", _company_show_context_menu)

    def _selected_company_task_ref():
        sel = company_tasks_tv.selection()
        if not sel:
            focus = company_tasks_tv.focus()
            if not focus:
                return (None, None)
            sel = (focus,)
        iid = sel[0]
        return _company_todo_rows.get(iid, (None, None))

    def _selected_company_task_index():
        i_task, _ = _selected_company_task_ref()
        return i_task
    
    # ---- Helpers
    def _adjust_display_bridge(d):
        return adjust_if_weekend_or_holiday(d)

    def _month_last_day_local(y, m): return _cal.monthrange(y, m)[1]
    def _parse_date_local(s):
        try: return _dt.date.fromisoformat(s) if s else None
        except Exception: return None

    def _resolve_company_idx_from_company():
        items = getattr(app, "items", [])
        name = (company.get("name") or "").strip()
        for i, c in enumerate(items):
            if (c.get("name") or "").strip() == name:
                return i
        return getattr(app, "_current_detail_idx", None)
    
    import uuid

    def _monday_of_this_week(d):
        return d - _dt.timedelta(days=d.weekday())  # Monday

    def _start_of_this_month(d):
        return d.replace(day=1)

    def _start_anchor_for_recur(freq, today=None):
        today = today or _dt.date.today()
        if freq in ("weekly", "biweekly"):
            # “current week” → Monday of the current week
            return _monday_of_this_week(today)
        if freq in ("monthly", "semi-monthly", "quarterly"):
            # start this month so we don’t backfill years of history
            return _start_of_this_month(today)
        return today  # one-off default

    def _open_add_task_dialog():
        """
        Use the Dashboard's modern task dialog, prefilled for this company.
        """
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            messagebox.showinfo("Add Task", "Dashboard is not available.")
            return
    
        import uuid
        comp_idx = _resolve_company_idx_from_company()
        comp_name = company.get("name") or ""
    
        # Default the Type to PAYROLL (you can change this)
        init = {
            "title": "",
            "kind": "PAYROLL",
            "company_idx": comp_idx,
            "company_name": comp_name,
            "is_enabled": True,
            "notify_days": 4,
            "method": "none",
            "action_lead_days": 0,
            "recurrence": {"freq": "one-off"},  # user can switch to recurring in the dialog
            "due": _dt.date.today().isoformat(),
        }
    
        # Open the same dialog the Dashboard uses
        cur = dash._task_dialog(title="Add Task", init=init)
        if not cur:
            return
    
        # Save like Dashboard
        cur.setdefault("id", str(uuid.uuid4()))
        dash.store.tasks.append(cur)
        dash.store.save()
    
        # Refresh both places
        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()

    def _occurs_on_local(t, day):
        rec = t.get("recurrence", {"freq":"one-off"})
        freq = rec.get("freq", "one-off")
        if freq == "one-off":
            d = _parse_date_local(t.get("due"))
            return (d == day)
        if freq == "monthly":
            dom = min(int(rec.get("dom", 1)), _month_last_day_local(day.year, day.month))
            return day.day == dom
        if freq == "semi-monthly":
            d1 = min(int(rec.get("dom", 5)),  _month_last_day_local(day.year, day.month))
            d2 = min(int(rec.get("dom2",20)), _month_last_day_local(day.year, day.month))
            return day.day in (d1, d2)
        if freq == "weekly":
            anchor = _parse_date_local(rec.get("anchor_date"))
            wk = rec.get("weekday")
            if wk is not None:
                start_on = _parse_date_local(t.get("start_on")) or anchor or _monday_of_this_week(day)
                if day < start_on or day.weekday() != int(wk):
                    return False
                return ((day - start_on).days % 7) == 0
            return bool(anchor) and (day - anchor).days >= 0 and ((day - anchor).days % 7 == 0)

        if freq == "biweekly":
            anchor = _parse_date_local(rec.get("anchor_date"))
            wk = rec.get("weekday")
            if wk is not None:
                start_on = _parse_date_local(t.get("start_on")) or anchor or _monday_of_this_week(day)
                if day < start_on or day.weekday() != int(wk):
                    return False
                return ((day - start_on).days % 14) == 0
            return bool(anchor) and (day - anchor).days >= 0 and ((day - anchor).days % 14 == 0)

        if freq == "quarterly":
            months = rec.get("months") or [1,4,7,10]
            if day.month not in months: return False
            dom = min(int(rec.get("dom", 15)), _month_last_day_local(day.year, day.month))
            return day.day == dom
        return False

    # Prefer dashboard helpers if present (holiday/weekend adjustment & occurs_on)
    def _occurs_on_bridge(t, day):
        dash = getattr(app, "dashboard", None)
        if dash and getattr(dash, "store", None):
            return dash.store.occurs_on(t, day)
        return _occurs_on_local(t, day)


    # Source of tasks for THIS company (match by idx OR by normalized name)
    def _company_tasks_source():
        idx = _resolve_company_idx_from_company()
        name_key = (company.get("name") or "").strip().lower()
        out = []
        dash = getattr(app, "dashboard", None)
        if dash and getattr(dash, "store", None):
            for t in dash.store.tasks:
                if not t.get("is_enabled", True):
                    continue
                if (idx is not None and t.get("company_idx") == idx) or \
                ((t.get("company_name") or "").strip().lower() == name_key):
                    out.append(t)
        return out

    def _stop_company_recurring():
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return
        i_task = _selected_company_task_index()
        if i_task is None:
            messagebox.showinfo("Stop Recurrence", "Select a task first.")
            return
        t = dash.store.tasks[i_task]
        rec = (t.get("recurrence") or {})
        if (rec.get("freq") or "one-off") == "one-off":
            messagebox.showinfo("Stop Recurrence", "This task is not recurring.")
            return
        t["end_on"] = _dt.date.today().isoformat()
        dash.store.save()
        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()

    def _pause_company_recurring():
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return
        i_task = _selected_company_task_index()
        if i_task is None:
            messagebox.showinfo("Pause Recurrence", "Select a task first.")
            return

        t = dash.store.tasks[i_task]
        rec = (t.get("recurrence") or {})
        if (rec.get("freq") or "one-off") == "one-off":
            messagebox.showinfo("Pause Recurrence", "This task is not recurring.")
            return

        if t.get("is_paused"):
            messagebox.showinfo("Pause Recurrence", "This recurring task is already paused.")
            return

        today = _dt.date.today().isoformat()
        t["pause_from"] = today
        # Clear any previous resume marker; this new pause period supersedes it
        t["resume_from"] = ""
        t["is_paused"] = True

        dash.store.save()
        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()

    def _resume_company_recurring():
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return
        i_task = _selected_company_task_index()
        if i_task is None:
            messagebox.showinfo("Resume Recurrence", "Select a task first.")
            return

        t = dash.store.tasks[i_task]
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

        dash.store.save()
        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()


    def _edit_company_task():
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return

        i_task, occ_date = _selected_company_task_ref()
        if i_task is None:
            messagebox.showinfo("Edit Task", "Select a task first.")
            return

        cur = json.loads(json.dumps(dash.store.tasks[i_task]))  # deep copy
        r = dash._task_dialog(title="Edit Task", init=cur)
        if not r:
            return

        rec = (cur.get("recurrence") or {})
        freq = (rec.get("freq") or "one-off").lower()

        # If a recurring occurrence row is being edited, preserve history by splitting
        if freq != "one-off" and occ_date:
            dash.store.split_recurring_task_from_date(i_task, r, occ_date)
        else:
            dash.store.tasks[i_task] = r
            dash.store.save()

        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()


    def _delete_company_task():
        dash = getattr(app, "dashboard", None)
        if not dash or not getattr(dash, "store", None):
            return
        i_task = _selected_company_task_index()
        if i_task is None:
            messagebox.showinfo("Delete Task", "Select a task first.")
            return
        if not messagebox.askyesno("Confirm", "Delete this task?"):
            return
        del dash.store.tasks[i_task]
        dash.store.save()
        dash._refresh_todo_feed()
        _safe_redraw_dashboard()
        _refresh_company_tasks_tv()

    def _refresh_company_tasks_tv():
        company_tasks_tv.delete(*company_tasks_tv.get_children())
        _company_todo_rows.clear()

        tasks = _company_tasks_source()
        if not tasks:
            return

        today = _dt.date.today()
        past_days = 9999 if (hasattr(app, "dashboard") and getattr(app.dashboard, "_show_all_past", False)) else 7
        window_start = today - _dt.timedelta(days=past_days)
        window_end   = today + _dt.timedelta(days=14)
        yesterday = today - _dt.timedelta(days=1)

        def occurs_between(t, start_d, end_d):
            comp = set(t.get("completed", []) or [])
            start_on = _parse_date_local(t.get("start_on"))
            d = start_d - _dt.timedelta(days=5)
            stop = end_d + _dt.timedelta(days=5)
            while d <= stop:
                if _occurs_on_bridge(t, d):
                    orig = d
                    if start_on and orig < start_on:
                        d += _dt.timedelta(days=1)
                        continue
                    disp = display_date_for(t, orig)
                    yield orig, disp, (orig.isoformat() in comp)
                d += _dt.timedelta(days=1)


        rows = []

        # 1) Past (in window): keep (gray if done)
        for t in tasks:
            kind  = t.get("kind","")
            title = t.get("title","")
            for orig, disp, is_done in occurs_between(t, window_start, yesterday):
                if window_start <= disp <= yesterday:
                    rows.append((disp, is_done, t, kind, title, orig))

        # 2) Today & future: nearest-upcoming rule (match Dashboard)
        for t in tasks:
            kind  = t.get("kind","")
            title = t.get("title","")
            futures = [(disp, is_done, t, kind, title, orig)
                    for (orig, disp, is_done) in occurs_between(t, today, window_end)
                    if today <= disp <= window_end]
            if not futures:
                continue
            futures.sort(key=lambda x: x[0])      # earliest display date first
            first = futures[0]
            if first[1] is True:
                # earliest already done — keep it (gray) and also show next pending if any
                rows.append(first)
                nxt = next((f for f in futures[1:] if f[1] is False), None)
                if nxt:
                    rows.append(nxt)
            else:
                # earliest pending — show only that one
                rows.append(first)

        # --- Dedupe by (task id, display date) ---
        seen = set()
        deduped = []
        for disp, is_done, t, kind, title, orig in rows:
            tid = t.get("id")
            key = (tid, disp)
            if key in seen:
                continue
            seen.add(key)
            deduped.append((disp, is_done, t, kind, title, orig))
        rows = deduped

        # Paint (deadline shows DISPLAY date)
        rows.sort(key=lambda r: r[0], reverse=True)

        # Map task id -> dashboard index so toggling persists
        index_by_id = {}
        dash = getattr(app, "dashboard", None)
        if dash and getattr(dash, "store", None):
            for i, tt in enumerate(dash.store.tasks):
                index_by_id[tt.get("id")] = i

        for disp, is_done, t, kind, title, orig in rows:
            canc = set(t.get("cancelled", []) or [])
            is_cancelled = (orig.isoformat() in canc) or (disp.isoformat() in canc)

            if is_cancelled:
                mark = "×"
            else:
                mark = "☑" if is_done else "☐"

            iid = company_tasks_tv.insert("", "end", values=(mark, title, kind, disp.isoformat()))
            i_task = index_by_id.get(t.get("id"))
            _company_todo_rows[iid] = (i_task, orig)

            tags = list(calc_tags_for_occurrence(t, disp, (is_done or is_cancelled), today))
            if (not is_done) and (not is_cancelled) and int(t.get("action_lead_days", 0) or 0) > 0:
                if "submission" not in tags:
                    tags.append("submission")
            if is_cancelled and "cancelled" not in tags:
                tags = [tg for tg in tags if tg != "todo"]
                if "done" not in tags:
                    tags.insert(0, "done")
                tags.append("cancelled")

            company_tasks_tv.item(iid, tags=tuple(tags))

                

    # Build now and once more after mount (in case outer state updates post-build)
    _refresh_company_tasks_tv()
    right.after(0, _refresh_company_tasks_tv)


    # ---------- LEFT: IDs/Accounts, Tax Rates, Address, Memo ----------
    ttk.Label(left, text="IDs / Accounts", font=("Segoe UI", 11, "bold")).pack(anchor="w")
    ids = ttk.Frame(left, style="Card.TFrame"); ids.pack(fill=tk.X, pady=(4,8))
    def _line(frame, text): ttk.Label(frame, text=text).pack(anchor="w")
    _line(ids, f"EIN: {company.get('ein','') or '—'}")
    _line(ids, f"EDD Number: {company.get('edd_number','') or '—'}")
    _line(ids, f"Sales Tax Account: {company.get('sales_tax_account','') or '—'}")
    _line(ids, f"Account Manager: {company.get('acct_mgr','') or '—'}")

    ttk.Label(left, text="Tax Rates", font=("Segoe UI", 11, "bold")).pack(anchor="w")
    tr = ttk.Frame(left, style="Card.TFrame"); tr.pack(fill="x", pady=(4,8))
    ttk.Label(tr, text=f"UI Rate (%): {company.get('ui_rate','') or '—'}").grid(row=0, column=0, sticky="w", padx=(0,12))
    ttk.Label(tr, text=f"Sales Tax Rate (%): {company.get('sales_tax_rate','') or '—'}").grid(row=0, column=1, sticky="w", padx=(0,12))
    ttk.Label(tr, text=f"Other: {company.get('other_tax_rates','') or '—'}").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2,0))

    btnr = ttk.Frame(left, style="Card.TFrame"); btnr.pack(fill="x", pady=(2,8))
    ttk.Button(btnr, text="Edit Rates", command=(edit_rates_cb or (lambda: None))).pack(side="left")
    ttk.Button(btnr, text="Refresh Sales Tax (quarterly)", command=(refresh_sales_cb or (lambda: None))).pack(side="left", padx=(6,0))
    
    acct_mgr_display = (company.get("acct_mgr","") or "—")
    ttk.Label(left, text=f"Company Address  (Mgr: {acct_mgr_display})", font=("Segoe UI", 11, "bold")).pack(anchor="w")

    addr = ttk.Frame(left, style="Card.TFrame"); addr.pack(fill=tk.X, pady=(4,8))
    a1 = company.get("addr1","") or "—"
    a2 = company.get("addr2","") or ""
    city = company.get("city","") or ""
    st   = company.get("state","") or ""
    zp   = company.get("zip","") or ""
    ttk.Label(addr, text=a1).pack(anchor="w")
    if a2:
        ttk.Label(addr, text=a2).pack(anchor="w")
    line2 = ", ".join(p for p in [city, st] if p)
    if line2 or zp:
        ttk.Label(addr, text=f"{line2} {zp}".strip()).pack(anchor="w")

    ttk.Label(left, text="Memo", font=("Segoe UI", 11, "bold")).pack(anchor="w")
    memo_txt = ScrolledText(left, width=56, height=4, wrap="word")
    memo_txt.pack(fill=tk.BOTH, expand=True, pady=(4,0))
    memo_txt.insert("1.0", company.get("memo",""))
    memo_txt.configure(state="disabled")

    return prof
