from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))
        
from dataclasses import dataclass, field
from pathlib import Path
import datetime as _dt, json, uuid, calendar as _cal
from typing import Iterable, Iterator, Tuple, List, Dict, Any
from functions.utils.app_logging import get_logger

WEEKDAY_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
WD_NAME_TO_INT = {name: i for i, name in enumerate(WEEKDAY_NAMES)}
DUE_DATE = 3  # highlight window (used by UI, but harmless here)
SUBMISSION_METHODS = {"mail", "direct_deposit"}
LOG = get_logger("tasks_model")

def _parse_date(s: str | None) -> _dt.date | None:
    if not s: return None
    s = str(s).strip()
    try: return _dt.date.fromisoformat(s)
    except Exception: pass
    if len(s) == 8 and s.isdigit():
        try: return _dt.datetime.strptime(s, "%Y%m%d").date()
        except Exception: return None
    return None

def _parse_date_safe(s):
    try:
        return _dt.date.fromisoformat(s) if s else None
    except Exception:
        return None

def _month_last_day(y: int, m: int) -> int:
    return _cal.monthrange(y, m)[1]

def next_monthly_on_or_after(start: _dt.date, dom: int) -> _dt.date:
    y, m = start.year, start.month
    dom = min(int(dom), _month_last_day(y, m))
    cand = _dt.date(y, m, dom)
    if cand >= start:
        return cand
    m = 1 if m == 12 else m + 1
    y = y + 1 if m == 1 else y
    dom = min(int(dom), _month_last_day(y, m))
    return _dt.date(y, m, dom)

def next_semi_monthly_on_or_after(start: _dt.date, d1: int, d2: int) -> _dt.date:
    if d1 > d2: d1, d2 = d2, d1
    y, m = start.year, start.month
    c1 = _dt.date(y, m, min(d1, _month_last_day(y, m)))
    c2 = _dt.date(y, m, min(d2, _month_last_day(y, m)))
    cands = [d for d in (c1, c2) if d >= start]
    if cands: return min(cands)
    m = 1 if m == 12 else m + 1
    y = y + 1 if m == 1 else y
    c1 = _dt.date(y, m, min(d1, _month_last_day(y, m)))
    c2 = _dt.date(y, m, min(d2, _month_last_day(y, m)))
    return min(c1, c2)

def next_quarterly_on_or_after(start: _dt.date, months: List[int], dom: int) -> _dt.date:
    months = sorted(set(int(x) for x in months)) or [1,4,7,10]
    y, m = start.year, start.month
    for off in range(0, 16):  # scan ~4 years worst-case
        yy = y + (m + off - 1) // 12
        mm = (m + off - 1) % 12 + 1
        if mm not in months: continue
        dd = _dt.date(yy, mm, min(int(dom), _month_last_day(yy, mm)))
        if dd >= start: return dd
    return start

def calc_tags_for_occurrence(task, display_date: _dt.date, is_done: bool, today: _dt.date | None = None):
    """
    Return an ordered list of Treeview tags for one task occurrence.
    Order matters for UI background precedence.
    - done/todo base tag
    - 'submission' (mail/direct_deposit) overrides 'due' yellow
    - else 'due' if overdue or due within DUE_DATE days
    """
    LOG.debug("calc_tags_for_occurrence title=%s date=%s done=%s",
                (task.get("title") or task.get("name") or "")[:40], display_date, is_done)
    today = today or _dt.date.today()
    tags = ["done"] if is_done else ["todo"]

    if is_done:
        return tags

    # submission (orange) wins over generic due (yellow)
    method = (task.get("method") or "").lower()
    lead_flag = int(task.get("action_lead_days", 0) or 0) > 0
    if method in SUBMISSION_METHODS or lead_flag:
        tags.append("submission")
        return tags

    if display_date < today or (display_date - today).days <= DUE_DATE:
        tags.append("due")
    return tags


# US holiday helpers (same behavior as current UI)
_US_HOL_CACHE: Dict[int, set] = {}
def _nth_weekday_of_month(year, month, weekday, n):
    d = _dt.date(year, month, 1)
    while d.weekday() != weekday: d += _dt.timedelta(days=1)
    return d + _dt.timedelta(days=(n-1)*7)
def _last_weekday_of_month(year, month, weekday):
    d = _dt.date(year, month, _month_last_day(year, month))
    while d.weekday() != weekday: d -= _dt.timedelta(days=1)
    return d
def _us_federal_holidays(year: int) -> set:
    if year in _US_HOL_CACHE: return _US_HOL_CACHE[year]
    H = {
        _dt.date(year, 1, 1), _dt.date(year, 6, 19), _dt.date(year, 7, 4),
        _dt.date(year,11,11), _dt.date(year,12,25),
        _nth_weekday_of_month(year, 1, 0, 3),
        _nth_weekday_of_month(year, 2, 0, 3),
        _last_weekday_of_month(year, 5, 0),
        _nth_weekday_of_month(year, 9, 0, 1),
        _nth_weekday_of_month(year,10, 0, 2),
        _nth_weekday_of_month(year,11, 3, 4),
    }
    _US_HOL_CACHE[year] = H
    return H
def adjust_if_weekend_or_holiday(d: _dt.date) -> _dt.date:
    while d.weekday() >= 5 or d in _us_federal_holidays(d.year):
        d -= _dt.timedelta(days=1)
    return d

def _is_business_day(d: _dt.date) -> bool:
    """Return True if the date is a weekday and not a US federal holiday."""
    return d.weekday() < 5 and d not in _us_federal_holidays(d.year)

def _shift_business_days(d: _dt.date, days: int) -> _dt.date:
    """Shift a date by a given number of business days (skipping weekends and holidays)."""
    if days == 0:
        return d
    step = 1 if days > 0 else -1
    remaining = abs(days)
    cur = d
    while remaining:
        cur += _dt.timedelta(days=step)
        if _is_business_day(cur):
            remaining -= 1
    return cur

def display_date_for(task: Dict[str, Any], orig: _dt.date) -> _dt.date:
    """
    Compute the UI display date for a task occurrence:
    - If action_lead_days > 0, shift earlier by that many business days
      (weekends/holidays don't count).
    - Orig date is first rolled back off weekends/holidays if needed.
    """
    lead = int(task.get("action_lead_days", 0) or 0)
    base = adjust_if_weekend_or_holiday(orig)
    disp = _shift_business_days(base, -lead) if lead > 0 else base
    return adjust_if_weekend_or_holiday(disp)


@dataclass
class TasksStore:
    data_dir: Path
    app: Any | None = None
    filename: str = "tasks.json"
    tasks: List[Dict[str, Any]] = field(default_factory=list)

    def __post_init__(self):
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.path = (self.data_dir / self.filename)
        self.load()

    # ---------- persistence ----------
    def load(self) -> None:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            data = []
        # normalize
        for t in data:
            t.setdefault("id", str(uuid.uuid4()))
            t["kind"] = str(t.get("kind", "OTHER")).upper()
            t.setdefault("recurrence", {"freq": "one-off"})
            t.setdefault("completed", [])
            t.setdefault("cancelled", [])
            t.setdefault("is_enabled", True)
            # new pause / resume fields
            t.setdefault("is_paused", False)
            t.setdefault("pause_from", "")
            t.setdefault("resume_from", "")
            if t.get("company_idx") is None and t.get("company_name") and self.app:
                idx = next(
                    (i for i, c in enumerate(getattr(self.app, "items", []))
                    if c.get("name") == t["company_name"]),
                    None,
                )
                t["company_idx"] = idx
        self.tasks = data

    def save(self) -> None:
        try:
            self.path.write_text(json.dumps(self.tasks, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception:
            # surface errors in UI callers; here we just raise
            raise

    # ---------- recurrence ----------
    def occurs_on(self, task: Dict[str, Any], day: _dt.date) -> bool:
        rec = task.get("recurrence", {"freq": "one-off"})
        freq = (rec.get("freq") or "one-off").lower()

        end_on = _parse_date_safe(task.get("end_on", ""))
        if end_on and day > end_on:
            return False

        s = task.get("start_on")
        if s:
            try:
                start_on = _dt.date.fromisoformat(s)
                if day < start_on:
                    return False
            except Exception:
                pass

        # ---- pause / resume window ----
        pause_from = _parse_date_safe(task.get("pause_from", ""))
        resume_from = _parse_date_safe(task.get("resume_from", ""))

        # If currently paused, hide occurrences on/after the pause start date
        if task.get("is_paused") and pause_from and day >= pause_from:
            return False

        # After resuming, permanently skip occurrences that fell in the paused window
        # (only if we actually have a real interval: pause_from < resume_from)
        if (
            pause_from
            and resume_from
            and pause_from < resume_from
            and pause_from <= day < resume_from
        ):
            return False

        if freq == "one-off":
            return _parse_date(task.get("due")) == day

        if freq == "monthly":
            dom = min(int(rec.get("dom", 1)), _month_last_day(day.year, day.month))
            return day.day == dom

        if freq == "semi-monthly":
            d1 = min(int(rec.get("dom", 5)),  _month_last_day(day.year, day.month))
            d2 = min(int(rec.get("dom2",20)), _month_last_day(day.year, day.month))
            return day.day in (d1, d2)

        if freq in ("weekly","biweekly"):
            anchor = _parse_date(rec.get("anchor_date"))
            wk = rec.get("weekday")
            period = 7 if freq == "weekly" else 14
            if wk is not None:
                base = _parse_date(task.get("start_on")) or anchor
                if not base or day < base or day.weekday() != int(wk): return False
                return ((day - base).days % period) == 0
            if not anchor: return False
            delta = (day - anchor).days
            return delta >= 0 and (delta % period) == 0

        if freq == "quarterly":
            months = rec.get("months") or [1,4,7,10]
            if day.month not in months: return False
            dom = min(int(rec.get("dom", 15)), _month_last_day(day.year, day.month))
            return day.day == dom

        return False

    def iter_occurrences(
        self, task: Dict[str, Any], start_d: _dt.date, end_d: _dt.date
    ) -> Iterator[Tuple[_dt.date, _dt.date, bool]]:
        """Yield (orig_date, display_date, is_done) scanning a buffer window."""

        end_on = _parse_date_safe(task.get("end_on", ""))
        if end_on and end_d > end_on:
            end_d = end_on

        comp = set(task.get("completed") or [])
        scan_start = start_d - _dt.timedelta(days=3)
        scan_end   = end_d   + _dt.timedelta(days=3)
        d = scan_start
        while d <= scan_end:
            if self.occurs_on(task, d):
                disp = display_date_for(task, d)
                yield d, disp, (d.isoformat() in comp)
            d += _dt.timedelta(days=1)

    # ---------- state flips ----------
    def toggle_done_for_date(self, task, date_obj):
        comp = task.setdefault("completed", [])
        s_actual = date_obj.isoformat()
        s_disp   = display_date_for(task, date_obj).isoformat()
        was_done = (s_actual in comp) or (s_disp in comp)
        if was_done:
            if s_actual in comp: comp.remove(s_actual)
            if s_disp   in comp: comp.remove(s_disp)
        else:
            if s_actual not in comp: comp.append(s_actual)
            if s_disp != s_actual and s_disp not in comp: comp.append(s_disp)
        self.save()
        
    def toggle_cancel_for_date(self, task, date_obj):
        canc = task.setdefault("cancelled", [])
        comp = task.setdefault("completed", [])
        s_actual = date_obj.isoformat()
        s_disp   = display_date_for(task, date_obj).isoformat()

        was_cancelled = (s_actual in canc) or (s_disp in canc)
        if was_cancelled:
            if s_actual in canc: canc.remove(s_actual)
            if s_disp   in canc: canc.remove(s_disp)
        else:
            # When cancelling, ensure it's not marked done
            if s_actual in comp: comp.remove(s_actual)
            if s_disp   in comp: comp.remove(s_disp)
            if s_actual not in canc: canc.append(s_actual)
            if s_disp != s_actual and s_disp not in canc: canc.append(s_disp)
        self.save()

    def set_state_for_date(self, task, date_obj, state: str):
        """
        Enforce a specific state ("todo" | "done" | "cancel") for one occurrence.
        """
        comp = task.setdefault("completed", [])
        canc = task.setdefault("cancelled", [])
        s_actual = date_obj.isoformat()
        s_disp   = display_date_for(task, date_obj).isoformat()

        def _remove(bag):
            if s_actual in bag: bag.remove(s_actual)
            if s_disp   in bag: bag.remove(s_disp)

        if state == "todo":
            _remove(comp); _remove(canc)
        elif state == "done":
            _remove(canc); _remove(comp)
            if s_actual not in comp: comp.append(s_actual)
            if s_disp != s_actual and s_disp not in comp: comp.append(s_disp)
        elif state == "cancel":
            _remove(comp); _remove(canc)
            if s_actual not in canc: canc.append(s_actual)
            if s_disp != s_actual and s_disp not in canc: canc.append(s_disp)
        else:
            raise ValueError(f"Unknown state: {state}")
        self.save()

    def compute_pause_start_date(
        self,
        task: Dict[str, Any],
        today: _dt.date | None = None,
    ) -> _dt.date | None:
        """
        Decide where a recurring task's pause should start.

        Rule:
        - Look at all DONE and CANCELLED dates (actual dates, not display dates).
        - Find the most recent such date <= today (the "last closed" occurrence).
        - Pause from the *next* occurrence after that date.
        - If there is no closed occurrence, pause from the earliest occurrence
          on/after the task's start (or today if start date is unknown).

        Returns a date (pause_from) or None if we cannot find any reasonable
        occurrence (in which case callers can fall back to today).
        """
        today = today or _dt.date.today()
        rec = (task.get("recurrence") or {})
        freq = (rec.get("freq") or "one-off").lower()

        # Only meaningful for recurring tasks
        if freq == "one-off":
            return None

        # Collect closed dates (done or cancelled)
        comp = [ _parse_date_safe(s) for s in (task.get("completed") or []) ]
        canc = [ _parse_date_safe(s) for s in (task.get("cancelled") or []) ]
        closed = [d for d in (comp + canc) if d is not None and d <= today]

        last_closed = max(closed) if closed else None

        if last_closed:
            # Example: 11/1 done, 11/8 cancelled → last_closed = 11/8
            base = last_closed
        else:
            # No closed occurrences: use start_on if available, otherwise today
            base = _parse_date_safe(task.get("start_on", "")) or today

        # If we *do* have a last_closed, we want the NEXT occurrence after it.
        # If not, we want the first occurrence on/after base.
        start_scan = base + _dt.timedelta(days=1) if last_closed else base

        # Scan forward for a reasonable horizon to find the next occurrence.
        # 3 years is overkill but safe for all normal recurrences.
        horizon = today + _dt.timedelta(days=365 * 3)
        d = start_scan
        while d <= horizon:
            if self.occurs_on(task, d):
                return d
            d += _dt.timedelta(days=1)

        # Fallback: none found
        return None

    def batch_pause_for_companies(
        self,
        company_idxs: list[int],
        company_names: list[str],
    ) -> int:
        """
        Pause all *active* recurring tasks for the given companies.

        - One-off tasks are ignored.
        - Tasks with any end_on set ("stopped") are NOT changed.
        - Already-paused tasks are NOT changed.
        - pause_from is computed from the most recent done/cancelled occurrence,
          as per compute_pause_start_date().
        """
        if not company_idxs and not company_names:
            return 0

        today = _dt.date.today()
        idx_set = set(company_idxs)
        name_cf_set = { (n or "").strip().casefold() for n in company_names if n }

        changed = 0

        for t in self.tasks:
            rec = (t.get("recurrence") or {})
            freq = (rec.get("freq") or "one-off").lower()
            if freq == "one-off":
                continue

            # Skip stopped tasks (any end_on means "stopped")
            if t.get("end_on"):
                continue

            # Skip disabled tasks
            if not t.get("is_enabled", True):
                continue

            # Skip if already paused
            if t.get("is_paused"):
                continue

            # Check company match by index or name
            idx = t.get("company_idx")
            cname_cf = (t.get("company_name") or "").strip().casefold()
            if (idx not in idx_set) and (cname_cf not in name_cf_set):
                continue

            pause_start = self.compute_pause_start_date(t, today=today) or today
            t["pause_from"] = pause_start.isoformat()
            # New pause period wipes any older resume marker
            t["resume_from"] = ""
            t["is_paused"] = True
            changed += 1

        if changed:
            self.save()
        return changed

    def batch_resume_for_companies(
        self,
        company_idxs: list[int],
        company_names: list[str],
    ) -> int:
        """
        Resume all paused recurring tasks for the given companies.

        - One-off tasks are ignored.
        - Tasks with any end_on set ("stopped") are NOT changed.
        - Only tasks currently is_paused=True are touched.
        - resume_from is set to today; occurrences between pause_from and
          resume_from remain permanently skipped (per occurs_on()).
        """
        if not company_idxs and not company_names:
            return 0

        today = _dt.date.today()
        idx_set = set(company_idxs)
        name_cf_set = { (n or "").strip().casefold() for n in company_names if n }

        changed = 0

        for t in self.tasks:
            rec = (t.get("recurrence") or {})
            freq = (rec.get("freq") or "one-off").lower()
            if freq == "one-off":
                continue

            # Skip stopped tasks
            if t.get("end_on"):
                continue

            if not t.get("is_paused"):
                continue

            idx = t.get("company_idx")
            cname_cf = (t.get("company_name") or "").strip().casefold()
            if (idx not in idx_set) and (cname_cf not in name_cf_set):
                continue

            t["resume_from"] = today.isoformat()
            t["is_paused"] = False
            changed += 1

        if changed:
            self.save()
        return changed
