from __future__ import annotations

import datetime
import uuid
from typing import Any

# ---- Status / category constants ----

FILE_REQUEST_STATUSES = [
    "Needed",
    "Requested",
    "Received",
    "Waiting on Client",
    "Waiting on Third Party",
    "Waiting on Government",
    "Not Needed",
    "Archived",
]

REMINDER_STATUSES = [
    "Active",
    "Done This Year",
    "Inactive",
    "Archived",
]

ISSUE_STATUSES = [
    "Open",
    "Waiting on Client",
    "Waiting on Government",
    "Resolved",
    "Closed",
    "Archived",
]

CATEGORIES = [
    "Income Tax",
    "Business Tax",
    "Bookkeeping",
    "Payroll",
    "Sales Tax",
    "Government Notice",
    "IPP",
    "Retirement",
    "Entity",
    "Other",
]

ISSUE_TYPES = [
    "Government Notice",
    "Tax Return Issue",
    "File Problem",
    "Client Habit",
    "IPP",
    "Payroll",
    "Bookkeeping",
    "Other",
]

CLOSED_FILE_STATUSES = frozenset({"Received", "Not Needed", "Archived"})
OPEN_ISSUE_STATUSES = frozenset({"Open", "Waiting on Client", "Waiting on Government"})

FILE_SOURCES = ["Client", "Third Party", "Government", "Internal", "Other"]
PRIORITIES = ["Low", "Normal", "High", "Urgent"]

# Common file templates: (display label, default category)
FILE_TEMPLATES: list[tuple[str, str]] = [
    ("— Custom —", ""),
    ("W-2", "Income Tax"),
    ("1099-NEC", "Income Tax"),
    ("1099-MISC", "Income Tax"),
    ("1099-INT", "Income Tax"),
    ("1099-DIV", "Income Tax"),
    ("1098 (Mortgage)", "Income Tax"),
    ("K-1", "Income Tax"),
    ("Schedule C P&L", "Income Tax"),
    ("Charitable donations", "Income Tax"),
    ("Property tax bill", "Income Tax"),
    ("Driver license / ID", "Income Tax"),
    ("Bank statements", "Bookkeeping"),
    ("Credit card statements", "Bookkeeping"),
    ("Payroll reports", "Payroll"),
    ("W-3 / 941", "Payroll"),
    ("Sales tax return", "Sales Tax"),
    ("IPP valuation / actuarial", "IPP"),
    ("Form 5500", "Retirement"),
    ("Government notice", "Government Notice"),
    ("Prior-year tax return", "Income Tax"),
]

FILE_TEMPLATE_LABELS = [t[0] for t in FILE_TEMPLATES]


def default_tax_year() -> str:
    return str(datetime.date.today().year)


def file_status_tag(status: str) -> str:
    st = str(status or "").strip()
    return {
        "Needed": "tr_needed",
        "Requested": "tr_requested",
        "Received": "tr_received",
        "Waiting on Client": "tr_waiting",
        "Waiting on Third Party": "tr_waiting",
        "Waiting on Government": "tr_waiting",
        "Not Needed": "tr_closed",
        "Archived": "tr_archived",
    }.get(st, "tr_default")


def reminder_status_tag(status: str) -> str:
    st = str(status or "").strip()
    return {
        "Active": "tr_active",
        "Done This Year": "tr_received",
        "Inactive": "tr_closed",
        "Archived": "tr_archived",
    }.get(st, "tr_default")


def issue_status_tag(status: str) -> str:
    st = str(status or "").strip()
    return {
        "Open": "tr_open",
        "Waiting on Client": "tr_waiting",
        "Waiting on Government": "tr_waiting",
        "Resolved": "tr_received",
        "Closed": "tr_closed",
        "Archived": "tr_archived",
    }.get(st, "tr_default")


def now_ts() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


def today_str() -> str:
    return datetime.date.today().isoformat()


def new_id() -> str:
    return str(uuid.uuid4())


def ensure_tracker_fields(client: dict) -> None:
    if not isinstance(client, dict):
        return
    client.setdefault("file_requests", [])
    client.setdefault("annual_reminders", [])
    client.setdefault("client_issues", [])


def _find_item(items: list, item_id: str) -> dict | None:
    iid = str(item_id or "").strip()
    if not iid:
        return None
    for it in items:
        if isinstance(it, dict) and str(it.get("id", "") or "").strip() == iid:
            return it
    return None


def _is_ipp_file_open(fr: dict) -> bool:
    if not isinstance(fr, dict) or fr.get("archived"):
        return False
    if (fr.get("category") or "").strip() != "IPP":
        return False
    return fr.get("status") not in CLOSED_FILE_STATUSES


def _is_ipp_issue_open(iss: dict) -> bool:
    if not isinstance(iss, dict) or iss.get("archived"):
        return False
    if iss.get("status") not in OPEN_ISSUE_STATUSES:
        return False
    cat = (iss.get("category") or "").strip()
    typ = (iss.get("type") or "").strip()
    return cat == "IPP" or typ == "IPP"


def count_tracker_summary(client: dict) -> dict[str, int]:
    """Counts for the Tracker tab summary row and Profile summary."""
    ensure_tracker_fields(client)
    missing = 0
    requested = 0
    received = 0
    open_issues = 0
    ipp_open = 0
    active_reminders = 0

    for fr in client.get("file_requests") or []:
        if not isinstance(fr, dict) or fr.get("archived"):
            continue
        st = str(fr.get("status", "") or "")
        if st == "Needed":
            missing += 1
        elif st == "Requested":
            requested += 1
        elif st == "Received":
            received += 1
        if _is_ipp_file_open(fr):
            ipp_open += 1

    for rem in client.get("annual_reminders") or []:
        if not isinstance(rem, dict) or rem.get("archived"):
            continue
        st = str(rem.get("status", "") or "")
        if st == "Active":
            active_reminders += 1
            if (rem.get("category") or "").strip() == "IPP":
                ipp_open += 1

    for iss in client.get("client_issues") or []:
        if not isinstance(iss, dict) or iss.get("archived"):
            continue
        st = str(iss.get("status", "") or "")
        if st in OPEN_ISSUE_STATUSES:
            open_issues += 1
            if _is_ipp_issue_open(iss):
                ipp_open += 1

    return {
        "missing": missing,
        "requested": requested,
        "received": received,
        "open_issues": open_issues,
        "ipp_open": ipp_open,
        "active_reminders": active_reminders,
    }


def count_open_tracker_items(client: dict) -> tuple[int, int, int]:
    """Return (open_files, open_reminders, open_issues) for dashboard."""
    ensure_tracker_fields(client)
    files = 0
    reminders = 0
    issues = 0

    for fr in client.get("file_requests") or []:
        if not isinstance(fr, dict) or fr.get("archived"):
            continue
        if str(fr.get("status", "") or "") not in CLOSED_FILE_STATUSES:
            files += 1

    for rem in client.get("annual_reminders") or []:
        if not isinstance(rem, dict) or rem.get("archived"):
            continue
        if str(rem.get("status", "") or "") == "Active":
            reminders += 1

    for iss in client.get("client_issues") or []:
        if not isinstance(iss, dict) or iss.get("archived"):
            continue
        if str(iss.get("status", "") or "") in OPEN_ISSUE_STATUSES:
            issues += 1

    return files, reminders, issues


def append_tracker_log(client: dict, text: str) -> None:
    ensure_tracker_fields(client)
    client.setdefault("logs", []).append({
        "ts": now_ts(),
        "text": str(text or "").strip(),
        "done": False,
        "log_type": "memo",
    })


def _base_file_request(data: dict | None = None) -> dict[str, Any]:
    d = data or {}
    ts = now_ts()
    return {
        "id": str(d.get("id") or new_id()),
        "tax_year": str(d.get("tax_year", "") or ""),
        "category": str(d.get("category", "Income Tax") or "Income Tax"),
        "name": str(d.get("name", "") or ""),
        "status": str(d.get("status", "Needed") or "Needed"),
        "requested_date": str(d.get("requested_date", "") or ""),
        "received_date": str(d.get("received_date", "") or ""),
        "due_date": str(d.get("due_date", "") or ""),
        "repeat_next_year": bool(d.get("repeat_next_year", True)),
        "source": str(d.get("source", "Client") or "Client"),
        "priority": str(d.get("priority", "Normal") or "Normal"),
        "note": str(d.get("note", "") or ""),
        "created_ts": str(d.get("created_ts") or ts),
        "updated_ts": str(d.get("updated_ts") or ts),
        "archived": bool(d.get("archived", False)),
    }


def _base_annual_reminder(data: dict | None = None) -> dict[str, Any]:
    d = data or {}
    ts = now_ts()
    return {
        "id": str(d.get("id") or new_id()),
        "title": str(d.get("title", "") or ""),
        "category": str(d.get("category", "Other") or "Other"),
        "status": str(d.get("status", "Active") or "Active"),
        "priority": str(d.get("priority", "Normal") or "Normal"),
        "applies_every_year": bool(d.get("applies_every_year", True)),
        "tax_year": str(d.get("tax_year", "") or ""),
        "note": str(d.get("note", "") or ""),
        "created_ts": str(d.get("created_ts") or ts),
        "updated_ts": str(d.get("updated_ts") or ts),
        "archived": bool(d.get("archived", False)),
    }


def _base_client_issue(data: dict | None = None) -> dict[str, Any]:
    d = data or {}
    ts = now_ts()
    return {
        "id": str(d.get("id") or new_id()),
        "type": str(d.get("type", "Other") or "Other"),
        "category": str(d.get("category", "Other") or "Other"),
        "tax_year": str(d.get("tax_year", "") or ""),
        "title": str(d.get("title", "") or ""),
        "status": str(d.get("status", "Open") or "Open"),
        "priority": str(d.get("priority", "Normal") or "Normal"),
        "opened_date": str(d.get("opened_date", "") or ""),
        "closed_date": str(d.get("closed_date", "") or ""),
        "note": str(d.get("note", "") or ""),
        "related_file_request_id": str(d.get("related_file_request_id", "") or ""),
        "created_ts": str(d.get("created_ts") or ts),
        "updated_ts": str(d.get("updated_ts") or ts),
        "archived": bool(d.get("archived", False)),
    }


def add_file_request(client: dict, data: dict) -> dict:
    ensure_tracker_fields(client)
    item = _base_file_request(data)
    client["file_requests"].append(item)
    return item


def update_file_request(client: dict, item_id: str, patch: dict) -> dict | None:
    ensure_tracker_fields(client)
    item = _find_item(client["file_requests"], item_id)
    if item is None:
        return None
    for k, v in (patch or {}).items():
        if k in ("created_ts", "id"):
            continue
        item[k] = v
    item["updated_ts"] = now_ts()
    return item


def add_annual_reminder(client: dict, data: dict) -> dict:
    ensure_tracker_fields(client)
    item = _base_annual_reminder(data)
    client["annual_reminders"].append(item)
    return item


def update_annual_reminder(client: dict, item_id: str, patch: dict) -> dict | None:
    ensure_tracker_fields(client)
    item = _find_item(client["annual_reminders"], item_id)
    if item is None:
        return None
    for k, v in (patch or {}).items():
        if k in ("created_ts", "id"):
            continue
        item[k] = v
    item["updated_ts"] = now_ts()
    return item


def add_client_issue(client: dict, data: dict) -> dict:
    ensure_tracker_fields(client)
    item = _base_client_issue(data)
    client["client_issues"].append(item)
    return item


def update_client_issue(client: dict, item_id: str, patch: dict) -> dict | None:
    ensure_tracker_fields(client)
    item = _find_item(client["client_issues"], item_id)
    if item is None:
        return None
    for k, v in (patch or {}).items():
        if k in ("created_ts", "id"):
            continue
        item[k] = v
    item["updated_ts"] = now_ts()
    return item


def roll_forward_file_requests(client: dict, from_year: str, to_year: str) -> int:
    ensure_tracker_fields(client)
    fy = str(from_year or "").strip()
    ty = str(to_year or "").strip()
    if not fy or not ty or fy == ty:
        return 0

    existing = {
        (
            str(x.get("tax_year", "") or "").strip(),
            str(x.get("category", "") or "").strip().casefold(),
            str(x.get("name", "") or "").strip().casefold(),
        )
        for x in client.get("file_requests") or []
        if isinstance(x, dict)
    }

    copied = 0
    for src in client.get("file_requests") or []:
        if not isinstance(src, dict):
            continue
        if src.get("archived"):
            continue
        if str(src.get("tax_year", "") or "").strip() != fy:
            continue
        if not src.get("repeat_next_year"):
            continue

        key = (ty, str(src.get("category", "") or "").strip().casefold(), str(src.get("name", "") or "").strip().casefold())
        if key in existing:
            continue

        note = str(src.get("note", "") or "")
        prefix = f"Rolled forward from {fy}. "
        if note and not note.startswith(prefix):
            note = prefix + note
        elif not note:
            note = prefix.strip()

        new_item = _base_file_request({
            "tax_year": ty,
            "category": src.get("category", "Income Tax"),
            "name": src.get("name", ""),
            "status": "Needed",
            "requested_date": "",
            "received_date": "",
            "due_date": src.get("due_date", ""),
            "repeat_next_year": src.get("repeat_next_year", True),
            "source": src.get("source", "Client"),
            "priority": src.get("priority", "Normal"),
            "note": note,
            "archived": False,
        })
        client["file_requests"].append(new_item)
        existing.add(key)
        copied += 1

    if copied:
        append_tracker_log(client, f"Rolled forward {copied} file request(s) from {fy} to {ty}.")
    return copied


def collect_tax_years(client: dict) -> list[str]:
    """All tax years referenced in tracker data, plus current/previous year."""
    ensure_tracker_fields(client)
    years: set[str] = set()
    today = datetime.date.today()
    years.add(str(today.year))
    years.add(str(today.year - 1))

    for coll in ("file_requests", "annual_reminders", "client_issues"):
        for it in client.get(coll) or []:
            if not isinstance(it, dict):
                continue
            y = str(it.get("tax_year", "") or "").strip()
            if y:
                years.add(y)

    def _sort_key(s: str):
        try:
            return (0, -int(s))
        except ValueError:
            return (1, s.casefold())

    return sorted(years, key=_sort_key)
