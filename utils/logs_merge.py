"""
Unified log + work-item rows for client Logs tab and global Notes page.
"""
from __future__ import annotations

import datetime


def norm_minute(ts: str) -> str:
    s = str(ts or "").strip().replace("T", " ").replace("Z", "")
    return s[:16]


def parse_sort_ts(raw: str) -> datetime.datetime:
    s = str(raw or "").strip()
    if not s:
        return datetime.datetime.min
    s_iso = s.replace("Z", "+00:00")
    try:
        return datetime.datetime.fromisoformat(s_iso)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s[:19], fmt)
        except ValueError:
            pass
    return datetime.datetime.min


def split_history_text(text: str) -> tuple[str, str]:
    t = (text or "").strip()
    if " | " in t:
        a, b = t.split(" | ", 1)
        return a.strip(), b.strip()
    return t, ""


def build_merged_rows(client: dict) -> list[dict]:
    """
    Per-client merged rows. Each dict:
      status, task, time_disp, note_disp, sort_ts, tag, meta
    meta keys: kind in memo|history|work|active; log_index?, work_item_id?
    """
    rows: list[dict] = []

    completed_keys: set[tuple[str, str]] = set()
    for wi in client.get("work_items") or []:
        if not isinstance(wi, dict):
            continue
        if str(wi.get("status", "") or "").strip().lower() != "completed":
            continue
        tn = (wi.get("task_name") or "").strip()
        ca = wi.get("completed_at") or wi.get("updated_at") or ""
        if tn:
            completed_keys.add((tn.casefold(), norm_minute(str(ca))))

    aw = client.get("active_work") or {}
    active_wid = ""
    if isinstance(aw, dict) and aw:
        active_wid = str(aw.get("work_item_id", "") or "").strip()
        tname = str(aw.get("task_name") or "").strip() or "—"
        since = str(aw.get("started_at") or aw.get("created_at") or "").strip() or "—"
        rows.append({
            "status": "Active",
            "task": tname,
            "time_disp": since,
            "note_disp": "",
            "sort_ts": parse_sort_ts(since),
            "tag": "task",
            "meta": {"kind": "active", "work_item_id": active_wid},
        })

    for wi in client.get("work_items") or []:
        if not isinstance(wi, dict):
            continue
        st = str(wi.get("status", "") or "").strip().lower()
        wid = str(wi.get("id", "") or "").strip()
        tname = str(wi.get("task_name", "") or "").strip()
        if not tname:
            tname = "—"

        if st == "active":
            if active_wid and wid == active_wid:
                continue
            ts = str(wi.get("started_at") or wi.get("updated_at") or wi.get("created_at") or "").strip() or "—"
            rows.append({
                "status": "Active",
                "task": tname,
                "time_disp": ts,
                "note_disp": str(wi.get("note") or "").replace("\n", " ").strip(),
                "sort_ts": parse_sort_ts(ts),
                "tag": "task",
                "meta": {"kind": "work", "work_item_id": wid},
            })
        elif st == "on_hold":
            ts = str(wi.get("held_at") or wi.get("updated_at") or "").strip() or "—"
            note = str(wi.get("note") or "").replace("\n", " ").strip()
            rows.append({
                "status": "On hold",
                "task": tname,
                "time_disp": ts,
                "note_disp": note,
                "sort_ts": parse_sort_ts(ts),
                "tag": "task",
                "meta": {"kind": "work", "work_item_id": wid},
            })
        elif st == "completed":
            ts = str(wi.get("completed_at") or wi.get("updated_at") or "").strip() or "—"
            note = str(wi.get("note") or "").replace("\n", " ").strip()
            rows.append({
                "status": "Finished",
                "task": tname,
                "time_disp": ts,
                "note_disp": note,
                "sort_ts": parse_sort_ts(ts),
                "tag": "finished",
                "meta": {"kind": "work", "work_item_id": wid},
            })

    for i, entry in enumerate(client.get("logs") or []):
        if not isinstance(entry, dict):
            continue
        lt = str(entry.get("log_type", "") or "").strip().lower()
        if not lt:
            lt = "memo"
        ts_raw = entry.get("ts", "")
        ts_disp = str(ts_raw or "").strip()
        if entry.get("edited"):
            ts_disp = f"{ts_disp} (Edited)" if ts_disp else "(Edited)"
        text = str(entry.get("text", "") or "").strip()
        sort_ts = parse_sort_ts(str(ts_raw))

        if lt == "history":
            task_part, note_part = split_history_text(text)
            key = (task_part.casefold(), norm_minute(str(ts_raw)))
            if key in completed_keys:
                continue
            rows.append({
                "status": "Finished",
                "task": task_part or "—",
                "time_disp": ts_disp,
                "note_disp": note_part,
                "sort_ts": sort_ts,
                "tag": "finished",
                "meta": {"kind": "history", "log_index": i},
            })
        else:
            done = bool(entry.get("done"))
            st_label = "N/A"
            if done:
                st_label = "N/A ✓"
            rows.append({
                "status": st_label,
                "task": text or "—",
                "time_disp": ts_disp,
                "note_disp": "",
                "sort_ts": sort_ts,
                "tag": "done" if done else "active",
                "meta": {"kind": "memo", "log_index": i},
            })

    rows.sort(key=lambda r: r["sort_ts"], reverse=True)
    return rows


def build_all_clients_merged_rows(items: list) -> list[dict]:
    """Flatten merge across clients; each row includes client_idx + client_name."""
    out: list[dict] = []
    for i, c in enumerate(items or []):
        if not isinstance(c, dict):
            continue
        cname = (c.get("name") or "").strip() or "(unnamed)"
        for r in build_merged_rows(c):
            row = dict(r)
            row["client_idx"] = i
            row["client_name"] = cname
            out.append(row)
    out.sort(key=lambda r: r["sort_ts"], reverse=True)
    return out
