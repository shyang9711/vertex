
from __future__ import annotations

import re
import json
from pathlib import Path
from typing import Dict, Any, Optional, List

_PHONE_DIGITS_RE = re.compile(r"\d")

def compose_person_name(first: str, middle: str, last: str, nickname: str) -> str:
    parts = [first.strip(), middle.strip(), last.strip()]
    base = " ".join([p for p in parts if p])
    if nickname.strip():
        return f'{base} ("{nickname.strip()}")' if base else nickname.strip()
    return base

def ensure_relation_dict(x) -> Dict[str, str]:
    """
    Normalize personnel dict.
    If only 'name' exists, move it to first_name (migration) and compose display 'name'.
    """
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
        "role":                (str(x.get("role","")) or "officer").strip().lower() or "officer",
        "linked_client_id":    str(x.get("linked_client_id","") or "").strip(),
        "linked_client_label": str(x.get("linked_client_label","") or "").strip(),
    }
    if o["name"] and not (o["first_name"] or o["middle_name"] or o["last_name"] or o["nickname"]):
        o["first_name"] = o["name"]
    composed = compose_person_name(o["first_name"], o["middle_name"], o["last_name"], o["nickname"])
    if composed:
        o["name"] = composed
    return o

def display_relation_name(o: Dict[str, str]) -> str:
    o = ensure_relation_dict(o)
    return o.get("name","").strip()

def _digits(s: str) -> str:
    return re.sub(r"\D+", "", s or "")

def get_client_uid(client: Dict[str, Any]) -> str:
    # Prefer explicit id if you already store it
    cid = (client.get("id") or "").strip()
    if cid:
        return cid

    if client.get("is_individual"):
        ssn = _digits(client.get("ssn", "")) or _digits(client.get("ein", ""))
        return f"ssn:{ssn}" if ssn else ""
    else:
        ein = _digits(client.get("ein", ""))
        return f"ein:{ein}" if ein else ""

def find_client_by_uid(clients: List[Dict[str, Any]], uid: str) -> Optional[Dict[str, Any]]:
    uid = (uid or "").strip()
    if not uid:
        return None
    for c in clients:
        if get_client_uid(c) == uid:
            return c
    return None

def _normalize_phone(x: str) -> str:
    return "".join(ch for ch in (x or "") if ch.isdigit())

def _safe_lower(x: str) -> str:
    return (x or "").strip().lower()

def _client_ref_for(app, client_idx: int) -> str:
    items = getattr(app, "items", []) or []
    if 0 <= client_idx < len(items):
        c = items[client_idx]
        cid = str(c.get("id") or "").strip()
        if cid:
            # IMPORTANT: in your app id is already like "ein:#########" or "ssn:#########"
            return cid
        # fallback if id missing
        return get_client_uid(c)
    return ""

def ensure_relation_link(x) -> Dict[str, str]:
    """
    Normalize a relation link dict.
    Canonical shape:
      { "other_id": "ein:#########", "other_label": "...", "role": "..." }
    """
    if not isinstance(x, dict):
        x = {"other_id": str(x).strip()}

    other_id = str(x.get("other_id") or x.get("id") or x.get("linked_client_id") or "").strip()
    other_label = str(x.get("other_label") or x.get("label") or x.get("linked_client_label") or "").strip()
    role = (str(x.get("role") or "") or "").strip().lower()

    return {
        "other_id": other_id,
        "other_label": other_label,
        "role": role,
    }


def merge_relations(existing: List[Dict[str, Any]], incoming: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Merge/dedupe relations by other_id. Prefer non-empty label/role from incoming.
    """
    out: Dict[str, Dict[str, str]] = {}

    for r in (existing or []):
        rr = ensure_relation_link(r)
        oid = rr["other_id"]
        if not oid:
            continue
        out[oid] = rr

    for r in (incoming or []):
        rr = ensure_relation_link(r)
        oid = rr["other_id"]
        if not oid:
            continue

        cur = out.get(oid) or {"other_id": oid, "other_label": "", "role": ""}
        # Prefer incoming label/role if provided
        if rr.get("other_label"):
            cur["other_label"] = rr["other_label"]
        if rr.get("role"):
            cur["role"] = rr["role"]
        out[oid] = cur

    # stable order
    return list(out.values())


def migrate_officer_business_links_to_relations(client: Dict[str, Any]) -> Dict[str, Any]:
    """
    Legacy migration:
    - If client.officers contains rows that represent linked *entities* (role == "business" AND linked_client_id),
      move/merge them into client.relations.
    - Return the updated client dict (modifies in place).
    """
    if not isinstance(client, dict):
        return client

    officers = client.get("officers", []) or []
    relations = client.get("relations", []) or []

    rel_add: List[Dict[str, Any]] = []
    cleaned_officers: List[Any] = []

    for o in officers:
        if not isinstance(o, dict):
            cleaned_officers.append(o)
            continue

        role = (str(o.get("role") or "") or "").strip().lower()
        lid = str(o.get("linked_client_id") or "").strip()
        lab = str(o.get("linked_client_label") or "").strip()

        # "business link row" heuristic: role business + has linked_client_id
        if role == "business" and lid:
            rel_add.append({"other_id": lid, "other_label": lab, "role": role})
            # drop it from officers (so officers remains real people)
            continue

        cleaned_officers.append(o)

    client["officers"] = cleaned_officers
    client["relations"] = merge_relations(relations, rel_add)
    return client

def _migrations_file(data_root: Path) -> Path:
    return Path(data_root) / "migrations.json"

def load_migration_flags(data_root: Path) -> Dict[str, Any]:
    p = _migrations_file(data_root)
    try:
        if p.exists():
            with p.open("r", encoding="utf-8") as f:
                d = json.load(f)
                return d if isinstance(d, dict) else {}
    except Exception:
        pass
    return {}

def save_migration_flags(data_root: Path, flags: Dict[str, Any]) -> None:
    p = _migrations_file(data_root)
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        with p.open("w", encoding="utf-8") as f:
            json.dump(flags or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        # migrations must never crash the app
        pass

def is_migration_done(data_root: Path, key: str) -> bool:
    flags = load_migration_flags(data_root)
    val = flags.get(key)
    if isinstance(val, dict):
        return bool(val.get("done"))
    return bool(val)

def mark_migration_done(data_root: Path, key: str, meta: Dict[str, Any] | None = None) -> None:
    flags = load_migration_flags(data_root)
    flags[key] = {"done": True, **(meta or {})}
    save_migration_flags(data_root, flags)

def link_clients_relations(app, this_id: str, other_id: str, link: bool = True, role: str = "", other_label: str = "", this_label: str = ""):
    """
    Symmetric link/unlink using client['relations'] (NOT officers).
    Expects app.items to be a list[dict] of clients and app.save_clients_data() exists (optional).
    """
    this_id = (this_id or "").strip()
    other_id = (other_id or "").strip()
    if not this_id or not other_id:
        return

    items = getattr(app, "items", []) or []
    a = find_client_by_uid(items, this_id)
    b = find_client_by_uid(items, other_id)
    if not isinstance(a, dict) or not isinstance(b, dict):
        return

    # default labels if not provided
    if not other_label:
        other_label = str(b.get("name") or "").strip()
    if not this_label:
        this_label = str(a.get("name") or "").strip()

    a_rels = a.get("relations", []) or []
    b_rels = b.get("relations", []) or []

    if link:
        a_new = [{"other_id": other_id, "other_label": other_label, "role": (role or "").strip().lower()}]
        b_new = [{"other_id": this_id, "other_label": this_label, "role": (role or "").strip().lower()}]
        a["relations"] = merge_relations(a_rels, a_new)
        b["relations"] = merge_relations(b_rels, b_new)
    else:
        def _drop(rels, oid):
            out = []
            for r in (rels or []):
                rr = ensure_relation_link(r)
                if rr.get("other_id") and rr["other_id"] != oid:
                    out.append(rr)
            return out

        a["relations"] = _drop(a_rels, other_id)
        b["relations"] = _drop(b_rels, this_id)

    # Persist if available
    if hasattr(app, "save_clients_data"):
        try:
            app.save_clients_data()
        except Exception:
            pass


def _linked_id_to_client_idx(app, link_id: str):
    link_id = (link_id or "").strip()
    if not link_id or ":" not in link_id:
        return None
    kind, val = link_id.split(":", 1)
    kind = (kind or "").strip().lower()
    val = (val or "").strip()

    items = getattr(app, "items", []) or []

    # direct id match (if link_id is exactly stored id)
    for i, c in enumerate(items):
        if str(c.get("id") or "").strip() == link_id:
            return i

    if kind == "idx":
        try:
            i = int(val)
            return i if 0 <= i < len(items) else None
        except Exception:
            return None

    if kind == "client":
        for i, c in enumerate(items):
            if str(c.get("id", "") or "").strip() == val:
                return i
        return None

    if kind in ("ein", "ssn"):
        target = "".join(ch for ch in val if ch.isdigit())[:9]
        if not target:
            return None

        for i, c in enumerate(items):
            if kind == "ein":
                got = "".join(ch for ch in (c.get("ein","") or "") if ch.isdigit())[:9]
            else:
                got = "".join(ch for ch in (c.get("ssn","") or "") if ch.isdigit())[:9]
                if not got:
                    # (optional) if some individuals store SSN in ein field
                    got = "".join(ch for ch in (c.get("ein","") or "") if ch.isdigit())[:9]

            if got == target:
                return i
        return None

    return None


def _find_matching_person_index(client_dict: dict, src_person: dict):
    """
    Returns (role_key, index) in linkee client where the person matches.
    Match order: email, phone, then first+last.
    """
    src = ensure_relation_dict(src_person)

    src_email = _safe_lower(src.get("email"))
    src_phone = _normalize_phone(src.get("phone"))
    src_first = (src.get("first_name") or "").strip()
    src_last  = (src.get("last_name") or "").strip()

    for rk in ("officers", "employees", "spouses"):
        arr = client_dict.get(rk, []) or []
        for j, p in enumerate(arr):
            if not isinstance(p, dict):
                continue

            p2 = ensure_relation_dict(p)

            if src_email and _safe_lower(p2.get("email")) == src_email:
                return rk, j
            if src_phone and _normalize_phone(p2.get("phone")) == src_phone:
                return rk, j
            if src_first and src_last and (
                (p2.get("first_name","").strip() == src_first) and
                (p2.get("last_name","").strip() == src_last)
            ):
                return rk, j

    return (None, None)


def apply_reciprocal_link(app, linker_client_idx: int, linker_role_key: str, linker_person_idx: int):
    """
    Reads linker person's linked_client_id, finds that linkee client,
    then writes reciprocal linked_client_id onto the matching person in linkee.
    """
    items = getattr(app, "items", []) or []
    if not (0 <= linker_client_idx < len(items)):
        return

    linker_client = items[linker_client_idx]
    arr = linker_client.get(linker_role_key, []) or []
    if not (0 <= linker_person_idx < len(arr)):
        return

    linker_person = arr[linker_person_idx]
    if not isinstance(linker_person, dict):
        return

    link_id = (linker_person.get("linked_client_id") or "").strip()
    if not link_id:
        return

    linkee_idx = _linked_id_to_client_idx(app, link_id)
    if linkee_idx is None or not (0 <= linkee_idx < len(items)):
        return

    linkee_client = items[linkee_idx]
    rk2, pidx2 = _find_matching_person_index(linkee_client, linker_person)
    if rk2 is None:
        # If your “auto-bidirectional” expects a person to exist on linkee side,
        # then this is why you "don't see" the link back.
        return

    backref = _client_ref_for(app, linker_client_idx)
    if not backref:
        return

    linkee_arr = linkee_client.get(rk2, []) or []
    linkee_person = linkee_arr[pidx2]
    if isinstance(linkee_person, dict):
        linkee_person["linked_client_id"] = backref

def normalize_client_schema(client: Dict[str, Any]) -> Dict[str, Any]:
    """
    One-stop schema normalization for a client dict:
    - Ensure relations exists and is normalized.
    - Migrate legacy officer 'business link rows' into relations.
    - Keep officers as real people only.
    """
    if not isinstance(client, dict):
        return client

    # Ensure list fields exist
    if not isinstance(client.get("officers"), list):
        client["officers"] = []
    if not isinstance(client.get("relations"), list):
        client["relations"] = []

    # Normalize existing relations entries
    client["relations"] = [ensure_relation_link(r) for r in (client.get("relations") or []) if isinstance(r, dict)]

    # Migrate legacy officer->relations
    migrate_officer_business_links_to_relations(client)

    # Final normalize/dedupe
    client["relations"] = merge_relations(client.get("relations") or [], [])
    return client


def merge_client_update(existing: Dict[str, Any], incoming: Dict[str, Any]) -> Dict[str, Any]:
    """
    Merge incoming client payload into existing client dict, *and* merge officers->relations migration.

    Key behavior:
    - officers: merged (append) and normalized as people (does NOT keep legacy business-link rows)
    - relations: merged/deduped by other_id
    - also migrates any legacy business-link rows found in either officers list into relations
    """
    if not isinstance(existing, dict):
        existing = {}
    if not isinstance(incoming, dict):
        incoming = {}

    # Start from existing, overwrite simple fields from incoming
    out = dict(existing)
    for k, v in incoming.items():
        if k in ("officers", "relations"):
            continue
        out[k] = v

    # Merge officers (people)
    ex_off = existing.get("officers", []) or []
    in_off = incoming.get("officers", []) or []
    merged_officers_raw = []
    for x in (ex_off + in_off):
        if isinstance(x, dict):
            merged_officers_raw.append(ensure_relation_dict(x))
        else:
            merged_officers_raw.append(x)

    out["officers"] = merged_officers_raw

    # Merge relations (entity links)
    ex_rel = existing.get("relations", []) or []
    in_rel = incoming.get("relations", []) or []
    out["relations"] = merge_relations(ex_rel, in_rel)

    # IMPORTANT: migrate any legacy officer business-link rows into relations (and remove them from officers)
    normalize_client_schema(out)

    return out