
from __future__ import annotations

import re
import json
import os
import hashlib
import ssl
import urllib.request
import urllib.parse
from pathlib import Path
from typing import Dict, Any, Optional, List
import datetime as dt

_PHONE_DIGITS_RE = re.compile(r"\d")
PHONE_DIGITS_RE = _PHONE_DIGITS_RE  # Alias for backward compatibility
_AM_WS_RE = re.compile(r"\s+")

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
    Preserves id if present (for entity links).
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
    
    # Preserve id if present (for entity links) - this is the primary field now
    # Check id first, then other_id, then linked_client_id for backward compatibility
    link_id = str(x.get("id") or x.get("other_id") or x.get("linked_client_id") or "").strip()
    if link_id:
        o["id"] = link_id
    
    # Also preserve other_id for backward compatibility (if different from id)
    other_id = str(x.get("other_id") or "").strip()
    if other_id and other_id != link_id:
        o["other_id"] = other_id
    other_label = str(x.get("other_label") or "").strip()
    if other_label:
        o["other_label"] = other_label
    
    if o["name"] and not (o["first_name"] or o["middle_name"] or o["last_name"] or o["nickname"]):
        o["first_name"] = o["name"]
    composed = compose_person_name(o["first_name"], o["middle_name"], o["last_name"], o["nickname"])
    if composed:
        o["name"] = composed
    return o

def display_relation_name(o: Dict[str, str]) -> str:
    o = ensure_relation_dict(o)
    return o.get("name","").strip()


# Normalization functions
def normalize_phone_digits(s: str) -> str:
    """Extract phone digits, return last 10 digits."""
    digits = "".join(PHONE_DIGITS_RE.findall(s or ""))
    return digits[-10:] if len(digits) >= 10 else digits


def normalize_ein_digits(s: str) -> str:
    """Extract EIN digits, return last 9 digits."""
    return "".join(PHONE_DIGITS_RE.findall(s or ""))[-9:]


def normalize_ssn_digits(s: str) -> str:
    """Extract SSN digits, return last 9 digits."""
    return "".join(PHONE_DIGITS_RE.findall(s or ""))[-9:]


def normalize_logs(logs):
    """Normalize log entries to standard format."""
    out = []
    for x in logs or []:
        if isinstance(x, dict):
            out.append({
                "ts":   str(x.get("ts","")).strip(),
                "user": str(x.get("user","")).strip(),
                "text": str(x.get("text","")).strip(),
                "done": bool(x.get("done", False)),
            })
        else:
            out.append({"ts":"", "user":"", "text":str(x), "done": False})
    return out


# Text processing
def tokenize(s: str) -> List[str]:
    """Tokenize string into words (lowercase, alphanumeric + @._-&)."""
    if s is None:
        return []
    s = str(s).lower().strip()
    parts = re.split(r"[^a-z0-9@._\-&]+", s)
    return [p for p in parts if p]


def norm_text(s: str) -> str:
    """Normalize text by tokenizing and rejoining."""
    return " ".join(tokenize(s))


# Relations helpers
def relations_to_display_lines(relations: List[Dict[str,str]]) -> List[str]:
    """Convert relations list to display name strings."""
    return [display_relation_name(o) for o in (relations or []) if display_relation_name(o)]


def parse_emails_from_field(email_field: str) -> List[str]:
    """Parse comma-, semicolon-, and newline-separated email string into list of trimmed, non-empty emails."""
    if not email_field:
        return []
    s = str(email_field).strip()
    out = []
    for part in re.split(r"[,;\n]", s):
        e = part.strip()
        if e:
            out.append(e)
    return out


def email_display_string(email_field: str) -> str:
    """Format email field for display: multiple emails shown one per line (\\n)."""
    emails = parse_emails_from_field(email_field)
    return "\n".join(emails) if emails else ""


def relations_to_flat_emails(relations: List[Dict[str,str]]) -> List[str]:
    """Extract unique emails from relations list. Treats each relation's email as comma/semicolon-separated (multiple)."""
    seen, out = set(), []
    for o in relations or []:
        raw = str(ensure_relation_dict(o).get("email","")).strip()
        for e in parse_emails_from_field(raw):
            if e and e not in seen:
                seen.add(e)
                out.append(e)
    return out


def relations_to_flat_phones(relations: List[Dict[str,str]]) -> List[str]:
    """Extract unique phones from relations list."""
    seen, out = set(), []
    for o in relations or []:
        p = str(ensure_relation_dict(o).get("phone","")).strip()
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


def is_valid_person_payload(data):
    """Check if data is a valid person payload tuple/list."""
    return isinstance(data, (tuple, list)) and len(data) == 3 and data[0] is not None


# Account manager helpers
def _account_manager_key(am: dict) -> str:
    """Stable, case-insensitive key for deduping account managers."""
    if not isinstance(am, dict):
        am = {"name": str(am)}
    name  = _AM_WS_RE.sub(" ", str(am.get("name", "") or "").strip()).casefold()
    email = str(am.get("email", "") or "").strip().casefold()
    phone = normalize_phone_digits(str(am.get("phone", "") or "").strip())
    return f"{name}|{email}|{phone}"


def _account_manager_id_from_key(key: str) -> str:
    """Deterministic short id from key (so imports don't create duplicates)."""
    h = hashlib.sha1((key or "").encode("utf-8", errors="ignore")).hexdigest()
    return f"am_{h[:12]}"


# Date/Quarter helpers
def today_date() -> dt.date:
    """Get today's date."""
    return dt.date.today()


def quarter_start(d: dt.date) -> dt.date:
    """Get the start date of the quarter containing d."""
    q = (d.month - 1) // 3
    first_month = q * 3 + 1
    return dt.date(d.year, first_month, 1)


def new_quarter_started(last_checked_iso: str | None) -> bool:
    """Check if a new quarter has started since last_checked_iso."""
    try:
        if not last_checked_iso:
            return True
        last = dt.date.fromisoformat(last_checked_iso)
    except Exception:
        return True
    start_now = quarter_start(today_date())
    return last < start_now


def safe_fetch_sales_tax_rate(state: str, city: str) -> Optional[float]:
    """
    Optional API call to fetch a sales tax rate by state/city.
    Configure env var TAX_API_URL like:
      https://example.com/tax?state={state}&city={city}
    Returns None on failure/offline.
    """
    state = (state or "").strip()
    city  = (city or "").strip()
    if not state:
        return None
    tmpl = os.environ.get("TAX_API_URL", "")
    if not tmpl:
        return None
    url = tmpl.format(state=state, city=urllib.parse.quote(city))
    try:
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(url, timeout=5, context=ctx) as resp:
            if resp.status != 200:
                return None
            data = resp.read().decode("utf-8", errors="ignore")
            m = re.search(r"(\d+(?:\.\d+)?)", data)
            if not m:
                return None
            val = float(m.group(1))
            if val <= 0.2:  # treat as fraction
                val *= 100.0
            return round(val, 4)
    except Exception:
        return None

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


def _client_matches_uid(client: Dict[str, Any], uid: str) -> bool:
    """
    Return True if this client is the one referred to by uid.
    uid can be: get_client_uid(client), or "ein:<9>", "ssn:<9>", "client:<id>", or raw id.
    Ensures linking works when clients have id stored as UUID but link dialog uses ein:/ssn:.
    """
    if not isinstance(client, dict) or not uid:
        return False
    uid = uid.strip()
    # Exact match with canonical get_client_uid
    if get_client_uid(client) == uid:
        return True
    # ein:<digits> -> match by EIN
    if uid.lower().startswith("ein:"):
        want = normalize_ein_digits(uid.split(":", 1)[1])
        if want and normalize_ein_digits(client.get("ein", "")) == want:
            return True
    # ssn:<digits> -> match by SSN (or EIN for individuals that use EIN field)
    if uid.lower().startswith("ssn:"):
        want = normalize_ssn_digits(uid.split(":", 1)[1])
        have = normalize_ssn_digits(client.get("ssn", "") or client.get("ein", ""))
        if want and have == want:
            return True
    # client:<id> or raw id
    raw_id = str(client.get("id") or "").strip()
    if uid.lower().startswith("client:"):
        want = (uid.split(":", 1)[1] or "").strip()
        return want and raw_id == want
    return raw_id == uid


def find_client_by_uid(clients: List[Dict[str, Any]], uid: str) -> Optional[Dict[str, Any]]:
    """
    Find a client by uid in any supported form: get_client_uid(c), ein:<9>, ssn:<9>, client:<id>, or raw id.
    This makes business-business and spouse-spouse (and all) linking work even when client['id']
    is stored as a UUID while the link dialog passes ein:/ssn: from candidates.
    """
    uid = (uid or "").strip()
    if not uid:
        return None
    for c in clients:
        if _client_matches_uid(c, uid):
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
      { "id": "ein:#########", "role": "...", ...all other fields... }
    Uses "id" instead of "other_id", and does not save "other_label" (name is saved instead).
    Preserves all fields from the input dict, ensuring id and role are set.
    """
    if not isinstance(x, dict):
        x = {"id": str(x).strip()}

    # Extract link-specific fields - use "id" instead of "other_id"
    link_id = str(x.get("id") or x.get("other_id") or x.get("linked_client_id") or "").strip()
    role = (str(x.get("role") or "") or "").strip().lower()

    # Start with all fields from the input dict
    result = dict(x)
    
    # Ensure link fields are set correctly - use "id" instead of "other_id"
    result["id"] = link_id
    result["role"] = role
    
    # Remove deprecated fields
    result.pop("other_id", None)
    result.pop("other_label", None)
    result.pop("linked_client_id", None)
    result.pop("linked_client_label", None)
    
    return result


def merge_relations(existing: List[Dict[str, Any]], incoming: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Merge/dedupe relations by id. Prefer non-empty fields from incoming.
    Preserves all relation fields (name, first_name, last_name, email, phone, address, etc.)
    Uses "id" instead of "other_id".
    IMPORTANT: Also includes relations without id (person relations, not entity links).
    """
    out: Dict[str, Dict[str, Any]] = {}
    relations_without_id: List[Dict[str, Any]] = []  # Store relations without id separately

    # First, add all existing relations (including those without id)
    for r in (existing or []):
        # Use ensure_relation_link which now preserves all fields
        rr = ensure_relation_link(r)
        link_id = rr.get("id") or ""
        
        if link_id:
            # Has id - store with all fields
            out[link_id] = dict(rr)
        else:
            # No id - treat as person relation (not entity link)
            # Store separately to preserve them
            relations_without_id.append(dict(rr))

    # Then, merge incoming relations (prefer incoming data)
    for r in (incoming or []):
        # Use ensure_relation_link which now preserves all fields
        rr = ensure_relation_link(r)
        link_id = rr.get("id") or ""
        
        if not link_id:
            # Incoming relation without id - add to relations_without_id if not duplicate
            # Check for duplicates by comparing key fields
            is_duplicate = False
            for existing_no_id in relations_without_id:
                if (existing_no_id.get("first_name") == rr.get("first_name") and
                    existing_no_id.get("last_name") == rr.get("last_name") and
                    existing_no_id.get("email") == rr.get("email")):
                    is_duplicate = True
                    break
            if not is_duplicate:
                relations_without_id.append(dict(rr))
            continue
        
        # If existing, prefer incoming non-empty fields (especially data fields)
        if link_id in out:
            cur = out[link_id]
            for key, value in rr.items():
                # Always update role from incoming
                if key == "role":
                    cur[key] = value
                # For data fields, prefer incoming if it's non-empty
                elif value and (not cur.get(key) or key in ("name", "first_name", "last_name", "email", "phone", "addr1", "addr2", "city", "state", "zip", "dob")):
                    cur[key] = value
            out[link_id] = cur
        else:
            out[link_id] = dict(rr)

    # Return relations with id first, then relations without id
    return list(out.values()) + relations_without_id


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

def _build_full_relation_from_client(src_client: dict, src_id: str, role_value: str) -> dict:
    """
    Build a full relation record from a client with all data fields.
    Returns a relation dict with other_id, other_label, role, and all client data.
    """
    src_label = str(src_client.get("name") or "").strip() or src_id
    is_ind = bool(src_client.get("is_individual")) or ((src_client.get("entity_type") or "").strip().casefold() == "individual")
    
    # Get contact info
    email = str(src_client.get("email") or "").strip()
    phone = str(src_client.get("phone") or "").strip()
    
    # Get name fields
    first_name = str(src_client.get("first_name") or "").strip()
    middle_name = str(src_client.get("middle_name") or "").strip()
    last_name = str(src_client.get("last_name") or "").strip()
    nickname = str(src_client.get("nickname") or "").strip()
    
    # For individuals, try to split name if first/last not available
    if is_ind and not first_name and not last_name:
        name_parts = [p for p in re.split(r"\s+", src_label) if p]
        if len(name_parts) >= 2:
            first_name = name_parts[0]
            last_name = " ".join(name_parts[1:])
        elif len(name_parts) == 1:
            first_name = name_parts[0]
    
    # Build full relation record
    rel_record = {
        "name": src_label,
        "first_name": first_name,
        "middle_name": middle_name,
        "last_name": last_name,
        "nickname": nickname,
        "email": email,
        "phone": phone,
        "addr1": str(src_client.get("addr1") or "").strip(),
        "addr2": str(src_client.get("addr2") or "").strip(),
        "city": str(src_client.get("city") or "").strip(),
        "state": str(src_client.get("state") or "").strip(),
        "zip": str(src_client.get("zip") or "").strip(),
        "dob": str(src_client.get("dob") or "").strip() if isinstance(src_client.get("dob"), str) else "",
        "role": (role_value or "").strip().lower(),
        "id": src_id,  # Use "id" instead of "other_id"
    }
    
    return ensure_relation_link(rel_record)


def _inverse_role(role: str) -> str:
    """Return the role for the reverse direction of a relation (e.g. parent <-> child)."""
    r = (role or "").strip().lower()
    if r == "child":
        return "parent"
    if r == "parent":
        return "child"
    if r == "owner":
        return "business"
    if r == "business":
        return "owner"
    return r


def sync_inverse_relations(clients: List[Dict[str, Any]]) -> int:
    """
    Ensure every client has back-links for relations pointing to them.
    If entity D has Chris Lim in its relations, Chris Lim will get D in its relations.
    Mutates clients in place. Returns the number of clients that were updated.
    """
    if not clients:
        return 0
    # Build id -> client index so we can resolve relation ids
    id_to_client: Dict[str, Dict[str, Any]] = {}
    for c in clients:
        uid = get_client_uid(c)
        if uid:
            id_to_client[uid] = c
    # Also match by id field directly for relation link resolution
    for c in clients:
        raw_id = str(c.get("id") or "").strip()
        if raw_id and raw_id not in id_to_client:
            id_to_client[raw_id] = c

    updated = 0
    for c in clients:
        c_id = get_client_uid(c)
        if not c_id:
            continue
        c_rels = c.get("relations", []) or []
        c_rels_by_id = {str(ensure_relation_link(r).get("id") or "").strip(): r for r in c_rels if ensure_relation_link(r).get("id")}

        for other in clients:
            if other is c:
                continue
            other_id = get_client_uid(other)
            if not other_id:
                continue
            other_rels = other.get("relations", []) or []
            for rel in other_rels:
                rr = ensure_relation_link(rel)
                rel_id = str(rr.get("id") or "").strip()
                if not rel_id or rel_id != c_id:
                    continue
                # other points to c; ensure c has a relation back to other
                if c_rels_by_id.get(other_id):
                    continue
                forward_role = (rr.get("role") or "").strip().lower()
                back_role = _inverse_role(forward_role)
                back_rel = _build_full_relation_from_client(other, other_id, back_role)
                c["relations"] = merge_relations(c.get("relations", []) or [], [back_rel])
                c_rels_by_id[other_id] = back_rel
                updated += 1
                break

    return updated


def link_clients_relations(app, this_id: str, other_id: str, link: bool = True, role: str = "", other_label: str = "", this_label: str = ""):
    """
    Symmetric link/unlink using client['relations'] (NOT officers).
    Expects app.items to be a list[dict] of clients and app.save_clients_data() exists (optional).
    Creates full relation records with all client data fields.
    """
    print(f"[helpers][LINK] link_clients_relations: this_id='{this_id}', other_id='{other_id}', link={link}, role='{role}'")
    this_id = (this_id or "").strip()
    other_id = (other_id or "").strip()
    if not this_id or not other_id:
        print(f"[helpers][LINK] link_clients_relations: Missing IDs - this_id='{this_id}', other_id='{other_id}'")
        return

    items = getattr(app, "items", []) or []
    print(f"[helpers][LINK] link_clients_relations: Found {len(items)} items")
    a = find_client_by_uid(items, this_id)
    b = find_client_by_uid(items, other_id)
    print(f"[helpers][LINK] link_clients_relations: Client A: {a is not None}, Client B: {b is not None}")
    if not isinstance(a, dict) or not isinstance(b, dict):
        print(f"[helpers][LINK] link_clients_relations: One or both clients not found or invalid")
        return

    # default labels if not provided
    if not other_label:
        other_label = str(b.get("name") or "").strip()
    if not this_label:
        this_label = str(a.get("name") or "").strip()

    a_rels = a.get("relations", []) or []
    b_rels = b.get("relations", []) or []

    if link:
        # Determine roles based on client types and provided role
        a_is_ind = bool(a.get("is_individual")) or ((a.get("entity_type") or "").strip().casefold() == "individual")
        b_is_ind = bool(b.get("is_individual")) or ((b.get("entity_type") or "").strip().casefold() == "individual")
        
        role_lower = (role or "").strip().lower()
        
        # Determine roles based on relationship type
        if not a_is_ind and b_is_ind:
            # Business → Individual
            # A (business) sees B (individual) with role: business owner, employee, or officer
            # B (individual) sees A (business) with role: business
            if role_lower in ("business owner", "businessowner", "employee", "officer"):
                role_a_to_b = role_lower
                role_b_to_a = "business"
            else:
                # Default to business owner if invalid role
                role_a_to_b = "business owner"
                role_b_to_a = "business"
        elif a_is_ind and not b_is_ind:
            # Individual → Business
            # A (individual) sees B (business) with role: business
            # B (business) sees A (individual) with role: owner
            role_a_to_b = "business"
            role_b_to_a = "owner"
        elif a_is_ind and b_is_ind:
            # Individual → Individual
            # Handle bidirectional roles: spouse, parent/child, relative
            if role_lower == "parent":
                role_a_to_b = "parent"
                role_b_to_a = "child"
            elif role_lower == "child":
                role_a_to_b = "child"
                role_b_to_a = "parent"
            elif role_lower in ("spouse", "relative"):
                # Symmetric roles
                role_a_to_b = role_lower
                role_b_to_a = role_lower
            else:
                # Default to relative if invalid
                role_a_to_b = "relative"
                role_b_to_a = "relative"
        else:
            # Business → Business (both are businesses)
            role_a_to_b = role_lower or "business"
            role_b_to_a = role_lower or "business"
        
        # Build full relation records with all client data
        print(f"[helpers][LINK] link_clients_relations: Building relations - role_a_to_b='{role_a_to_b}', role_b_to_a='{role_b_to_a}'")
        a_new = [_build_full_relation_from_client(b, other_id, role_a_to_b)]
        b_new = [_build_full_relation_from_client(a, this_id, role_b_to_a)]
        print(f"[helpers][LINK] link_clients_relations: a_new: {a_new}")
        print(f"[helpers][LINK] link_clients_relations: b_new: {b_new}")
        a["relations"] = merge_relations(a_rels, a_new)
        b["relations"] = merge_relations(b_rels, b_new)
        print(f"[helpers][LINK] link_clients_relations: After merge - a relations: {len(a.get('relations', []))}, b relations: {len(b.get('relations', []))}")
        print(f"[helpers][LINK] link_clients_relations: a relations data: {a.get('relations', [])}")
        print(f"[helpers][LINK] link_clients_relations: b relations data: {b.get('relations', [])}")
    else:
        print(f"[helpers][LINK] link_clients_relations: Unlinking this_id='{this_id}' from other_id='{other_id}'")
        def _relation_points_to_client(rel_id: str, target_uid: str) -> bool:
            """True if the relation id refers to the same client as target_uid (handles ein:/ssn:/raw formats)."""
            if not (rel_id and target_uid):
                return False
            if rel_id.strip() == target_uid.strip():
                return True
            client_rel = find_client_by_uid(items, rel_id)
            client_tgt = find_client_by_uid(items, target_uid)
            return client_rel is not None and client_tgt is not None and client_rel is client_tgt

        def _drop(rels, oid):
            out = []
            print(f"[helpers][LINK] link_clients_relations: _drop: Processing {len(rels or [])} relations, removing oid='{oid}'")
            for i, r in enumerate(rels or []):
                rr = ensure_relation_link(r)
                rel_id = (rr.get("id") or "").strip()
                match = _relation_points_to_client(rel_id, oid)
                print(f"[helpers][LINK] link_clients_relations: _drop: Relation {i} - id='{rel_id}', oid='{oid}', match={match}")
                if not match:
                    out.append(rr)
                    print(f"[helpers][LINK] link_clients_relations: _drop: Keeping relation {i}")
                else:
                    print(f"[helpers][LINK] link_clients_relations: _drop: Removing relation {i}")
            print(f"[helpers][LINK] link_clients_relations: _drop: Returning {len(out)} relations (removed {len(rels or []) - len(out)})")
            return out

        a_relations_before = len(a_rels)
        b_relations_before = len(b_rels)
        a["relations"] = _drop(a_rels, other_id)
        b["relations"] = _drop(b_rels, this_id)
        a_relations_after = len(a.get("relations", []))
        b_relations_after = len(b.get("relations", []))
        print(f"[helpers][LINK] link_clients_relations: After unlink - a relations: {a_relations_before} -> {a_relations_after}, b relations: {b_relations_before} -> {b_relations_after}")

    # Persist if available
    if hasattr(app, "save_clients_data"):
        try:
            print(f"[helpers][LINK] link_clients_relations: Calling save_clients_data")
            app.save_clients_data()
            print(f"[helpers][LINK] link_clients_relations: save_clients_data completed")
        except Exception as e:
            print(f"[helpers][LINK] link_clients_relations: Error saving: {e}")
            import traceback
            traceback.print_exc()


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