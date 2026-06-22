# de9c_to_csv.py
import re
import csv
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# =========================
# Config
# =========================
# this is the *placeholder* a PDF row will get if it doesn't have a real SSN
DEFAULT_SSN = "000000000"
DEFAULT_WAGE_PLAN_CODE = "S"

STATE_CALIFORNIA = "California"
STATE_NEW_YORK = "New York"
STATE_HAWAII = "Hawaii"
STATE_OPTIONS = (STATE_CALIFORNIA, STATE_NEW_YORK, STATE_HAWAII)

STRICT_SSN_LINE_RE = re.compile(r"^\d{9}$")
MONEY_LINE_RE = re.compile(r"^\d+(?:,\d{3})*\.\d{2}$")
COMMA_NAME_LINE_RE = re.compile(r"^[A-Za-z][A-Za-z\s'\-]+,\s*[A-Za-z].*$")
CITY_STATE_ZIP_COMMA_RE = re.compile(r"^.+,\s*[A-Z]{2}\s*\d", re.IGNORECASE)

# match CA address lines like "BRENTWOOD CA 94513"
ADDRESS_LINE_RE = re.compile(
    r"^[A-Z0-9 .,'/-]+ CA \d{5}(?:-\d{4})?$"
)

# text that should never become an employee
FORBIDDEN_NAME_PHRASES = [
    # Generic non-human text / headings / boilerplate
    "california quarterly contribution", "return and report of wages",
    "quarterly contribution", "report of wages", "employer account",
    "employees full-time", "employees full time", "payroll period",
    "yr qtr", "quarter ended", "due", "delinquent if",
    "page number", "page", "instructions", "or received by", "not postmarked",
    "keep for your records", "do not mail", "rev", "copy", "continuation",
    # Totals / column headings
    "f. total subject wages", "g. pit wages", "h. pit withheld",
    "i. total subject wages this page", "j. total pit wages this page", "k. total pit withheld this page",
    "l. grand total subject wages", "m. grand total pit wages", "n. grand total pit withheld",
    # Addresses / companies (generic)
    "address", "po box", "zip", "california", "inc", "llc", "corp", "company", "co.", "ltd",
    # Hawaii UC-B6 / UC-B6A
    "state of hawaii", "form uc-b6", "unemployment insurance", "federal id number",
    "account number", "for quarter ending", "delinquent after", "do not write",
    "total wages this page", "add to form uc-b6", "employee's ssa",
    # New York NYS-45
    "state of new york", "nys-45", "quarterly combined withholding",
    "wage reporting", "unemployment insurance return", "employer registration",
]

# column/header variants we've actually seen in your PDFs
COLHEAD_KEYWORDS = [
    "d. social security number",
    "social security number",
    "employee name",
    "last name",
    "first name",
    "(m.i.)",
    "(mi)",
    "m.i.",
    "m i",
]

STOP_KEYWORDS_NAMEISH = {
    "total", "subject", "wages", "pit", "withheld", "page", "grand",
    "return", "report", "quarterly", "employer", "account", "number",
    "keep", "mail", "rev", "copy", "continuation", "employees", "payroll",
    "period", "mo", "due", "instructions", "california", "zip", "address",
    "received", "postmarked", "inc", "llc", "corp", "company", "co.", "ltd"
}

CSV_HEADERS = [
    "SSN",
    "First Name",
    "Middle Initial",
    "Last Name",
    "Total Subject Wages",
    "Personal Income Tax Wages",
    "Personal Income Tax Withheld",
    "Wage Plan Code",
]

HEADER_NAME_PAT = re.compile(r"^\s*e\.\s*employee\s+name", re.IGNORECASE)
SSN_LINE_RE = re.compile(r"^\s*(\d{9})(?:\s+(.*))?$")

SUFFIX_SET = {"JR", "SR", "II", "III", "IV", "V", "JR.", "SR.", "II.", "III.", "IV.", "V."}

# =========================
# PDF extraction
# =========================
def extract_text_from_pdf(pdf_path: Path) -> str:
    try:
        import fitz
        parts = []
        with fitz.open(pdf_path) as doc:
            for page in doc:
                parts.append(page.get_text("text"))
        text = "\n".join(parts).replace("\r", "\n")
        text = re.sub(r"[ \t]+", " ", text)
        return text
    except Exception as e1:
        last_err = e1

    try:
        from pdfminer.high_level import extract_text as miner
        text = miner(str(pdf_path)) or ""
        text = re.sub(r"[ \t]+", " ", text.replace("\r", "\n"))
        if text.strip():
            return text
    except Exception as e2:
        last_err = e2

    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(str(pdf_path))
        parts = [pg.extract_text() or "" for pg in reader.pages]
        text = "\n".join(parts).replace("\r", "\n")
        text = re.sub(r"[ \t]+", " ", text)
        if text.strip():
            return text
    except Exception as e3:
        last_err = e3

    raise RuntimeError(f"Failed to extract text; install PyMuPDF or pdfminer.six or PyPDF2. Last error: {last_err}")

# =========================
# helpers
# =========================
def _clean_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", (s or "")).strip()

def _is_forbidden_name_line(s: str, *, allow_digits: bool = False) -> bool:
    s = s or ""
    s_l = s.lower()
    if ADDRESS_LINE_RE.match(s.strip()):
        return True
    if not allow_digits and re.search(r"\d", s_l):
        return True
    return any(ph in s_l for ph in FORBIDDEN_NAME_PHRASES)

def _is_de9c_colhead(s: str) -> bool:
    if not s:
        return False
    s_l = s.lower().strip()
    # parentheses-style column names
    if s_l.startswith("(") and s_l.endswith(")"):
        return True
    # d. something
    if re.match(r"^[a-z]\.\s", s_l):
        return True
    for kw in COLHEAD_KEYWORDS:
        if kw in s_l:
            return True
    return False

def _is_nameish(line: str) -> bool:
    if not line or _is_forbidden_name_line(line) or _is_de9c_colhead(line):
        return False
    low = line.lower()
    if _contains_stop_keyword(line):
        return False
    toks = re.findall(r"[A-Za-z][A-Za-z\-']*", line)
    return 1 <= len(toks) <= 6

def _money_triplet(line: str):
    m = re.match(
        r"^\s*(\d+(?:,\d{3})*(?:\.\d{2})?)\s+(\d+(?:,\d{3})*(?:\.\d{2})?)\s+(\d+(?:,\d{3})*(?:\.\d{2})?)\s*$",
        line or "",
    )
    return m.groups() if m else None

def _money_token(line: str):
    m = re.match(r"^\s*(\d+(?:,\d{3})*(?:\.\d{2})?)\s*$", line or "")
    return m.group(1) if m else None

def _money_to_float(s: str) -> float:
    try:
        return float((s or "").replace(",", ""))
    except Exception:
        return float("nan")

def _plausible(F: str, G: str, H: str) -> bool:
    f, g, h = _money_to_float(F), _money_to_float(G), _money_to_float(H)
    if any(v != v for v in (f, g, h)):
        return False
    if not (f >= g >= h >= 0):
        return False
    if f == g == h == 0:
        return False
    return True

def _looks_like_page_total(F: str, G: str, H: str) -> bool:
    # your PDF has big page totals right after headers
    if not (F and G and H):
        return False
    if F == G:
        try:
            val = float(F.replace(",", ""))
            return val > 20000  # tweakable
        except Exception:
            return False
    return False

def _scan_amounts(block: list[str], start_idx: int = 0):
    """
    EXACTLY the same shape as your current file: always returns 4 values.
    """
    t = start_idx
    while t < len(block):
        mg = _money_triplet(block[t])
        if mg and _plausible(*mg):
            return mg[0], mg[1], mg[2], t
        if t + 2 < len(block):
            a = _money_token(block[t])
            b = _money_token(block[t + 1])
            c = _money_token(block[t + 2])
            if a and b and c and _plausible(a, b, c):
                return a, b, c, t + 2
        t += 1
    return None, None, None, start_idx

def _titlecase_hyphenated(token: str) -> str:
    return "-".join(part[:1].upper() + part[1:].lower() for part in token.split("-"))

def _titlecase_name_preserve_hyphens(s: str) -> str:
    return " ".join(_titlecase_hyphenated(t) for t in s.split())

def _is_mi_token(tok: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z]\.?", tok or ""))

def _is_suffix(tok: str) -> bool:
    return tok.upper() in SUFFIX_SET

def _enforce_firstname_limit(name: str) -> str:
    name = (name or "").strip()
    if " " in name:
        first_part = name.split()[0]
        return first_part if len(first_part) <= 11 else first_part[:11]
    return name if len(name) <= 11 else name[:11]

def _enforce_lastname_limit(name: str) -> str:
    name = (name or "").strip()
    if len(name) <= 19:
        return name
    if " " in name:
        first_part = name.split()[0]
        return first_part if len(first_part) <= 19 else first_part[:19]
    return name[:19]

def _parse_comma_name(name_line: str) -> tuple[str, str, str]:
    """Parse 'LAST, FIRST [MIDDLE]' names used on HI UC-B6A and NY NYS-45."""
    name_line = _clean_spaces(name_line)
    if "," not in name_line:
        return _extract_name([name_line])

    last_part, rest = name_line.split(",", 1)
    last = _enforce_lastname_limit(_titlecase_name_preserve_hyphens(last_part.strip()))
    tokens = re.findall(r"[A-Za-z][A-Za-z\-']*", rest.strip())
    if not tokens:
        return "", "", last

    first_raw = tokens[0]
    mi = ""
    idx = 1
    if idx < len(tokens) and _is_mi_token(tokens[idx]):
        mi = tokens[idx][0].upper()
        idx += 1
    elif idx < len(tokens) and len(tokens[idx]) == 1:
        mi = tokens[idx].upper()
        idx += 1
    if idx < len(tokens):
        first_raw = " ".join([first_raw] + tokens[idx:])

    first = _enforce_firstname_limit(_titlecase_name_preserve_hyphens(first_raw))
    return first, mi, last

def _next_nonempty(lines: list[str], start: int) -> int:
    i = start
    while i < len(lines) and not lines[i]:
        i += 1
    return i

def _collect_money_lines(lines: list[str], start: int, *, max_count: int = 5) -> tuple[list[str], int]:
    amounts: list[str] = []
    i = start
    while i < len(lines) and len(amounts) < max_count:
        if not lines[i]:
            i += 1
            continue
        if MONEY_LINE_RE.match(lines[i]):
            amounts.append(lines[i])
            i += 1
            continue
        break
    return amounts, i

def _row_from_employee(ssn: str, first: str, mi: str, last: str, amounts: list[str], *, capture: str, name_line: str, block_idx: int):
    F = amounts[0] if len(amounts) >= 1 else ""
    G = amounts[1] if len(amounts) >= 2 else ""
    H = amounts[2] if len(amounts) >= 3 else ""
    row = {
        "SSN": ssn,
        "First Name": first,
        "Middle Initial": mi,
        "Last Name": last,
        "Total Subject Wages": F,
        "Personal Income Tax Wages": G,
        "Personal Income Tax Withheld": H,
        "Wage Plan Code": DEFAULT_WAGE_PLAN_CODE,
        "_dbg_name_lines": name_line,
    }
    dbg = {
        "block": block_idx,
        "capture": capture,
        "ssn": ssn,
        "name_lines": name_line,
        "first": first,
        "mi": mi,
        "last": last,
        "F": F,
        "G": G,
        "H": H,
    }
    return row, dbg

def parse_ssn_comma_name_amounts_text_with_debug(
    text: str,
    *,
    capture_label: str,
    min_amounts: int = 1,
    max_amounts: int = 5,
):
    """Parse Intuit-style state wage reports: SSN, LAST,FIRST name, then money column(s)."""
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    debug_rows = []
    block_idx = 0
    i = 0
    n = len(lines)

    while i < n:
        if not STRICT_SSN_LINE_RE.match(lines[i]):
            i += 1
            continue

        ssn = lines[i]
        j = _next_nonempty(lines, i + 1)
        if j >= n:
            break

        name_line = lines[j]
        if not _is_comma_person_name_line(name_line):
            i += 1
            continue

        k = _next_nonempty(lines, j + 1)
        if k >= n:
            break

        amounts, next_idx = _collect_money_lines(lines, k, max_count=max_amounts)
        if len(amounts) < min_amounts:
            i += 1
            continue

        first, mi, last = _parse_comma_name(name_line)
        if not (first and last):
            i += 1
            continue

        row, dbg = _row_from_employee(
            ssn, first, mi, last, amounts,
            capture=capture_label,
            name_line=name_line,
            block_idx=block_idx,
        )
        rows.append(row)
        debug_rows.append(dbg)
        block_idx += 1
        i = next_idx

    return rows, debug_rows

def parse_hawaii_ucb6a_text_with_debug(text: str):
    return parse_ssn_comma_name_amounts_text_with_debug(
        text,
        capture_label="hawaii-ucb6a",
        min_amounts=1,
        max_amounts=1,
    )

def _is_comma_person_name_line(s: str) -> bool:
    s = _clean_spaces(s)
    if not COMMA_NAME_LINE_RE.match(s) or _is_forbidden_name_line(s):
        return False
    if CITY_STATE_ZIP_COMMA_RE.match(s):
        return False
    _left, right = s.split(",", 1)
    right = right.strip()
    if re.fullmatch(r"[A-Za-z]{2}", right):
        return False
    return True

def _is_nys45_wage_type(s: str) -> bool:
    return (s or "").upper() in {"R", "O"}

def _is_single_letter_token(s: str) -> bool:
    return len((s or "").strip()) == 1 and s.isalpha()

def _contains_stop_keyword(s: str) -> bool:
    low = (s or "").lower()
    return any(re.search(rf"\b{re.escape(kw)}\b", low) for kw in STOP_KEYWORDS_NAMEISH)

def _looks_like_nys45_name_part(s: str) -> bool:
    s = _clean_spaces(s)
    if not s or STRICT_SSN_LINE_RE.match(s) or MONEY_LINE_RE.match(s):
        return False
    if _is_nys45_wage_type(s) or _is_forbidden_name_line(s):
        return False
    if _contains_stop_keyword(s):
        return False
    return bool(re.search(r"[A-Za-z]", s))

def parse_ny_nys45_ssn_names_text_with_debug(text: str):
    """Parse Intuit NYS-45 Part C PDFs: SSN, last name, first name, R/O, optional MI, wages."""
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    debug_rows = []
    block_idx = 0
    i = 0
    n = len(lines)

    while i < n:
        if not STRICT_SSN_LINE_RE.match(lines[i]):
            i += 1
            continue

        ssn = lines[i]
        j = _next_nonempty(lines, i + 1)
        if j >= n - 1:
            i += 1
            continue
        if not (_looks_like_nys45_name_part(lines[j]) and _looks_like_nys45_name_part(lines[j + 1])):
            i += 1
            continue

        last_raw = lines[j]
        first_raw = lines[j + 1]
        k = j + 2
        if k >= n or not _is_nys45_wage_type(lines[k]):
            i += 1
            continue
        k += 1

        mi = ""
        if k < n and _is_single_letter_token(lines[k]) and k + 1 < n and MONEY_LINE_RE.match(lines[k + 1]):
            mi = lines[k].upper()
            k += 1

        amounts, next_k = _collect_money_lines(lines, k, max_count=3)
        if len(amounts) < 2:
            i += 1
            continue

        f_val = _money_to_float(amounts[0])
        if not (f_val > 0 and f_val < 500000):
            i += 1
            continue

        first = _enforce_firstname_limit(_titlecase_name_preserve_hyphens(first_raw))
        last = _enforce_lastname_limit(_titlecase_name_preserve_hyphens(last_raw))
        if not (first and last):
            i += 1
            continue

        padded = (amounts + ["", ""])[:3]
        row, dbg = _row_from_employee(
            ssn,
            first,
            mi,
            last,
            padded,
            capture="ny-nys45-ssn",
            name_line=f"{last_raw}, {first_raw}",
            block_idx=block_idx,
        )
        rows.append(row)
        debug_rows.append(dbg)
        block_idx += 1
        i = next_k

    return rows, debug_rows

def _part_c_start_index(lines: list[str]) -> int:
    for idx, ln in enumerate(lines):
        low = ln.lower()
        if "part c" in low and ("wage reporting" in low or "employee/payee" in low):
            return idx
    return 0

def _collect_part_c_ssns(lines: list[str], start: int) -> list[str]:
    ssns: list[str] = []
    for ln in lines[start:]:
        if STRICT_SSN_LINE_RE.match(ln):
            ssns.append(ln)
    return ssns

def parse_ny_nys45_partc_text_with_debug(text: str):
    """Parse modern NYS-45 Part C PDFs where names and SSNs extract in separate columns."""
    lines = [ln.strip() for ln in text.split("\n")]
    start = _part_c_start_index(lines)
    employees: list[dict] = []
    i = start
    n = len(lines)

    while i < n - 2:
        if not (_looks_like_nys45_name_part(lines[i]) and _looks_like_nys45_name_part(lines[i + 1])):
            i += 1
            continue

        last_raw = lines[i]
        first_raw = lines[i + 1]
        j = i + 2
        mi = ""
        if j < n and _is_nys45_wage_type(lines[j]):
            j += 1
        if j < n and _is_single_letter_token(lines[j]) and j + 1 < n and MONEY_LINE_RE.match(lines[j + 1]):
            mi = lines[j].upper()
            j += 1

        amounts, next_j = _collect_money_lines(lines, j, max_count=5)
        if len(amounts) < 3:
            i += 1
            continue

        f_val = _money_to_float(amounts[0])
        if not (f_val > 0 and f_val < 500000):
            i += 1
            continue

        first = _enforce_firstname_limit(_titlecase_name_preserve_hyphens(first_raw))
        last = _enforce_lastname_limit(_titlecase_name_preserve_hyphens(last_raw))
        if not (first and last):
            i += 1
            continue

        employees.append({
            "first": first,
            "mi": mi,
            "last": last,
            "amounts": amounts,
            "name_line": f"{last_raw}, {first_raw}",
        })
        i = next_j

    ssns = _collect_part_c_ssns(lines, start)
    rows = []
    debug_rows = []
    for idx, emp in enumerate(employees):
        ssn = ssns[idx] if idx < len(ssns) else DEFAULT_SSN
        row, dbg = _row_from_employee(
            ssn,
            emp["first"],
            emp["mi"],
            emp["last"],
            emp["amounts"],
            capture="ny-nys45-partc",
            name_line=emp["name_line"],
            block_idx=idx,
        )
        rows.append(row)
        debug_rows.append(dbg)

    return rows, debug_rows

def parse_ny_nys45_text_with_debug(text: str):
    rows, dbg = parse_ny_nys45_ssn_names_text_with_debug(text)
    if rows:
        return rows, dbg
    rows, dbg = parse_ssn_comma_name_amounts_text_with_debug(
        text,
        capture_label="ny-nys45-comma",
        min_amounts=3,
        max_amounts=3,
    )
    if rows:
        return rows, dbg
    rows, dbg = parse_ny_nys45_partc_text_with_debug(text)
    if rows:
        return rows, dbg
    return parse_de9c_text_with_debug(text)

def parse_payroll_text_with_debug(text: str, state: str):
    state = (state or STATE_CALIFORNIA).strip()
    if state == STATE_HAWAII:
        return parse_hawaii_ucb6a_text_with_debug(text)
    if state == STATE_NEW_YORK:
        return parse_ny_nys45_text_with_debug(text)
    return parse_de9c_text_with_debug(text)

def _extract_name(name_lines: list[str]) -> tuple[str, str, str]:
    if not name_lines:
        return "", "", ""
    tokens = re.findall(r"[A-Za-z][A-Za-z\-']*", " ".join(name_lines))
    if not tokens:
        return "", "", ""

    suffix = ""
    if tokens and _is_suffix(tokens[-1]):
        suffix = tokens.pop(-1).upper()

    mi = ""
    if len(tokens) >= 2 and _is_mi_token(tokens[1]):
        mi = tokens[1][0].upper()
        tokens = [tokens[0]] + tokens[2:]

    if len(tokens) == 1:
        first_raw = tokens[0]
        last_raw = ""
    else:
        first_raw = tokens[0]
        last_raw = " ".join(tokens[1:])

    first = _titlecase_name_preserve_hyphens(first_raw)
    last = _titlecase_name_preserve_hyphens(last_raw).strip()

    first = _enforce_firstname_limit(first)
    last = _enforce_lastname_limit(last)

    if suffix:
        last = (last + " " + suffix).strip()

    return first, mi, last

# =========================
# main parser
# =========================
def parse_de9c_text_with_debug(text: str):
    lines = [ln.strip() for ln in text.split("\n")]
    n = len(lines)
    i = 0
    rows = []
    debug_rows = []
    block_idx = 0

    seen_first_real_employee = False  # only let truly floating names after we saw one

    # NEW: counter to generate unique fake SSNs for "no-SSN" employees
    fake_ssn_counter = 0

    def _next_fake_ssn():
        return DEFAULT_SSN

    while i < n:
        line = lines[i]

        # CASE 0: header-like start (for all-no-SSN PDFs)
        if _is_de9c_colhead(line):
            j = i + 1
            # skip more header lines
            while j < n and _is_de9c_colhead(lines[j]):
                j += 1

            name_lines = []
            k = j
            while k < n:
                if _money_triplet(lines[k]) or _money_token(lines[k]):
                    break
                if _is_de9c_colhead(lines[k]) or _is_forbidden_name_line(lines[k]):
                    break
                name_lines.append(lines[k])
                k += 1

            F, G, H, amt_idx = _scan_amounts(lines, k)
            if name_lines and F is not None and not _looks_like_page_total(F, G, H):
                first, mi, last = _extract_name(name_lines)
                rows.append({
                    "SSN": _next_fake_ssn(),  # <-- instead of always DEFAULT_SSN
                    "First Name": first,
                    "Middle Initial": mi,
                    "Last Name": last,
                    "Total Subject Wages": F,
                    "Personal Income Tax Wages": G,
                    "Personal Income Tax Withheld": H,
                    "Wage Plan Code": DEFAULT_WAGE_PLAN_CODE,
                    "_dbg_name_lines": " | ".join(name_lines),
                })
                debug_rows.append({
                    "block": block_idx, "capture": "header-employee",
                    "ssn": rows[-1]["SSN"],
                    "name_lines": " | ".join(name_lines),
                    "first": first, "mi": mi, "last": last,
                    "F": F, "G": G, "H": H
                })
                seen_first_real_employee = True
                i = amt_idx + 1
                continue

            i += 1
            continue

        # CASE 1: normal SSN block
        m = SSN_LINE_RE.match(line)
        if m:
            ssn = m.group(1) or _next_fake_ssn()
            after = _clean_spaces(m.group(2) or "")

            j = i + 1
            while j < n and not SSN_LINE_RE.match(lines[j]):
                j += 1
            block = lines[i + 1: j]

            name_lines = []
            if after and not _is_forbidden_name_line(after, allow_digits=True) and not _is_de9c_colhead(after):
                name_lines.append(after)
            k = 0
            while k < min(5, len(block)):
                if _money_triplet(block[k]) or _money_token(block[k]):
                    break
                if _is_de9c_colhead(block[k]):
                    k += 1
                    continue
                if not _is_forbidden_name_line(block[k]) and block[k]:
                    name_lines.append(block[k])
                k += 1

            first, mi, last = _extract_name(name_lines)
            F, G, H, end_idx = _scan_amounts(block, k)

            rows.append({
                "SSN": ssn,
                "First Name": first,
                "Middle Initial": mi,
                "Last Name": last,
                "Total Subject Wages": F or "",
                "Personal Income Tax Wages": G or "",
                "Personal Income Tax Withheld": H or "",
                "Wage Plan Code": DEFAULT_WAGE_PLAN_CODE,
                "_dbg_name_lines": " | ".join(name_lines),
            })
            debug_rows.append({
                "block": block_idx, "capture": "ssn-first", "ssn": ssn,
                "name_lines": " | ".join(name_lines),
                "first": first, "mi": mi, "last": last,
                "F": F, "G": G, "H": H
            })
            seen_first_real_employee = True

            # CASE 1b: extra names inside this block
            p = end_idx
            while p < len(block):
                ln = block[p]
                if not ln:
                    p += 1
                    continue
                if _is_de9c_colhead(ln):
                    p += 1
                    continue

                if (ln.isupper() and not _is_forbidden_name_line(ln)) or _is_nameish(ln):
                    inner_names = [ln]
                    look = p + 1
                    while look < len(block):
                        nxt = block[look]
                        if _money_triplet(nxt) or _money_token(nxt):
                            break
                        if _is_forbidden_name_line(nxt) or _is_de9c_colhead(nxt):
                            break
                        if (nxt.isupper() and not _is_forbidden_name_line(nxt)) or _is_nameish(nxt):
                            inner_names.append(nxt)
                            look += 1
                            if len(inner_names) >= 3:
                                break
                            continue
                        break
                    F2, G2, H2, consumed2 = _scan_amounts(block, look)
                    if F2 is not None and not _looks_like_page_total(F2, G2, H2):
                        f2, mi2, l2 = _extract_name(inner_names)
                        rows.append({
                            "SSN": _next_fake_ssn(),
                            "First Name": f2,
                            "Middle Initial": mi2,
                            "Last Name": l2,
                            "Total Subject Wages": F2,
                            "Personal Income Tax Wages": G2,
                            "Personal Income Tax Withheld": H2,
                            "Wage Plan Code": DEFAULT_WAGE_PLAN_CODE,
                            "_dbg_name_lines": " | ".join(inner_names),
                        })
                        debug_rows.append({
                            "block": block_idx, "capture": "ssn-extra",
                            "ssn": rows[-1]["SSN"],
                            "name_lines": " | ".join(inner_names),
                            "first": f2, "mi": mi2, "last": l2,
                            "F": F2, "G": G2, "H": H2
                        })
                        p = consumed2 if consumed2 is not None else look
                        continue
                p += 1

            i = j
            block_idx += 1
            continue

        # CASE 2: floating employee (only after we have at least one real/parsed employee)
        if seen_first_real_employee:
            if ((line.isupper() and not _is_forbidden_name_line(line)) or _is_nameish(line)) and not _is_de9c_colhead(line):
                F, G, H, amt_idx = _scan_amounts(lines, i + 1)
                if F is not None and not _looks_like_page_total(F, G, H):
                    first, mi, last = _extract_name([line])
                    rows.append({
                        "SSN": _next_fake_ssn(),
                        "First Name": first,
                        "Middle Initial": mi,
                        "Last Name": last,
                        "Total Subject Wages": F,
                        "Personal Income Tax Wages": G,
                        "Personal Income Tax Withheld": H,
                        "Wage Plan Code": DEFAULT_WAGE_PLAN_CODE,
                        "_dbg_name_lines": line,
                    })
                    debug_rows.append({
                        "block": -1, "capture": "floating-extra",
                        "ssn": rows[-1]["SSN"],
                        "name_lines": line,
                        "first": first, "mi": mi, "last": last,
                        "F": F, "G": G, "H": H
                    })
                    i = amt_idx + 1
                    continue

        i += 1

    return rows, debug_rows

# =========================
# post-filter
# =========================
def _filter_out_header_rows(rows):
    """Remove header/non-employee artifacts and normalize missing SSNs.

    - Any row missing/invalid SSN gets DEFAULT_SSN.
    - Drops header/boilerplate rows and non-employee rows (prevents ',,,,,,,' lines).
    """
    out = []
    dropped_header = 0
    dropped_non_employee = 0

    for r in rows:
        src = (r.get("_dbg_name_lines") or "").strip()

        # Drop obvious header rows
        if HEADER_NAME_PAT.match(src):
            dropped_header += 1
            continue

        # Remove debug-only field before output
        r.pop("_dbg_name_lines", None)

        # Normalize SSN: if missing/invalid, force DEFAULT_SSN
        ssn_digits = re.sub(r"\D", "", (r.get("SSN") or ""))
        if len(ssn_digits) != 9:
            ssn_digits = DEFAULT_SSN
        r["SSN"] = ssn_digits

        # Require BOTH first+last name (prevents B./A. header fragments becoming "people")
        first_name = (r.get("First Name") or "").strip()
        last_name = (r.get("Last Name") or "").strip()
        has_name = bool(first_name and last_name)

        # Require at least one money value
        has_money = bool(
            (r.get("Total Subject Wages") or "").strip()
            or (r.get("Personal Income Tax Wages") or "").strip()
            or (r.get("Personal Income Tax Withheld") or "").strip()
        )

        # Extra guard against addresses/boilerplate accidentally being captured as a name line
        if src and _is_forbidden_name_line(src, allow_digits=True):
            dropped_non_employee += 1
            continue

        if not (has_name and has_money):
            dropped_non_employee += 1
            continue

        def _is_initial_token(s: str) -> bool:
            s = (s or "").strip()
            if not s:
                return False
            s2 = s.replace(".", "")
            return len(s2) <= 1 and s2.isalpha()
        
        if _is_initial_token(first_name) and _is_initial_token(last_name):
            dropped_non_employee += 1
            continue

        out.append(r)

    return out, dropped_header, dropped_non_employee


# =========================
# csv writer
# =========================
def write_csv_no_header(out_path: Path, rows):
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        for r in rows:
            w.writerow({k: (r.get(k, "") or "") for k in CSV_HEADERS})

# =========================
# GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DE9 / Payroll → CSV")
        self.geometry("800x500")
        self.resizable(True, True)

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        self.state_var = tk.StringVar(value=STATE_CALIFORNIA)
        ttk.Label(frm, text="State:").grid(row=0, column=0, sticky="w")
        state_cb = ttk.Combobox(
            frm,
            textvariable=self.state_var,
            values=STATE_OPTIONS,
            state="readonly",
            width=18,
        )
        state_cb.grid(row=0, column=1, sticky="w", pady=(0, 8))
        state_cb.bind("<<ComboboxSelected>>", self._on_state_change)

        self.path_var = tk.StringVar(value="")
        self.pdf_label = ttk.Label(frm, text=self._pdf_label_text())
        self.pdf_label.grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.path_var, width=70).grid(row=2, column=0, sticky="we", pady=(2, 8))
        ttk.Button(frm, text="Browse…", command=self.pick_pdf).grid(row=2, column=1, padx=(8, 0))
        ttk.Button(frm, text="Convert to CSV", command=self.convert).grid(row=2, column=2, padx=(8, 0))

        self.log = tk.Text(frm, height=24)
        self.log.grid(row=3, column=0, columnspan=3, sticky="nsew")
        frm.rowconfigure(3, weight=1)
        frm.columnconfigure(0, weight=1)

    def _pdf_label_text(self) -> str:
        state = self.state_var.get()
        if state == STATE_HAWAII:
            return "Hawaii UC-B6 / UC-B6A PDF file:"
        if state == STATE_NEW_YORK:
            return "New York NYS-45 PDF file:"
        return "California DE9C PDF file:"

    def _on_state_change(self, _event=None):
        self.pdf_label.configure(text=self._pdf_label_text())

    def pick_pdf(self):
        p = filedialog.askopenfilename(
            title=f"Select {self.state_var.get()} payroll PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if p:
            self.path_var.set(p)

    def logln(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def convert(self):
        pdf_path = Path(self.path_var.get().strip() or "")
        
        if not pdf_path.is_file():
            messagebox.showerror("Error", "Please choose a valid PDF file (a .pdf file, not a folder).")
            return

        if pdf_path.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Selected file is not a .pdf.")
            return
        
        try:
            state = self.state_var.get()
            text = extract_text_from_pdf(pdf_path)
            rows, dbg = parse_payroll_text_with_debug(text, state)
            rows, dropped_header, dropped_non_employee = _filter_out_header_rows(rows)
        except Exception as e:
            self.logln("[ERROR] " + str(e))
            messagebox.showerror("Error", f"Parse failed:\n{e}")
            return

        out_path = pdf_path.with_name(pdf_path.stem + " CSV.csv")
        write_csv_no_header(out_path, rows)

        self.log.delete("1.0", "end")
        self.logln(
            f"State: {self.state_var.get()}. Parsed {len(rows)} employee rows. "
            f"Removed {dropped_header} header rows and {dropped_non_employee} non-employee rows. "
            f"Saved to: {out_path}"
        )
        self.logln("")
        for d in dbg[:80]:
            self.logln(f'[{d["capture"]}] {d["name_lines"]} → {d["first"]} {d["mi"]} {d["last"]} | {d["F"]},{d["G"]},{d["H"]}')
            
        messagebox.showinfo(
            "Done",
            f"CSV saved:\n{out_path}\n\nRemoved {dropped_header} header rows and {dropped_non_employee} non-employee rows."
        )

# =========================
# CLI
# =========================
def _print_debug_to_console(pdf_path: Path, state: str = STATE_CALIFORNIA):
    text = extract_text_from_pdf(pdf_path)
    rows, dbg = parse_payroll_text_with_debug(text, state)
    rows, dropped_header, dropped_non_employee = _filter_out_header_rows(rows)
    print(f"State: {state}")
    print(f"Rows: {len(rows)} (removed {dropped_header} header rows, {dropped_non_employee} non-employee rows)")

    for r in rows:
        print(r)

if __name__ == "__main__":
    if "--debug" in sys.argv:
        try:
            pdf_idx = sys.argv.index("--pdf") + 1
            pdf_path = Path(sys.argv[pdf_idx])
        except Exception:
            print("Usage: python de9c_to_csv.py --debug --pdf <PDF_PATH> [--state California|New York|Hawaii]")  # noqa
            sys.exit(2)
        state = STATE_CALIFORNIA
        if "--state" in sys.argv:
            try:
                state = sys.argv[sys.argv.index("--state") + 1]
            except Exception:
                print("Usage: python de9c_to_csv.py --debug --pdf <PDF_PATH> [--state California|New York|Hawaii]")  # noqa
                sys.exit(2)
        if not pdf_path.is_file():
            print(f"PDF not found: {pdf_path}")
            sys.exit(2)
        _print_debug_to_console(pdf_path, state)
    else:
        App().mainloop()
