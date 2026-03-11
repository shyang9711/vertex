import re
import sys
import os
from datetime import datetime

MONTH_ABBREV = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTH_ABBREV)}


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from PDF using PyMuPDF (fitz)."""
    import fitz
    parts = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            parts.append(page.get_text("text"))
    text = "\n".join(parts).replace("\r", "\n")
    text = text.replace("\t", " ")
    text = re.sub(r" {2,}", " ", text)
    return text


RE_STATEMENT_PERIOD = re.compile(
    r"(?:Statement\s+Period:\s*)?"
    r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),?\s+(\d{4})\s+through\s+",
    re.I
)

RE_MONTH_DAY = re.compile(
    r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s*$",
    re.I
)

RE_MONTH_DAY_WITH_REST = re.compile(
    r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(.+)$",
    re.I
)

RE_MONTH_ONLY = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$", re.I)
RE_AMOUNT_ONLY = re.compile(r"^\$?\s*([\d,]+)\.(\d{2})\s*(-)?\s*$")
RE_REF = re.compile(r"^\d{4,}\*?\s*$")
RE_ALNUM_REF = re.compile(r"^[A-Z0-9]{8,}\*?\s*$", re.I)
RE_REF_IN_DESC = re.compile(r"REF\s*[=#]?\s*([A-Z0-9]+)", re.I)


def _extract_statement_year(text: str) -> int | None:
    m = RE_STATEMENT_PERIOD.search(text)
    if m:
        y = int(m.group(3))
        if 1990 <= y <= 2035:
            return y
    for m in re.finditer(r"\b(20\d{2})\b", text[:3000]):
        y = int(m.group(1))
        if 1990 <= y <= 2035:
            return y
    return None


def _normalize_date(month_abbrev: str, day: str | int, year: int) -> str:
    mo = MONTH_NUM.get(month_abbrev.capitalize()[:3])
    if mo is None:
        return ""
    try:
        d = int(day)
        if 1 <= d <= 31 and 1990 <= year <= 2035:
            return f"{mo:02d}/{d:02d}/{year}"
    except ValueError:
        pass
    return ""


def _parse_amount_line(line: str) -> float | None:
    s = (line or "").strip()
    if not s:
        return None
    m = RE_AMOUNT_ONLY.match(s)
    if not m:
        return None
    val = float(m.group(1).replace(",", "") + "." + m.group(2))
    return -val if m.group(3) == "-" else val


def _get_date_start(lines: list[str], i: int):
    """
    Supports:
      - 'Sep 10 Real Time Payment Credit'
      - 'Sep 10'
      - 'Sep' / '10'
    Returns (month_abbrev, day, rest, next_index) or None
    """
    if i >= len(lines):
        return None

    s = (lines[i] or "").strip()
    if not s:
        return None

    m = RE_MONTH_DAY_WITH_REST.match(s)
    if m:
        return m.group(1).capitalize()[:3], m.group(2), m.group(3).strip(), i + 1

    m = RE_MONTH_DAY.match(s)
    if m:
        return m.group(1).capitalize()[:3], m.group(2), "", i + 1

    if RE_MONTH_ONLY.match(s) and i + 1 < len(lines):
        nxt = (lines[i + 1] or "").strip()
        if nxt.isdigit():
            return s.capitalize()[:3], nxt, "", i + 2

    return None


def _clean_lines(text: str) -> list[str]:
    return [ln.strip() for ln in text.split("\n")]


def _find_first(lines: list[str], predicate, start: int = 0) -> int:
    for i in range(start, len(lines)):
        if predicate(lines[i]):
            return i
    return -1


def _find_next_header(lines: list[str], start: int, headers: tuple[str, ...]) -> int:
    headers_lower = tuple(h.lower() for h in headers)
    for i in range(start, len(lines)):
        low = (lines[i] or "").strip().lower()
        if any(h in low for h in headers_lower):
            return i
    return len(lines)


def _slice_section(lines: list[str], start_headers: tuple[str, ...], end_headers: tuple[str, ...], start_at: int = 0) -> list[str]:
    start = _find_next_header(lines, start_at, start_headers)
    if start == len(lines):
        return []
    end = _find_next_header(lines, start + 1, end_headers)
    return lines[start:end]


def _parse_customer_deposits(section_lines: list[str], year: int) -> list[dict]:
    out = []
    i = 0
    while i < len(section_lines):
        ds = _get_date_start(section_lines, i)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        ref_num = ""
        amount_val = None

        if rest and RE_ALNUM_REF.match(rest):
            ref_num = rest.rstrip("*")

        while j < len(section_lines) and j < i + 6:
            nxt = (section_lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if not ref_num and RE_ALNUM_REF.match(nxt):
                ref_num = nxt.rstrip("*")
                j += 1
                continue

            if nxt == "$" and j + 1 < len(section_lines):
                amt = _parse_amount_line(section_lines[j + 1])
                if amt is not None:
                    amount_val = abs(amt)
                    j += 2
                    break

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = abs(amt)
                j += 1
                break

            if _get_date_start(section_lines, j):
                break

            j += 1

        if amount_val is not None:
            out.append({
                "date": _normalize_date(month_abbrev, day, year),
                "description": f"Deposit {ref_num}" if ref_num else "Customer Deposit",
                "amount": amount_val,
                "reference_number": ref_num
            })
            i = j
        else:
            i += 1

    return out


def _parse_other_deposits(section_lines: list[str], year: int) -> list[dict]:
    out = []
    i = 0
    while i < len(section_lines):
        ds = _get_date_start(section_lines, i)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        desc_parts = [rest] if rest else []
        amount_val = None

        while j < len(section_lines) and j < i + 20:
            nxt = (section_lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if _get_date_start(section_lines, j):
                break

            if nxt == "$" and j + 1 < len(section_lines):
                amt = _parse_amount_line(section_lines[j + 1])
                if amt is not None:
                    amount_val = abs(amt)
                    j += 2
                    break

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = abs(amt)
                j += 1
                break

            low = nxt.lower()
            if low not in {
                "other deposits",
                "other deposits (continued)",
                "date",
                "description of transaction",
                "ref number",
                "amount",
            }:
                desc_parts.append(nxt)

            j += 1

        if amount_val is not None:
            description = " ".join(x for x in desc_parts if x).strip()
            ref_num = ""
            m = RE_REF_IN_DESC.search(description)
            if m:
                ref_num = m.group(1)

            out.append({
                "date": _normalize_date(month_abbrev, day, year),
                "description": description,
                "amount": amount_val,
                "reference_number": ref_num
            })
            i = j
        else:
            i += 1

    return out


def _parse_other_withdrawals(section_lines: list[str], year: int) -> list[dict]:
    out = []
    i = 0
    while i < len(section_lines):
        ds = _get_date_start(section_lines, i)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        desc_parts = [rest] if rest else []
        amount_val = None

        while j < len(section_lines) and j < i + 20:
            nxt = (section_lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if _get_date_start(section_lines, j):
                break

            if nxt == "$" and j + 1 < len(section_lines):
                amt = _parse_amount_line(section_lines[j + 1])
                if amt is not None:
                    amount_val = -abs(amt)
                    j += 2
                    break

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = -abs(amt)
                j += 1
                break

            low = nxt.lower()
            if low not in {
                "other withdrawals",
                "other withdrawals (continued)",
                "date",
                "description of transaction",
                "ref number",
                "amount",
                "check",
            }:
                desc_parts.append(nxt)

            j += 1

        if amount_val is not None:
            description = " ".join(x for x in desc_parts if x).strip()
            ref_num = ""
            m = RE_REF_IN_DESC.search(description)
            if m:
                ref_num = m.group(1)

            out.append({
                "date": _normalize_date(month_abbrev, day, year),
                "description": description,
                "amount": amount_val,
                "reference_number": ref_num
            })
            i = j
        else:
            i += 1

    return out


def _parse_checks_conventional(section_lines: list[str], year: int) -> list[dict]:
    rows = []
    i = 0
    while i < len(section_lines):
        s = (section_lines[i] or "").strip()
        if not RE_REF.match(s):
            i += 1
            continue

        check_no = s.rstrip("*")
        ds = _get_date_start(section_lines, i + 1)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        ref_num = ""
        amount_val = None

        if rest and RE_ALNUM_REF.match(rest):
            ref_num = rest.rstrip("*")

        while j < len(section_lines) and j < i + 6:
            nxt = (section_lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if not ref_num and RE_ALNUM_REF.match(nxt):
                ref_num = nxt.rstrip("*")
                j += 1
                continue

            if nxt == "$" and j + 1 < len(section_lines):
                amt = _parse_amount_line(section_lines[j + 1])
                if amt is not None:
                    amount_val = -abs(amt)
                    j += 2
                    break

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = -abs(amt)
                j += 1
                break

            if RE_REF.match(nxt) or _get_date_start(section_lines, j):
                break

            j += 1

        if amount_val is not None:
            rows.append({
                "date": _normalize_date(month_abbrev, day, year),
                "description": f"Check {check_no}",
                "amount": amount_val,
                "reference_number": ref_num or check_no
            })
            i = j
        else:
            i += 1

    return rows


def _parse_checks_electronic(section_lines: list[str], year: int) -> list[dict]:
    rows = []
    i = 0
    while i < len(section_lines):
        s = (section_lines[i] or "").strip()
        if not RE_REF.match(s):
            i += 1
            continue

        check_no = s.rstrip("*")
        ds = _get_date_start(section_lines, i + 1)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        amount_val = None
        desc_parts = []

        if rest:
            m = re.match(r"^\s*([\d,]+\.\d{2})\s*(.*)$", rest)
            if m:
                amount_val = -abs(float(m.group(1).replace(",", "")))
                if m.group(2).strip():
                    desc_parts.append(m.group(2).strip())
            else:
                desc_parts.append(rest)

        while j < len(section_lines) and j < i + 8:
            nxt = (section_lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if RE_REF.match(nxt) or _get_date_start(section_lines, j):
                break

            m = re.match(r"^\s*([\d,]+\.\d{2})\s*(.*)$", nxt)
            if m and amount_val is None:
                amount_val = -abs(float(m.group(1).replace(",", "")))
                if m.group(2).strip():
                    desc_parts.append(m.group(2).strip())
                j += 1
                continue

            amt = _parse_amount_line(nxt)
            if amt is not None and amount_val is None:
                amount_val = -abs(amt)
                j += 1
                continue

            desc_parts.append(nxt)
            j += 1

        if amount_val is not None:
            rows.append({
                "date": _normalize_date(month_abbrev, day, year),
                "description": " ".join(desc_parts).strip() or f"Electronic Check {check_no}",
                "amount": amount_val,
                "reference_number": check_no
            })
            i = j
        else:
            i += 1

    return rows


def _dedupe_rows(rows: list[dict]) -> list[dict]:
    out = []
    seen = set()
    for r in rows:
        key = (
            r["date"],
            r["description"],
            round(float(r["amount"]), 2),
            r["reference_number"],
        )
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


def _parse_us_bank_checking_text(text: str) -> list[dict]:
    lines = _clean_lines(text)
    year = _extract_statement_year(text) or datetime.now().year

    # Start at transaction body
    start_idx = 0
    for idx, ln in enumerate(lines):
        if (ln or "").strip().lower().startswith("beginning balance on"):
            start_idx = idx
            break
    lines = lines[start_idx:]

    customer_section = _slice_section(
        lines,
        start_headers=("customer deposits",),
        end_headers=("other deposits", "other withdrawals", "checks presented conventionally", "checks presented electronically")
    )

    other_deposits_section = _slice_section(
        lines,
        start_headers=("other deposits",),
        end_headers=("other withdrawals", "checks presented conventionally", "checks presented electronically")
    )

    other_withdrawals_section = _slice_section(
        lines,
        start_headers=("other withdrawals",),
        end_headers=("checks presented conventionally", "checks presented electronically", "balance summary")
    )

    conventional_checks_section = _slice_section(
        lines,
        start_headers=("checks presented conventionally",),
        end_headers=("checks presented electronically", "balance summary")
    )

    electronic_checks_section = _slice_section(
        lines,
        start_headers=("checks presented electronically",),
        end_headers=("balance summary",)
    )

    rows = []
    rows.extend(_parse_customer_deposits(customer_section, year))
    rows.extend(_parse_other_deposits(other_deposits_section, year))
    rows.extend(_parse_other_withdrawals(other_withdrawals_section, year))
    rows.extend(_parse_checks_conventional(conventional_checks_section, year))
    rows.extend(_parse_checks_electronic(electronic_checks_section, year))

    return _dedupe_rows(rows)


def parse_us_bank_checking_text(text: str) -> list[dict]:
    return _parse_us_bank_checking_text(text)


def parse_us_bank_checking_pdf(pdf_path: str) -> list[dict]:
    text = extract_text_from_pdf(pdf_path)
    return parse_us_bank_checking_text(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_us_bank_checking.py <pdf_path> [--csv output.csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    rows = parse_us_bank_checking_pdf(pdf_path)
    print(f"Parsed {len(rows)} transactions")
    for r in rows:
        print(r)
    if len(sys.argv) >= 4 and sys.argv[2] == "--csv":
        import csv
        with open(sys.argv[3], "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["date", "description", "amount", "reference_number"])
            w.writeheader()
            w.writerows(rows)
        print(f"Wrote {sys.argv[3]}")


if __name__ == "__main__":
    main()