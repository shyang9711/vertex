"""
Parse U.S. Bank Business Checking statement PDF and pasted text.
Output: list of dicts with date, description, amount (negative = withdrawals, positive = deposits), reference_number.

Supports two layouts:
1) Customer Deposits (simple): Date line, Ref line, Amount line (e.g. Jan 15, 8612700360, 100.00).
2) Other Deposits / Card Withdrawals / Other Withdrawals: "Mon DD Description..." block, then amount ($X.XX or X.XX-).
"""
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


# Statement period: "Statement Period: Jan 7, 2025 through Jan 31, 2025" or "Feb 3, 2025 through Feb 28, 2025"
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

RE_AMOUNT = re.compile(r"^\s*\$?\s*([\d,]+)\.(\d{2})\s*(-)?\s*$")
RE_REF = re.compile(r"^\d{4,}\*?\s*$")
RE_ALNUM_REF = re.compile(r"^[A-Z0-9]{8,}\*?\s*$", re.I)
RE_REF_IN_DESC = re.compile(r"REF[=#]?\s*([A-Z0-9]+)", re.I)


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
    m = RE_AMOUNT.match(s)
    if not m:
        return None
    val = float(m.group(1).replace(",", "") + "." + m.group(2))
    return -val if m.group(3) == "-" else val


def _is_month_only(line: str) -> bool:
    s = (line or "").strip().capitalize()[:3]
    return s in MONTH_NUM


def _get_date_start(lines: list[str], i: int):
    """
    Returns (month_abbrev, day, rest, next_index) or None.
    Supports:
      - 'Sep 10 Real Time Payment Credit'
      - 'Sep 10'
      - 'Sep' / '10' / next line...
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

    if _is_month_only(s) and i + 1 < len(lines) and (lines[i + 1] or "").strip().isdigit():
        return s.capitalize()[:3], lines[i + 1].strip(), "", i + 2

    return None


def _is_header_noise(line: str) -> bool:
    low = (line or "").strip().lower()
    if not low:
        return True

    header_starts = (
        "business statement",
        "gold business checking",
        "(continued)",
        "u.s. bank national association",
        "account number",
        "statement period",
        "page ",
        "flucco",
        "dba ",
        "15206 ",
        "date",
        "description of transaction",
        "ref number",
        "amount",
        "check",
        "payee",
        "ending balance",
        "to contact u.s. bank",
        "reserve line",
        "balance summary",
        "analysis service charge detail",
        "conventional checks paid",
        "electronic checks paid",
        "total checks paid",
        "total other deposits",
        "total other withdrawals",
        "* gap in check sequence",
    )
    return any(low.startswith(x) for x in header_starts)


def _parse_summary_counts(lines: list[str]) -> dict[str, int]:
    """
    Reads the opening summary block:
      Customer Deposits / 1
      Other Deposits / 65
      Other Withdrawals / 18
      Checks Paid / 89
    """
    counts = {
        "customer_deposits": 0,
        "other_deposits": 0,
        "other_withdrawals": 0,
        "checks_paid": 0,
    }

    mapping = {
        "customer deposits": "customer_deposits",
        "other deposits": "other_deposits",
        "other withdrawals": "other_withdrawals",
        "checks paid": "checks_paid",
    }

    for i, line in enumerate(lines[:60]):
        low = (line or "").strip().lower()
        if low in mapping:
            for j in range(i + 1, min(i + 4, len(lines))):
                nxt = (lines[j] or "").strip()
                if nxt.isdigit():
                    counts[mapping[low]] = int(nxt)
                    break
    return counts


def _parse_customer_deposits(lines: list[str], year: int, expected_count: int) -> list[dict]:
    """
    Parse the detailed customer deposits at the start of the statement.
    Example:
      Sep 2
      8316110051
      5,727.00
    """
    out = []
    i = 0

    while i < len(lines) and len(out) < expected_count:
        ds = _get_date_start(lines, i)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        ref_num = ""
        amount_val = None

        if rest and RE_ALNUM_REF.match(rest):
            ref_num = rest.rstrip("*")

        while j < len(lines) and j < i + 6:
            nxt = (lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = abs(amt)
                j += 1
                break

            if nxt == "$" and j + 1 < len(lines):
                amt = _parse_amount_line(lines[j + 1])
                if amt is not None:
                    amount_val = abs(amt)
                    j += 2
                    break

            if not ref_num and RE_ALNUM_REF.match(nxt):
                ref_num = nxt.rstrip("*")
                j += 1
                continue

            if _get_date_start(lines, j):
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
            continue

        i += 1

    return out


def _parse_other_deposits(lines: list[str], year: int, expected_count: int) -> list[dict]:
    out = []
    i = 0

    while i < len(lines) and len(out) < expected_count:
        ds = _get_date_start(lines, i)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        desc_parts = [rest] if rest else []
        amount_val = None

        while j < len(lines) and j < i + 20:
            nxt = (lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if _get_date_start(lines, j):
                break

            if nxt == "$" and j + 1 < len(lines):
                amt = _parse_amount_line(lines[j + 1])
                if amt is not None:
                    amount_val = abs(amt)
                    j += 2
                    break

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = abs(amt)
                j += 1
                break

            if not _is_header_noise(nxt):
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
            continue

        i += 1

    return out


def _parse_other_withdrawals(lines: list[str], year: int, expected_count: int) -> list[dict]:
    out = []
    i = 0

    while i < len(lines) and len(out) < expected_count:
        ds = _get_date_start(lines, i)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds
        desc_parts = [rest] if rest else []
        amount_val = None

        while j < len(lines) and j < i + 20:
            nxt = (lines[j] or "").strip()
            if not nxt:
                j += 1
                continue

            if _get_date_start(lines, j):
                break

            if nxt == "$" and j + 1 < len(lines):
                amt = _parse_amount_line(lines[j + 1])
                if amt is not None:
                    amount_val = -abs(amt)
                    j += 2
                    break

            amt = _parse_amount_line(nxt)
            if amt is not None:
                amount_val = -abs(amt)
                j += 1
                break

            if not _is_header_noise(nxt):
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
            continue

        i += 1

    return out


def _parse_checks_sections(lines: list[str], year: int) -> list[dict]:
    """
    Parses both:
      1) Checks Presented Conventionally
         12383
         Sep 2
         8316639870
         1,715.55

      2) Checks Presented Electronically
         12418
         Sep 10
         500.78 CHECKPAYMT
         AUDI FINCL, INC.
    """
    rows = []
    i = 0
    mode = None

    while i < len(lines):
        low = (lines[i] or "").strip().lower()

        if "checks presented conventionally" in low:
            mode = "conventional"
            i += 1
            continue

        if "checks presented electronically" in low:
            mode = "electronic"
            i += 1
            continue

        if "balance summary" in low:
            mode = None
            i += 1
            continue

        if not mode:
            i += 1
            continue

        s = (lines[i] or "").strip()
        if not RE_REF.match(s):
            i += 1
            continue

        check_no = s.rstrip("*")
        ds = _get_date_start(lines, i + 1)
        if not ds:
            i += 1
            continue

        month_abbrev, day, rest, j = ds

        if mode == "conventional":
            ref_num = ""
            amount_val = None

            if rest and RE_ALNUM_REF.match(rest):
                ref_num = rest.rstrip("*")

            while j < len(lines) and j < i + 6:
                nxt = (lines[j] or "").strip()
                if not nxt:
                    j += 1
                    continue

                amt = _parse_amount_line(nxt)
                if amt is not None:
                    amount_val = -abs(amt)
                    j += 1
                    break

                if nxt == "$" and j + 1 < len(lines):
                    amt = _parse_amount_line(lines[j + 1])
                    if amt is not None:
                        amount_val = -abs(amt)
                        j += 2
                        break

                if not ref_num and RE_ALNUM_REF.match(nxt):
                    ref_num = nxt.rstrip("*")
                    j += 1
                    continue

                if RE_REF.match(nxt) or _get_date_start(lines, j):
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
                continue

        elif mode == "electronic":
            amount_val = None
            desc_parts = []

            while j < len(lines) and j < i + 6:
                nxt = (lines[j] or "").strip()
                if not nxt:
                    j += 1
                    continue

                if RE_REF.match(nxt) or _get_date_start(lines, j) or "balance summary" in nxt.lower():
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
                continue

        i += 1

    return rows


def _parse_us_bank_checking_text(text: str) -> list[dict]:
    lines = [ln.strip() for ln in text.split("\n")]
    year = _extract_statement_year(text) or datetime.now().year

    # Start at the detailed transaction section, not the decorative/header block.
    start_idx = 0
    for idx, ln in enumerate(lines):
        if (ln or "").strip().lower().startswith("beginning balance on"):
            start_idx = idx
            break
    lines = lines[start_idx:]

    counts = _parse_summary_counts(lines)

    rows = []

    # Customer Deposits appear immediately after the opening summary block in this statement style.
    rows.extend(_parse_customer_deposits(lines, year, counts["customer_deposits"]))

    # Other Deposits / Other Withdrawals can span multiple pages and continued headers.
    rows.extend(_parse_other_deposits(lines, year, counts["other_deposits"]))
    rows.extend(_parse_other_withdrawals(lines, year, counts["other_withdrawals"]))

    # Checks Presented Conventionally / Electronically
    rows.extend(_parse_checks_sections(lines, year))

    # Deduplicate while preserving order
    deduped = []
    seen = set()
    for r in rows:
        key = (r["date"], r["description"], round(float(r["amount"]), 2), r["reference_number"])
        if key not in seen:
            seen.add(key)
            deduped.append(r)

    return deduped


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
