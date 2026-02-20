"""
Parse Comerica Bank checking statement PDF and pasted text.
Output: date (MM/DD/YYYY), description, amount (negative = debits, positive = credits).

PDF format: Multi-line blocks per transaction
  Date (e.g. Jan 02)
  Amount (e.g. 7,280.28 or -4,589.96)
  Activity (description, one or more lines)
  Optional reference (digits)

Pasted format: Single line per transaction
  Mon DD Amount Description
  e.g. May 01 3,564.29 Bankcard 8076 Mtot Dep 250430 554402000721191 9488197080
"""
import re
import sys
import csv
import os
from datetime import datetime

MONTH_ABBREV = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTH_ABBREV)}

# Pasted: Mon DD Amount Description (amount may be negative)
COMERICA_PASTED_LINE = re.compile(
    r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*(\d{1,2})\s+(-?[\d,]+\.\d{2})\s+(.+)$",
    re.I
)

# PDF: date line only (Month Day), with or without space (Jan 02 or Jan02)
COMERICA_DATE_LINE = re.compile(
    r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*(\d{1,2})\s*$",
    re.I
)

# PDF: amount line only
COMERICA_AMOUNT_LINE = re.compile(
    r"^(-?[\d,]+\.\d{2})\s*$"
)

# Checks paid section: check number line (#612, @614, #10147). Must have #/@/* prefix so we don't match bank ref digits.
COMERICA_CHECK_NUM_LINE = re.compile(
    r"^[#@*](\d+)\s*$"
)

# Statement period: "January 1, 2026 to January 31, 2026" or "January 1, 2025 to January 31, 2025"
RE_STATEMENT_PERIOD = re.compile(
    r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+(\d{4})\s+to\s+",
    re.I
)

# Skip section headers / footers
SKIP_PATTERNS = [
    "Reference numbers",
    "Date Amount Activity",
    "Date\nAmount",
    "Total Electronic Deposits",
    "Total Number of Electronic",
    "Total Other Deposits",
    "Total Number of Other",
    "Total checks paid",
    "Total number of checks",
    "Total ATM/Debit Card",
    "Total Number of ATM",
    "Total Electronic withdrawals",
    "Total Number of Electronic",
    "this statement period",
    "Account summary",
    "Beginning balance",
    "Ending balance",
    "Page ",
    "Basic Business Checking",
    "Account number",
]


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


def _extract_statement_year(text: str) -> int | None:
    """Get statement year from 'January 1, 2026 to January 31, 2026' or similar."""
    m = RE_STATEMENT_PERIOD.search(text)
    if m:
        y = int(m.group(1))
        if 1990 <= y <= 2030:
            return y
    # Fallback: any 4-digit year in first 500 chars
    for m in re.finditer(r"\b(20\d{2})\b", text[:500]):
        y = int(m.group(1))
        if 1990 <= y <= 2030:
            return y
    return None


def _normalize_date(month_abbrev: str, day: str, year: int | None) -> str:
    """Return MM/DD/YYYY."""
    if year is None:
        year = datetime.now().year
    mo = MONTH_NUM.get(month_abbrev.capitalize()[:3])
    if mo is None:
        return ""
    try:
        d = int(day)
        if 1 <= d <= 31 and 1990 <= year <= 2030:
            return f"{mo:02d}/{d:02d}/{year}"
    except ValueError:
        pass
    return ""


def _is_skip_line(line: str) -> bool:
    line_strip = line.strip()
    if not line_strip:
        return True
    line_upper = line_strip.upper()
    for pat in SKIP_PATTERNS:
        if pat.upper() in line_upper or pat in line_strip:
            return True
    return False


def _parse_comerica_checks_section(lines: list, start_i: int, year: int) -> tuple[list[dict], int]:
    """
    Parse 'Checks paid this statement period' section.
    Block format: Check number (#612) -> Amount (-749.20) -> Date (Jan 03) -> Bank reference (digits).
    Returns (list of {date, description, amount, reference_number}, index after section).
    """
    rows = []
    i = start_i
    # Skip to first check number line (past section title and column headers)
    while i < len(lines):
        if COMERICA_CHECK_NUM_LINE.match(lines[i].strip()):
            break
        if "Total checks paid" in lines[i] or "Total number of checks" in lines[i]:
            return rows, i
        i += 1
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if "Total checks paid" in line or "Total number of checks" in line:
            break
        check_m = COMERICA_CHECK_NUM_LINE.match(line)
        if check_m:
            check_num = check_m.group(1)
            i += 1
            amount_str = None
            date_str = None
            while i < len(lines):
                cur = lines[i].strip()
                if not cur:
                    i += 1
                    continue
                if COMERICA_CHECK_NUM_LINE.match(cur):
                    break
                if COMERICA_AMOUNT_LINE.match(cur) and amount_str is None:
                    amount_str = cur.replace(",", "").strip()
                    i += 1
                    continue
                date_m = COMERICA_DATE_LINE.match(cur)
                if date_m and date_str is None:
                    date_str = _normalize_date(date_m.group(1), date_m.group(2), year)
                    i += 1
                    continue
                if re.match(r"^\d{7,}\s*$", cur):
                    i += 1
                    break
                i += 1
            if amount_str and date_str:
                try:
                    amount = float(amount_str)
                    rows.append({
                        "date": date_str,
                        "description": f"Check {check_num}",
                        "amount": amount,
                        "reference_number": check_num,
                    })
                except ValueError:
                    pass
            continue
        i += 1
    return rows, i


def parse_comerica_pasted_text(text: str) -> list[dict]:
    """
    Parse pasted Comerica text: one line per transaction.
    Format: Mon DD Amount Description
    Returns list of {date, description, amount}.
    """
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    year = _extract_statement_year(text)
    if year is None:
        year = datetime.now().year
    rows = []
    for line in lines:
        if _is_skip_line(line):
            continue
        m = COMERICA_PASTED_LINE.match(line)
        if m:
            month_abbrev, day, amount_str, description = m.groups()
            try:
                amount = float(amount_str.replace(",", ""))
                date_str = _normalize_date(month_abbrev, day, year)
                if date_str:
                    rows.append({"date": date_str, "description": description.strip(), "amount": amount})
            except ValueError:
                pass
    return rows


def parse_comerica_pdf_text(text: str) -> list[dict]:
    """
    Parse Comerica PDF-extracted text. Multi-line blocks: Date, Amount, Activity.
    Also parses 'Checks paid' section (check number -> amount -> date -> reference).
    Returns list of {date, description, amount} with optional reference_number for checks.
    """
    lines = text.split("\n")
    year = _extract_statement_year(text)
    if year is None:
        year = datetime.now().year
    rows = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if _is_skip_line(line):
            i += 1
            continue
        date_m = COMERICA_DATE_LINE.match(stripped)
        if date_m:
            month_abbrev, day = date_m.groups()
            i += 1
            amount_str = None
            activity_parts = []
            while i < len(lines):
                cur = lines[i].strip()
                if not cur:
                    i += 1
                    continue
                if COMERICA_DATE_LINE.match(cur):
                    break
                if COMERICA_CHECK_NUM_LINE.match(cur):
                    break
                if _is_skip_line(cur):
                    i += 1
                    continue
                amt_m = COMERICA_AMOUNT_LINE.match(cur)
                if amt_m and amount_str is None:
                    amount_str = amt_m.group(1).replace(",", "").strip()
                    i += 1
                    continue
                if re.match(r"^\d{7,}\s*$", cur):
                    i += 1
                    continue
                activity_parts.append(cur)
                i += 1
            if amount_str and activity_parts:
                try:
                    amount = float(amount_str)
                    date_str = _normalize_date(month_abbrev, day, year)
                    if date_str:
                        rows.append({
                            "date": date_str,
                            "description": " ".join(activity_parts).strip(),
                            "amount": amount
                        })
                except ValueError:
                    pass
            continue
        i += 1
    # Parse Checks paid section and append (with reference_number for Check Number)
    for j, ln in enumerate(lines):
        if "Checks paid this statement period" in ln:
            check_rows, _ = _parse_comerica_checks_section(lines, j, year)
            rows.extend(check_rows)
            break
    return rows


def parse_comerica_checking_text(text: str) -> list[dict]:
    """
    Parse Comerica statement text (PDF or pasted).
    Tries pasted (single-line) format first; falls back to PDF (multi-line) format.
    Returns list of {date, description, amount}.
    """
    pasted = parse_comerica_pasted_text(text)
    if pasted:
        return pasted
    return parse_comerica_pdf_text(text)


def parse_comerica_checking_pdf(pdf_path: str) -> list[dict]:
    """Parse Comerica checking statement PDF."""
    text = extract_text_from_pdf(pdf_path)
    return parse_comerica_checking_text(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_comerica_checking.py <pdf_path> [--csv output.csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    out_csv = None
    if len(sys.argv) >= 4 and sys.argv[2] == "--csv":
        out_csv = sys.argv[3]
    rows = parse_comerica_checking_pdf(pdf_path)
    if not rows:
        print("No transactions found.")
        sys.exit(0)
    print(f"Parsed {len(rows)} transactions")
    for r in rows[:15]:
        print(f"  {r['date']}  {r['amount']:>12.2f}  {r['description'][:60]}")
    if out_csv:
        with open(out_csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["date", "description", "amount"])
            w.writeheader()
            w.writerows(rows)
        print(f"Wrote {len(rows)} rows to {out_csv}")


if __name__ == "__main__":
    main()
