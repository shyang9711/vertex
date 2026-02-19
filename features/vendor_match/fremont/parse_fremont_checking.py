"""
Parse Fremont Bank checking statement PDF.
Output: date, description, amount (negative = debits, positive = credits).

Format: Fremont Bank Business Plus Checking
- Description line(s) (can span multiple lines)
- Amount and Date on same line: Amount (MM/DD)
- Balance on next line
- Amount can be negative (debits) or positive (credits)

Example:
SGWS of N. CA 3056254171
-1,571.10    01/02
26,751.99 

Usage:
  python parse_fremont_checking.py "path/to/statement.pdf"
  python parse_fremont_checking.py "path/to/statement.pdf" --csv "output.csv"
"""
import re
import sys
import csv
import os
from datetime import datetime


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


# Pattern to match amount and date on same line: Amount (MM/DD)
# Amount can be negative, may have commas, and date is MM/DD
FREMONT_AMOUNT_DATE_LINE = re.compile(
    r"^([-]?\s*[\d,]+\.\d{2})\s+(\d{1,2}/\d{1,2})\s*$"
)

# Single-line paste format: Description Amount MM/DD Balance (all on one line)
FREMONT_SINGLE_LINE = re.compile(
    r"^(.+?)\s+([-]?[\d,]+\.\d{2})\s+(\d{1,2}/\d{1,2})\s+([\d,]+\.\d{2})\s*$"
)

# Pattern to match balance line: Balance amount
FREMONT_BALANCE_LINE = re.compile(
    r"^([\d,]+\.\d{2})\s*$"
)

# Statement period / account as of — to extract year for MM/DD
RE_THROUGH_FULL = re.compile(r"through\s+(\d{1,2})/(\d{1,2})/(\d{4})", re.I)
RE_AS_OF_FULL = re.compile(r"(?:account\s+)?as\s+of\s+(\d{1,2})/(\d{1,2})/(\d{4})", re.I)
RE_PERIOD_RANGE = re.compile(
    r"(\d{1,2})/(\d{1,2})/(\d{4})\s*(?:-|through|to)\s*(\d{1,2})/(\d{1,2})/(\d{4})",
    re.I
)
RE_ANY_MM_DD_YYYY = re.compile(r"\d{1,2}/\d{1,2}/(\d{4})")
RE_MONTH_YEAR = re.compile(r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+(\d{4})", re.I)

# Statement date pattern
RE_STATEMENT_DATE = re.compile(r"STATEMENT DATE\s+(\d{1,2})/(\d{1,2})/(\d{4})", re.I)

# Skip these lines (header, balance lines, footers)
SKIP_PATTERNS = [
    "BALANCE LAST STATEMENT",
    "BALANCE THIS STATEMENT",
    "TOTAL DAYS IN STATEMENT PERIOD",
    "ITEMIZATION OF SERVICE CHARGE",
    "TOTAL CHARGE FOR",
    "TOTAL CHARGE",
    "Description Transaction Amount",
    "Page:",
    "STATEMENT DATE",
    "ACCOUNT NUMBER",
    "Phone:",
    "Website:",
    "SPLITHERE",
]


def _is_skip_line(line: str) -> bool:
    """Check if line should be skipped."""
    line_upper = line.upper().strip()
    for pattern in SKIP_PATTERNS:
        if pattern in line_upper:
            return True
    # Skip empty lines
    if not line.strip():
        return True
    # Skip lines that are just numbers or dates without context
    if re.match(r"^\d{1,2}/\d{1,2}$", line.strip()):
        return False  # This might be part of a transaction
    return False


def _extract_statement_period(text: str) -> tuple[int | None, int | None, int | None, int | None]:
    """(start_month, start_year, end_month, end_year). For Dec–Jan statements, used to assign year by transaction month."""
    lines = text.split("\n")
    head = " ".join(lines[:80])
    m = RE_PERIOD_RANGE.search(head)
    if m:
        sm, sd, sy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        em, ed, ey = int(m.group(4)), int(m.group(5)), int(m.group(6))
        if 1990 <= sy <= 2030 and 1990 <= ey <= 2030:
            return (sm, sy, em, ey)
    for pat in (RE_THROUGH_FULL, RE_AS_OF_FULL):
        m = pat.search(head)
        if m:
            em, ed, ey = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1990 <= ey <= 2030:
                return (None, None, em, ey)
    # Try to extract from statement date
    m = RE_STATEMENT_DATE.search(head)
    if m:
        em, ed, ey = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1990 <= ey <= 2030:
            return (None, None, em, ey)
    return (None, None, None, None)


def _extract_statement_year(text: str) -> int | None:
    """Retrieve statement year from header."""
    start_m, start_y, end_m, end_y = _extract_statement_period(text)
    if end_y is not None:
        return end_y
    if start_y is not None:
        return start_y
    lines = text.split("\n")
    head = " ".join(lines[:80])
    m = RE_MONTH_YEAR.search(head)
    if m:
        y = int(m.group(1))
        if 1990 <= y <= 2030:
            return y
    for line in lines[:40]:
        for m in RE_ANY_MM_DD_YYYY.finditer(line):
            y = int(m.group(1))
            if 1990 <= y <= 2030:
                return y
    return None


def _year_for_transaction_month(tx_month: int, period: tuple) -> int | None:
    """Return year for a transaction in tx_month (1-12) given period (start_m, start_y, end_m, end_y)."""
    start_m, start_y, end_m, end_y = period
    if end_y is None:
        return start_y
    if start_m is not None and start_y is not None and start_y != end_y and end_m is not None:
        if end_m < start_m:
            if tx_month <= end_m:
                return end_y
            return start_y
    return end_y or start_y


def _normalize_date_to_mm_dd_yyyy(date_str: str, statement_year: int | None, period: tuple = (None, None, None, None)) -> str:
    """Normalize MM/DD or MM/DD/YY to MM/DD/YYYY. Uses period to pick year when statement spans two years (e.g. Dec–Jan)."""
    if not date_str or not date_str.strip():
        return ""
    date_str = date_str.strip()
    parts = date_str.split("/")
    if len(parts) == 3:
        try:
            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
            year = 2000 + y if y < 100 and y < 50 else (1900 + y if y < 100 else y)
            return f"{m:02d}/{d:02d}/{year}"
        except (ValueError, IndexError):
            return date_str
    if len(parts) == 2:
        try:
            m, d = int(parts[0]), int(parts[1])
            year = _year_for_transaction_month(m, period)
            if year is None or not (1990 <= year <= 2030):
                year = statement_year if statement_year and 1990 <= statement_year <= 2030 else 2000
            return f"{m:02d}/{d:02d}/{year}"
        except (ValueError, IndexError):
            return date_str
    return date_str


def parse_fremont_checking_text(text: str) -> list[dict]:
    """
    Parse Fremont Bank checking statement text.
    Returns list of {"date": "MM/DD/YYYY", "description": "...", "amount": float}.
    Amount: negative = debit (money out), positive = credit (money in).
    """
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    i = 0
    
    # Find the start of transaction section (after "Description" header)
    in_transactions = False
    while i < len(lines):
        line = lines[i]
        if "Description" in line and "Transaction Amount" in line and "Date" in line:
            in_transactions = True
            i += 1
            break
        i += 1
    
    if not in_transactions:
        # Try to find transactions by looking for amount+date patterns
        i = 0
    
    period = _extract_statement_period(text)
    year = _extract_statement_year(text)
    if year is None:
        year = datetime.now().year  # pasted text often has no header

    while i < len(lines):
        line = lines[i]
        
        # Skip header/metadata lines
        if _is_skip_line(line):
            i += 1
            continue
        
        # Single-line paste format: Description Amount MM/DD Balance
        single_match = FREMONT_SINGLE_LINE.match(line.strip())
        if single_match:
            desc_str = single_match.group(1).strip()
            amount_str = single_match.group(2).replace(",", "").strip()
            date_str = single_match.group(3)
            # Never treat balance-only lines as transactions (no transaction amount)
            desc_upper = desc_str.upper()
            if desc_upper.startswith("BALANCE LAST STATEMENT") or desc_upper.startswith("BALANCE THIS STATEMENT"):
                i += 1
                continue
            try:
                amount = float(amount_str)
                normalized_date = _normalize_date_to_mm_dd_yyyy(date_str, year, period)
                rows.append({
                    "date": normalized_date,
                    "description": desc_str,
                    "amount": amount
                })
            except ValueError:
                pass
            i += 1
            continue
        
        # Check if this line has amount and date (multi-line PDF format)
        amount_date_match = FREMONT_AMOUNT_DATE_LINE.match(line)
        if amount_date_match:
            # Found amount and date line
            amount_str = amount_date_match.group(1).replace(",", "").strip()
            date_str = amount_date_match.group(2)
            
            # Look backwards for description lines
            desc_parts = []
            j = i - 1
            while j >= 0:
                prev_line = lines[j].strip()
                if not prev_line:
                    j -= 1
                    continue
                # Stop if we hit another amount+date line or balance line
                if FREMONT_AMOUNT_DATE_LINE.match(prev_line) or FREMONT_BALANCE_LINE.match(prev_line):
                    break
                # Stop if we hit a skip pattern
                if _is_skip_line(prev_line):
                    break
                desc_parts.insert(0, prev_line)
                j -= 1
            
            # Look ahead for balance (optional)
            balance_val = None
            if i + 1 < len(lines):
                balance_match = FREMONT_BALANCE_LINE.match(lines[i + 1].strip())
                if balance_match:
                    balance_str = balance_match.group(1).replace(",", "").strip()
                    try:
                        balance_val = float(balance_str)
                    except ValueError:
                        pass
            
            if desc_parts:
                try:
                    amount = float(amount_str)
                    full_desc = " ".join(desc_parts).strip()
                    normalized_date = _normalize_date_to_mm_dd_yyyy(date_str, year, period)
                    
                    rows.append({
                        "date": normalized_date,
                        "description": full_desc,
                        "amount": amount
                    })
                except ValueError:
                    pass
        
        i += 1
    
    return rows


def parse_fremont_checking_pdf(pdf_path: str) -> list[dict]:
    """Parse Fremont Bank checking statement PDF. Returns list of {date, description, amount}."""
    text = extract_text_from_pdf(pdf_path)
    return parse_fremont_checking_text(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_fremont_checking.py <pdf_path> [--csv output.csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    out_csv = None
    if len(sys.argv) >= 4 and sys.argv[2] == "--csv":
        out_csv = sys.argv[3]
    rows = parse_fremont_checking_pdf(pdf_path)
    if not rows:
        print("No transactions found in the PDF.")
        sys.exit(0)
    print(f"Parsed {len(rows)} transactions")
    col_widths = (10, 50, 12)
    print(f"{'Date':<{col_widths[0]}} {'Description':<{col_widths[1]}} {'Amount':>{col_widths[2]}}")
    print("-" * (sum(col_widths) + 2))
    for r in rows:
        print(f"{r['date']:<{col_widths[0]}} {r['description'][:col_widths[1]]:<{col_widths[1]}} {r['amount']:>{col_widths[2]}.2f}")
    if out_csv:
        with open(out_csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["date", "description", "amount"])
            w.writeheader()
            w.writerows(rows)
        print(f"\nWrote {len(rows)} rows to {out_csv}")


if __name__ == "__main__":
    main()
