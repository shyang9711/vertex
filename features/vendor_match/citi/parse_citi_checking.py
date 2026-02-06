"""
Parse Citi (Citibank) checking/business statement PDF.
Output: date, description, amount (negative = debits, positive = credits).

Format: CitiBusiness CHECKING ACTIVITY
- Transaction line: MM/DD Description Amount Balance
- Continuation lines (no date) are appended to description.
- Amount is either debit or credit; sign inferred from description keywords.

Usage:
  python parse_citi_checking.py "path/to/statement.pdf"
  python parse_citi_checking.py "path/to/statement.pdf" --csv "output.csv"
"""
import re
import sys
import csv
import os


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


# Single-line format (if PDF exports one line per tx): MM/DD description amount balance
CITI_TX_LINE_SINGLE = re.compile(
    r"^(\d{1,2}/\d{1,2})\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2}-?)\s*$"
)
# Multi-line format: date on line 1, description on line 2, amount on line 3, balance on line 4
CITI_DATE_LINE = re.compile(r"^(\d{1,2}/\d{1,2})\s*$")
CITI_AMOUNT_LINE = re.compile(r"^([\d,]+\.\d{2}-?)\s*$")

# Statement period / account as of — to extract year for MM/DD
RE_THROUGH_FULL = re.compile(r"through\s+(\d{1,2})/(\d{1,2})/(\d{4})", re.I)
RE_AS_OF_FULL = re.compile(r"(?:account\s+)?as\s+of\s+(\d{1,2})/(\d{1,2})/(\d{4})", re.I)
RE_PERIOD_RANGE = re.compile(
    r"(\d{1,2})/(\d{1,2})/(\d{4})\s*(?:-|through|to)\s*(\d{1,2})/(\d{1,2})/(\d{4})",
    re.I
)
RE_ANY_MM_DD_YYYY = re.compile(r"\d{1,2}/\d{1,2}/(\d{4})")
RE_MONTH_YEAR = re.compile(r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+(\d{4})", re.I)

# Beginning/opening balance — use with balance column to determine debit vs credit
RE_BEGINNING_BALANCE = re.compile(
    r"(?:Beginning|Opening|Previous|Prior)\s+Balance\s*[:\$]?\s*([\d,]+\.\d{2})",
    re.I
)


def _parse_balance_value(s: str) -> float | None:
    """Parse balance string (e.g. '3,361.81' or '4,554.48-') to float."""
    if not s:
        return None
    s = s.replace(",", "").strip().rstrip("-")
    try:
        return float(s)
    except ValueError:
        return None


def _extract_beginning_balance(text: str) -> float | None:
    """Extract beginning/opening balance from statement text."""
    for m in RE_BEGINNING_BALANCE.finditer(text):
        v = _parse_balance_value(m.group(1))
        if v is not None:
            return v
    return None


def _apply_balance_to_sign(rows: list[dict], beginning_balance: float | None) -> None:
    """
    Set amount sign from balance progression: balance increased -> credit (positive),
    balance decreased -> debit (negative). Uses beginning balance for first row, then running balance.
    """
    prev_balance = beginning_balance
    for row in rows:
        bal = row.get("balance")
        if bal is None:
            if prev_balance is not None and row.get("amount") is not None:
                row["amount"] = -abs(float(row["amount"]))
            continue
        try:
            curr_balance = float(bal) if isinstance(bal, (int, float)) else _parse_balance_value(str(bal))
        except (TypeError, ValueError):
            curr_balance = None
        if curr_balance is not None and prev_balance is not None:
            raw_amt = abs(float(row.get("amount", 0) or 0))
            if curr_balance > prev_balance:
                row["amount"] = raw_amt   # credit
            else:
                row["amount"] = -raw_amt  # debit
        prev_balance = curr_balance if curr_balance is not None else prev_balance
        row.pop("balance", None)


def _extract_statement_period(text: str) -> tuple[int | None, int | None, int | None, int | None]:
    """(start_month, start_year, end_month, end_year). For Dec–Jan statements, used to assign year by transaction month."""
    lines = text.split("\n")
    head = " ".join(lines[:80])
    m = RE_PERIOD_RANGE.search(head)
    if m:
        sm, sy = int(m.group(1)), int(m.group(3))
        em, ey = int(m.group(4)), int(m.group(6))
        if 1990 <= sy <= 2030 and 1990 <= ey <= 2030:
            return (sm, sy, em, ey)
    for pat in (RE_THROUGH_FULL, RE_AS_OF_FULL):
        m = pat.search(head)
        if m:
            em, ey = int(m.group(1)), int(m.group(3))
            if 1990 <= ey <= 2030:
                return (None, None, em, ey)
    return (None, None, None, None)


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


def _extract_statement_year(text: str) -> int | None:
    """Retrieve statement year from header (through, as of, period range, or month name + year)."""
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


# Phrases that indicate credit (money in). Order matters: check specific first.
CITI_CREDIT_PHRASES = (
    "ELECTRONIC CREDIT",
    "MISC DEPOSIT",
    "ATM DEPOSIT",
    "CHECK REVERSAL",
    "SERV CHARGE REV",
)


def _is_credit(description: str) -> bool:
    """True if transaction is a credit (money in). Debits (payments, purchases) are negative."""
    u = description.upper()
    for phrase in CITI_CREDIT_PHRASES:
        if phrase in u:
            return True
    return False


def _parse_citi_multiline(lines: list, in_section_start: int) -> list[dict]:
    """Parse when each transaction is split: line1=date, line2=desc, line3=amount, line4=balance, then optional continuation lines. Keeps balance for debit/credit from balance column."""
    rows = []
    i = in_section_start
    while i < len(lines):
        line = lines[i]
        if not line:
            i += 1
            continue
        if "Total Debits/Credits" in line:
            break
        date_m = CITI_DATE_LINE.match(line)
        if date_m:
            date_str = date_m.group(1)
            i += 1
            desc_parts = []
            amt_str = None
            balance_str = None
            while i < len(lines):
                cur = lines[i]
                if not cur:
                    i += 1
                    continue
                if CITI_DATE_LINE.match(cur):
                    break
                amt_clean = cur.replace(",", "").strip().rstrip("-")
                if re.match(r"^\d+\.\d{2}$", amt_clean) or re.match(r"^[\d,]+\.\d{2}-?\s*$", cur.strip()):
                    if amt_str is None:
                        amt_str = cur.replace(",", "").rstrip("-").strip()
                        i += 1
                        if i < len(lines) and re.match(r"^[\d,]+\.\d{2}-?\s*$", lines[i].strip()):
                            balance_str = lines[i].strip().replace(",", "").rstrip("-").strip()
                            i += 1
                        while i < len(lines):
                            cont = lines[i].strip()
                            if not cont or CITI_DATE_LINE.match(cont) or re.match(r"^[\d,]+\.\d{2}-?\s*$", cont):
                                break
                            if "Total Debits/Credits" in cont or (cont.startswith("-- ") and " of " in cont):
                                break
                            desc_parts.append(cont)
                            i += 1
                        break
                desc_parts.append(cur)
                i += 1
            if amt_str and desc_parts:
                try:
                    amount_raw = abs(float(amt_str.replace(",", "")))
                    balance_val = _parse_balance_value(balance_str) if balance_str else None
                    full_desc = " ".join(desc_parts).strip()
                    row = {"date": date_str, "description": full_desc, "amount": amount_raw}
                    if balance_val is not None:
                        row["balance"] = balance_val
                    rows.append(row)
                except ValueError:
                    pass
            continue
        i += 1
    return rows


def _parse_citi_singleline(lines: list, in_section_start: int) -> list[dict]:
    """Parse when each transaction is one line: MM/DD description amount balance. Keeps balance for debit/credit from balance column."""
    rows = []
    i = in_section_start
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if "Total Debits/Credits" in line:
            break
        m = CITI_TX_LINE_SINGLE.match(line)
        if m:
            date_str, desc, amt_str, balance_str = m.groups()
            amt_clean = amt_str.replace(",", "")
            try:
                amount_raw = abs(float(amt_clean))
            except ValueError:
                i += 1
                continue
            balance_val = _parse_balance_value(balance_str) if balance_str else None
            desc_parts = [desc.strip()]
            i += 1
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line:
                    i += 1
                    continue
                if CITI_TX_LINE_SINGLE.match(next_line):
                    break
                if next_line.startswith("-- ") and " of " in next_line:
                    break
                if "Page " in next_line and " of " in next_line:
                    i += 1
                    continue
                desc_parts.append(next_line)
                i += 1
            full_desc = " ".join(desc_parts).strip()
            row = {"date": date_str, "description": full_desc, "amount": amount_raw}
            if balance_val is not None:
                row["balance"] = balance_val
            rows.append(row)
            continue
        i += 1
    return rows


def parse_citi_checking_text(text: str) -> list[dict]:
    """
    Parse Citi checking/business statement text.
    Returns list of {"date": "dd/mm/yyyy", "description": "...", "amount": float}.
    Amount: negative = debit (money out), positive = credit (money in).
    Uses beginning balance + balance column to determine debit vs credit (balance down = debit, up = credit).
    Falls back to description keywords (ELECTRONIC CREDIT etc.) if no balance available.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if "CHECKING ACTIVITY" in line:
            i += 1
            start = i
            while start < len(lines):
                if CITI_DATE_LINE.match(lines[start]):
                    rows = _parse_citi_multiline(lines, start)
                    break
                if CITI_TX_LINE_SINGLE.match(lines[start]):
                    rows = _parse_citi_singleline(lines, start)
                    break
                start += 1
            break
        i += 1
    beginning_balance = _extract_beginning_balance(text)
    if beginning_balance is not None and any(r.get("balance") is not None for r in rows):
        _apply_balance_to_sign(rows, beginning_balance)
    else:
        for row in rows:
            raw = abs(float(row.get("amount", 0) or 0))
            if _is_credit(row.get("description", "")):
                row["amount"] = raw
            else:
                row["amount"] = -raw
    period = _extract_statement_period(text)
    year = _extract_statement_year(text)
    for row in rows:
        row["date"] = _normalize_date_to_mm_dd_yyyy(row.get("date") or "", year, period)
    return rows


def parse_citi_checking_pdf(pdf_path: str) -> list[dict]:
    """Parse Citi checking statement PDF. Returns list of {date, description, amount}."""
    text = extract_text_from_pdf(pdf_path)
    return parse_citi_checking_text(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_citi_checking.py <pdf_path> [--csv output.csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    out_csv = None
    if len(sys.argv) >= 4 and sys.argv[2] == "--csv":
        out_csv = sys.argv[3]
    rows = parse_citi_checking_pdf(pdf_path)
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
