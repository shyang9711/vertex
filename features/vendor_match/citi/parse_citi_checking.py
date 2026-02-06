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
    """Parse when each transaction is split: line1=date, line2=desc, line3=amount, line4=balance, then optional continuation lines."""
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
            while i < len(lines):
                cur = lines[i]
                if not cur:
                    i += 1
                    continue
                if CITI_DATE_LINE.match(cur):
                    break
                amt_clean = cur.replace(",", "").strip().rstrip("-")
                if re.match(r"^\d+\.\d{2}$", amt_clean) and amt_str is None:
                    amt_str = cur.replace(",", "").rstrip("-").strip()
                    i += 1
                    if i < len(lines) and re.match(r"^[\d,]+\.\d{2}-?\s*$", lines[i].strip()):
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
                    amount = float(amt_str.replace(",", ""))
                    if _is_credit(desc_parts[0]):
                        amount = abs(amount)
                    else:
                        amount = -abs(amount)
                    full_desc = " ".join(desc_parts).strip()
                    rows.append({"date": date_str, "description": full_desc, "amount": amount})
                except ValueError:
                    pass
            continue
        i += 1
    return rows


def _parse_citi_singleline(lines: list, in_section_start: int) -> list[dict]:
    """Parse when each transaction is one line: MM/DD description amount balance."""
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
            date_str, desc, amt_str, _balance = m.groups()
            amt_clean = amt_str.replace(",", "")
            try:
                amount = float(amt_clean)
            except ValueError:
                i += 1
                continue
            if _is_credit(desc):
                amount = abs(amount)
            else:
                amount = -abs(amount)
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
            rows.append({"date": date_str, "description": full_desc, "amount": amount})
            continue
        i += 1
    return rows


def parse_citi_checking_text(text: str) -> list[dict]:
    """
    Parse Citi checking/business statement text.
    Returns list of {"date": "MM/DD", "description": "...", "amount": float}.
    Amount: negative = debit (money out), positive = credit (money in).
    Handles both multi-line PDF layout (date, description, amount, balance on separate lines)
    and single-line layout (all on one line).
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
