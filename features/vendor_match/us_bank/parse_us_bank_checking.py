"""
Parse U.S. Bank Business Checking statement PDF and pasted text.
Output: list of dicts with date, description, amount (negative = withdrawals, positive = deposits), reference_number.

PDF format: Statement Period Jan 7, 2025 through Jan 31, 2025.
Sections: Customer Deposits (Date, Ref Number, Amount, Date, Ending Balance), and withdrawal sections if present.
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


# Statement period: "Statement Period: Jan 7, 2025 through Jan 31, 2025" or "Jan 7, 2025 through Jan 31, 2025"
RE_STATEMENT_PERIOD = re.compile(
    r"(?:Statement\s+Period:\s*)?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),?\s+(\d{4})\s+through\s+",
    re.I
)
RE_MONTH_DAY = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s*$", re.I)
RE_AMOUNT = re.compile(r"^\s*\$?([\d,]+\.\d{2})\s*$")
RE_REF = re.compile(r"^\d{4,}\s*$")


def _extract_statement_year(text: str) -> int | None:
    """Get statement year from Statement Period."""
    m = RE_STATEMENT_PERIOD.search(text)
    if m:
        y = int(m.group(3))
        if 1990 <= y <= 2030:
            return y
    for m in re.finditer(r"\b(20\d{2})\b", text[:2000]):
        y = int(m.group(1))
        if 1990 <= y <= 2030:
            return y
    return None


def _normalize_date(month_abbrev: str, day: str, year: int) -> str:
    """Return MM/DD/YYYY."""
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


def _parse_us_bank_checking_text(text: str) -> list[dict]:
    """
    Parse US Bank checking statement text.
    Customer Deposits: rows like "Jan 15", "8612700360", "100.00", "Jan 15", "100.00" (Date, Ref Number, Amount, Date, Balance).
    Withdrawals/checks sections may follow similar pattern with negative amounts or separate columns.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    year = _extract_statement_year(text)
    if not year:
        year = datetime.now().year

    rows = []
    in_deposits = False
    in_withdrawals = False
    i = 0

    while i < len(lines):
        line = lines[i]
        lower = (line or "").lower()
        if "customer deposits" in lower:
            in_deposits = True
            in_withdrawals = False
            i += 1
            continue
        if "withdrawal" in lower and "total" not in lower or "checks paid" in lower or "debit" in lower:
            in_withdrawals = True
            in_deposits = False
            i += 1
            continue
        if "account summary" in lower or "balance summary" in lower:
            i += 1
            continue
        if not line:
            i += 1
            continue

        # Customer Deposits: after "# Items" and "Ending Balance" and "Number of Days", we get rows: Mon DD, ref, amount, date, balance
        if in_deposits:
            date_m = RE_MONTH_DAY.match(line)
            if date_m:
                month_abbrev, day = date_m.groups()
                # Look ahead for ref number and amount
                j = i + 1
                ref_num = ""
                amount_str = None
                while j < len(lines) and j < i + 6:
                    nxt = lines[j].strip()
                    if not nxt:
                        j += 1
                        continue
                    if RE_REF.match(nxt) and not ref_num:
                        ref_num = nxt
                        j += 1
                        continue
                    amt = RE_AMOUNT.match(nxt.replace(",", ""))
                    if amt and amount_str is None and ref_num:  # require ref to avoid adding balance row
                        try:
                            amount_val = float(amt.group(1).replace(",", ""))
                            if amount_val > 0 and amount_val < 1e9:  # plausible deposit
                                date_str = _normalize_date(month_abbrev, day, year)
                                rows.append({
                                    "date": date_str,
                                    "description": f"Deposit {ref_num}" if ref_num else "Deposit",
                                    "amount": amount_val,
                                    "reference_number": ref_num
                                })
                                i = j
                                break
                        except ValueError:
                            pass
                    if RE_MONTH_DAY.match(nxt) and nxt != line:
                        break
                    j += 1
        i += 1

    # Also try alternate layout: "Jan 15" then "8612700360" then "100.00" on consecutive lines (no header row in between)
    if not rows:
        i = 0
        while i < len(lines):
            date_m = RE_MONTH_DAY.match(lines[i].strip())
            if date_m:
                month_abbrev, day = date_m.groups()
                # Next non-empty: ref or amount
                j = i + 1
                ref_num = ""
                amount_val = None
                while j < len(lines) and j < i + 5:
                    nxt = lines[j].strip()
                    if not nxt:
                        j += 1
                        continue
                    if RE_REF.match(nxt):
                        ref_num = nxt
                        j += 1
                        if j < len(lines):
                            amt = RE_AMOUNT.match(lines[j].strip().replace(",", ""))
                            if amt:
                                try:
                                    amount_val = float(amt.group(1).replace(",", ""))
                                except ValueError:
                                    pass
                                j += 1
                        break
                    amt = RE_AMOUNT.match(nxt.replace(",", ""))
                    if amt:
                        try:
                            amount_val = float(amt.group(1).replace(",", ""))
                        except ValueError:
                            pass
                        j += 1
                        break
                    break
                if amount_val is not None and 0 < amount_val < 1e9:
                    date_str = _normalize_date(month_abbrev, day, year)
                    rows.append({
                        "date": date_str,
                        "description": f"Deposit {ref_num}" if ref_num else "Deposit",
                        "amount": amount_val,
                        "reference_number": ref_num
                    })
                i = j
            else:
                i += 1

    return rows


def parse_us_bank_checking_text(text: str) -> list[dict]:
    """Parse US Bank checking statement text. Returns list of {date, description, amount, reference_number}."""
    return _parse_us_bank_checking_text(text)


def parse_us_bank_checking_pdf(pdf_path: str) -> list[dict]:
    """Parse US Bank checking statement PDF."""
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
