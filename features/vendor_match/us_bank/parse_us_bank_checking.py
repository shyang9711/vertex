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
    r"(?:Statement\s+Period:\s*)?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),?\s+(\d{4})\s+through\s+",
    re.I
)
RE_MONTH_DAY = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s*$", re.I)
# Transaction line start: "Feb 14 Wire Credit..." or "Feb 18 Debit Purchase..."
RE_TX_START = re.compile(r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(.+)$", re.I)
RE_AMOUNT = re.compile(r"^\s*\$?([\d,]+)\.(\d{2})\s*(-)?\s*$")  # optional minus suffix for withdrawal
RE_AMOUNT_LEADING_MINUS = re.compile(r"^\s*-\s*\$?([\d,]+)\.(\d{2})\s*$")  # leading minus e.g. -2,700.00
RE_REF = re.compile(r"^\d{4,}\s*$")
RE_REF_IN_DESC = re.compile(r"REF\s*#?\s*(\d+)", re.I)


def _extract_statement_year(text: str) -> int | None:
    """Get statement year from Statement Period (use end date)."""
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


def _parse_layout2_other_deposits_withdrawals(lines: list, year: int) -> list[dict]:
    """
    Parse "Other Deposits", "Card Withdrawals", "Other Withdrawals" layout.
    Transaction: line starting with "Mon DD Description...", then continuation lines, then amount ($X.XX or X.XX-).
    """
    rows = []
    i = 0
    current_section = None  # "other_deposits" | "card_withdrawals" | "other_withdrawals"
    while i < len(lines):
        line = lines[i]
        lower = (line or "").lower()
        if "other deposits" in lower and "withdrawal" not in lower:
            current_section = "other_deposits"
            i += 1
            continue
        if "card withdrawals" in lower:
            current_section = "card_withdrawals"
            i += 1
            continue
        if "other withdrawals" in lower:
            current_section = "other_withdrawals"
            i += 1
            continue
        if not line.strip():
            i += 1
            continue

        # Transaction start: "Feb 14 Wire Credit REF019296" or "Feb 18 Debit Purchase - VISA"
        tx_start = RE_TX_START.match(line.strip())
        if tx_start and current_section:
            month_abbrev, day, rest = tx_start.groups()
            desc_parts = [rest.strip()]
            j = i + 1
            amount_val = None
            is_negative = False
            # Try amount at end of first line (e.g. "Sep 15 Wire Credit 3,880.23" or "Sep 15 Debit 2,700.00-")
            rest_clean = rest.replace(",", "")
            for amt_re in (RE_AMOUNT, RE_AMOUNT_LEADING_MINUS):
                amt_m = amt_re.search(" " + rest_clean)
                if amt_m:
                    try:
                        amount_val = float(amt_m.group(1).replace(",", "") + "." + amt_m.group(2))
                        is_negative = amt_re is RE_AMOUNT_LEADING_MINUS or (amt_re is RE_AMOUNT and (amt_m.lastindex >= 3 and amt_m.group(3) == "-"))
                        break
                    except (ValueError, IndexError):
                        pass
            while j < len(lines) and j < i + 25:
                nxt = lines[j].strip()
                if not nxt:
                    j += 1
                    continue
                # Amount: "85,000.00" or "83.74-" or "-2,700.00" or after a lone "$"
                nxt_clean = nxt.replace(",", "")
                amt_m = RE_AMOUNT.match(nxt_clean)
                if not amt_m:
                    amt_m = RE_AMOUNT_LEADING_MINUS.match(nxt_clean)
                    if amt_m:
                        amount_val = float(amt_m.group(1).replace(",", "") + "." + amt_m.group(2))
                        is_negative = True
                        j += 1
                        break
                if nxt == "$":
                    j += 1
                    if j < len(lines):
                        amt_m = RE_AMOUNT.match(lines[j].strip())
                        if amt_m:
                            amount_val = float(amt_m.group(1).replace(",", "") + "." + amt_m.group(2))
                            is_negative = amt_m.group(3) == "-"
                            j += 1
                    break
                if amt_m:
                    amount_val = float(amt_m.group(1).replace(",", "") + "." + amt_m.group(2))
                    is_negative = amt_m.group(3) == "-"
                    j += 1
                    break
                # Next transaction start?
                if RE_TX_START.match(nxt):
                    break
                if "Number of Days" in nxt or "To Contact" in nxt:
                    break
                desc_parts.append(nxt)
                j += 1

            if amount_val is not None and amount_val != 0:
                abs_val = abs(amount_val)
                date_str = _normalize_date(month_abbrev, day, year)
                description = " ".join(desc_parts).strip()
                ref_num = ""
                ref_m = RE_REF_IN_DESC.search(description)
                if ref_m:
                    ref_num = ref_m.group(1)
                # Infer sign from section or description (Wire Credit / Deposit = positive; Debit = negative)
                desc_upper = description.upper()
                if amount_val < 0:
                    amt_signed = -abs_val
                elif current_section == "other_deposits" or "WIRE CREDIT" in desc_upper or "ELECTRONIC DEPOSIT" in desc_upper or (desc_upper.startswith("DEPOSIT") and "DEBIT" not in desc_upper) or " RET " in desc_upper or " RETURN" in desc_upper or " REFUND" in desc_upper:
                    amt_signed = abs_val
                else:
                    amt_signed = -abs_val if is_negative else abs_val
                rows.append({
                    "date": date_str,
                    "description": description,
                    "amount": amt_signed,
                    "reference_number": ref_num
                })
            i = j
            continue
        i += 1
    return rows


def _parse_customer_deposits(lines: list, year: int) -> list[dict]:
    """Parse Customer Deposits section: Date line, optional Ref line, Amount line. Returns list of deposit dicts."""
    out = []
    in_deposits = False
    i = 0
    while i < len(lines):
        line = lines[i]
        lower = (line or "").lower()
        if "customer deposits" in lower:
            in_deposits = True
            i += 1
            continue
        if "withdrawal" in lower and "total" not in lower:
            in_deposits = False
            i += 1
            continue
        if not line or not in_deposits:
            i += 1
            continue
        date_m = RE_MONTH_DAY.match(line)
        if date_m:
            month_abbrev, day = date_m.groups()
            j = i + 1
            ref_num = ""
            amount_val = None
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
                if not amt:
                    amt = RE_AMOUNT_LEADING_MINUS.match(nxt.replace(",", ""))
                if amt and not amount_val:
                    try:
                        amount_val = float(amt.group(1).replace(",", "") + "." + amt.group(2))
                        if amount_val > 0 and amount_val < 1e9:
                            date_str = _normalize_date(month_abbrev, day, year)
                            out.append({
                                "date": date_str,
                                "description": f"Deposit {ref_num}" if ref_num else "Customer Deposit",
                                "amount": amount_val,
                                "reference_number": ref_num
                            })
                            i = j
                            break
                    except (ValueError, IndexError):
                        pass
                if RE_MONTH_DAY.match(nxt) and nxt != line:
                    break
                j += 1
        i += 1
    return out


def _parse_checks_sections(lines: list, year: int) -> list[dict]:
    """
    Parse "Checks Presented Conventionally" and "Checks Presented Electronically" sections.
    Each line typically has check info and payee then amount. Include payee in description for vendor matching.
    """
    rows = []
    current_section = None
    i = 0
    while i < len(lines):
        line = lines[i]
        lower = (line or "").lower()
        if "checks presented conventionally" in lower:
            current_section = "conventional"
            i += 1
            continue
        if "checks presented electronically" in lower:
            current_section = "electronic"
            i += 1
            continue
        if "other deposits" in lower or "card withdrawals" in lower or "other withdrawals" in lower:
            current_section = None
            i += 1
            continue
        if not line.strip() or not current_section:
            i += 1
            continue
        stripped = line.strip().replace(",", "")
        amt_m = RE_AMOUNT.search(" " + stripped) or RE_AMOUNT_LEADING_MINUS.search(" " + stripped)
        if amt_m:
            try:
                amount_val = float(amt_m.group(1).replace(",", "") + "." + amt_m.group(2))
                if amount_val > 0 and amount_val < 1e9:
                    line_strip = line.strip()
                    tx_start = RE_TX_START.match(line_strip)
                    if tx_start:
                        month_abbrev, day, rest = tx_start.groups()
                        date_str = _normalize_date(month_abbrev, day, year)
                        rest_clean = rest.replace(",", "")
                        am2 = RE_AMOUNT.search(" " + rest_clean) or RE_AMOUNT_LEADING_MINUS.search(" " + rest_clean)
                        description = (rest_clean[:am2.start()].strip() if am2 else rest.strip()).replace("  ", " ")
                    else:
                        date_str = ""
                        end_amt = re.search(r"[\d,]+\.\d{2}\s*-?\s*$", line_strip)
                        description = line_strip[:end_amt.start()].strip() if end_amt else line_strip
                    rows.append({
                        "date": date_str,
                        "description": description or f"Check {current_section}",
                        "amount": -amount_val,
                        "reference_number": ""
                    })
            except (ValueError, IndexError, AttributeError):
                pass
        i += 1
    return rows


def _parse_us_bank_checking_text(text: str) -> list[dict]:
    """
    Parse US Bank checking statement text.
    Layout 2 (Other Deposits / Card Withdrawals / Other Withdrawals) first, then Layout 1 (Customer Deposits) always merged.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    year = _extract_statement_year(text)
    if not year:
        year = datetime.now().year

    rows = []

    # Layout 2: "Other Deposits", "Card Withdrawals", "Other Withdrawals" with "Mon DD Desc..." then amount
    rows = _parse_layout2_other_deposits_withdrawals(lines, year)

    # Layout 1: "Customer Deposits" — always merge (so Customer Deposit 5,272.00 etc. are included)
    rows.extend(_parse_customer_deposits(lines, year))

    # Checks Presented Conventionally / Electronically (payee in description for vendor match)
    rows.extend(_parse_checks_sections(lines, year))

    # Alternate: "Jan 15" then ref then amount on consecutive lines (no Customer Deposits header)
    if not rows:
        i = 0
        while i < len(lines):
            date_m = RE_MONTH_DAY.match(lines[i].strip())
            if date_m:
                month_abbrev, day = date_m.groups()
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
                                    amount_val = float(amt.group(1) + "." + amt.group(2))
                                except ValueError:
                                    pass
                                j += 1
                        break
                    amt = RE_AMOUNT.match(nxt.replace(",", ""))
                    if amt:
                        try:
                            amount_val = float(amt.group(1) + "." + amt.group(2))
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
