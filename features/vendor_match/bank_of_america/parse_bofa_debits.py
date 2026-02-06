"""
Parse a Bank of America statement PDF and output only withdrawals/debits (bank) or
purchases/charges (credit card). Output: date, description, amount.

Differences:
- BoA BANK (eStatement): "Withdrawals and other debits" | date MM/DD/YY | amount negative
- BoA CREDIT CARD:       "Purchases and Other Charges" | dates MM/DD (post + trans) | amount positive

Usage:
  python parse_bofa_debits.py "path/to/statement.pdf"
  python parse_bofa_debits.py "path/to/statement.pdf" --csv "output.csv"
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


def _is_bank_statement(text: str) -> bool:
    """True if PDF is a bank (checking) eStatement, False if credit card."""
    # Bank eStatement has "Withdrawals and other debits"; CC has "Purchases and Other Charges"
    return "Withdrawals and other debits" in text


# --- Bank eStatement: "Withdrawals and other debits" ---
# Line 1: date (MM/DD/YY), Line 2+: description, Last: amount (-1,234.56)
DATE_ONLY = re.compile(r"^(\d{1,2}/\d{1,2}/\d{2})$")
AMOUNT_LINE = re.compile(r"^(-?\d{1,3}(?:,\d{3})*\.\d{2})$")

# --- Credit Card: "Purchases and Other Charges" ---
# 5 lines per transaction: post_date (MM/DD), trans_date (MM/DD), description, reference, amount (positive = charge)
CC_DATE = re.compile(r"^(\d{1,2}/\d{1,2})$")
CC_REF = re.compile(r"^\d{15,}$")
CC_AMOUNT = re.compile(r"^(-?\s*\d{1,3}(?:,\d{3})*\.\d{2})$")
# Single-line format (e.g. November statement): PostDate TransDate Description RefNumber Amount
CC_SINGLE_LINE = re.compile(
    r"^(\d{1,2}/\d{1,2})\s+(\d{1,2}/\d{1,2})\s+(.+?)\s+(\d{15,})\s+([-]?\s*[\d,]+\.?\d{0,2})\s*$"
)


def _parse_bank_withdrawals(text: str) -> list[dict]:
    """
    Parse Bank of America eStatement text. Returns withdrawals/other debits.
    Each item: {"date": "MM/DD/YY", "description": "...", "amount": float (negative)}.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    in_section = False
    past_header = False
    rows = []
    cur_date = None
    cur_desc_lines = []

    for line in lines:
        if not line:
            continue
        if "Withdrawals and other debits" in line and "Date" not in line:
            in_section = True
            past_header = False
            continue
        if not in_section:
            continue
        if line in ("Date", "Description", "Amount"):
            past_header = True
            continue
        if "continued on the next page" in line or ("Deposits and other credits" in line and "Withdrawals" not in line):
            in_section = False
            continue
        amt_match = AMOUNT_LINE.match(line)
        if amt_match and past_header and cur_date is not None:
            amount = float(amt_match.group(1).replace(",", ""))
            if amount < 0:
                description = " ".join(cur_desc_lines).strip() if cur_desc_lines else ""
                rows.append({"date": cur_date, "description": description, "amount": amount})
            cur_date = None
            cur_desc_lines = []
            continue
        date_match = DATE_ONLY.match(line)
        if date_match and past_header:
            cur_date = date_match.group(1)
            cur_desc_lines = []
            continue
        if cur_date is not None and not AMOUNT_LINE.match(line):
            cur_desc_lines.append(line)
    return rows


def _parse_bank_deposits(text: str) -> list[dict]:
    """
    Parse Bank of America eStatement text. Returns deposits/other credits.
    Each item: {"date": "MM/DD/YY", "description": "...", "amount": float (positive)}.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    in_section = False
    past_header = False
    rows = []
    cur_date = None
    cur_desc_lines = []

    for line in lines:
        if not line:
            continue
        if "Deposits and other credits" in line and "Date" not in line:
            in_section = True
            past_header = False
            continue
        if not in_section:
            continue
        if line in ("Date", "Description", "Amount"):
            past_header = True
            continue
        if "continued on the next page" in line or ("Withdrawals and other debits" in line and "Deposits" not in line):
            in_section = False
            continue
        amt_match = AMOUNT_LINE.match(line)
        if amt_match and past_header and cur_date is not None:
            amount = float(amt_match.group(1).replace(",", ""))
            if amount > 0:
                description = " ".join(cur_desc_lines).strip() if cur_desc_lines else ""
                rows.append({"date": cur_date, "description": description, "amount": amount})
            cur_date = None
            cur_desc_lines = []
            continue
        date_match = DATE_ONLY.match(line)
        if date_match and past_header:
            cur_date = date_match.group(1)
            cur_desc_lines = []
            continue
        if cur_date is not None and not AMOUNT_LINE.match(line):
            cur_desc_lines.append(line)
    return rows


def _parse_cc_charges_single_line(text: str) -> list[dict]:
    """
    Parse BoA Credit Card when each transaction is on one line:
    PostDate TransDate Description ReferenceNumber Amount
    (e.g. November statement / table layout). Returns only positive amounts (purchases/charges).
    """
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    in_section = False
    for line in lines:
        if not line:
            continue
        if "Purchases and Other Charges" in line and "TOTAL" not in line:
            in_section = True
            continue
        if "TOTAL PURCHASES AND OTHER CHARGES FOR THIS PERIOD" in line:
            in_section = False
            continue
        if not in_section:
            continue
        m = CC_SINGLE_LINE.match(line)
        if m:
            _post, _trans, desc, _ref, amt_str = m.groups()
            amt_clean = amt_str.replace(",", "").replace(" ", "").strip()
            is_neg = amt_clean.startswith("-") or amt_str.strip().startswith("-")
            if amt_clean.startswith("-"):
                amt_clean = amt_clean[1:]
            try:
                amount = float(amt_clean)
                if is_neg:
                    amount = -amount
                if amount > 0:
                    rows.append({"date": _post, "description": desc.strip(), "amount": amount})
            except ValueError:
                pass
    return rows


def _parse_cc_charges(text: str) -> list[dict]:
    """
    Parse Bank of America Credit Card statement text. Returns purchases/other charges.
    Each item: {"date": "MM/DD" (posting), "description": "...", "amount": float (positive)}.
    Supports: (1) single-line table format; (2) 5 lines per transaction — post_date, trans_date, description, reference, amount.
    """
    single_line_rows = _parse_cc_charges_single_line(text)
    if single_line_rows:
        return single_line_rows
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    in_section = False
    # State: 0=want post_date, 1=want trans_date, 2=want description, 3=want ref, 4=want amount
    state = 0
    post_date = None
    trans_date = None
    description = None

    i = 0
    while i < len(lines):
        line = lines[i]
        if not line:
            i += 1
            continue

        if "Purchases and Other Charges" in line and "TOTAL" not in line:
            in_section = True
            state = 0
            i += 1
            continue
        if "TOTAL PURCHASES AND OTHER CHARGES FOR THIS PERIOD" in line:
            in_section = False
            state = 0
            i += 1
            continue
        if not in_section:
            i += 1
            continue

        if state == 0 and CC_DATE.match(line):
            post_date = CC_DATE.match(line).group(1)
            state = 1
            i += 1
            continue
        if state == 1 and CC_DATE.match(line):
            trans_date = CC_DATE.match(line).group(1)
            state = 2
            i += 1
            continue
        if state == 2:
            description = line
            state = 3
            i += 1
            continue
        if state == 3 and CC_REF.match(line):
            state = 4
            i += 1
            continue
        if state == 4:
            amt_match = CC_AMOUNT.match(line)
            if amt_match:
                raw = amt_match.group(1).replace(",", "").replace(" ", "").strip()
                amount = float(raw) if not raw.startswith("-") else -float(raw[1:])
                if amount > 0:
                    rows.append({
                        "date": post_date or trans_date or "",
                        "description": description or "",
                        "amount": amount,
                    })
            state = 0
            post_date = trans_date = description = None
            i += 1
            continue

        # Reset on unexpected line
        state = 0
        i += 1
    return rows


def _parse_cc_payments_single_line(text: str) -> list[dict]:
    """
    Parse BoA Credit Card "Payments and Other Credits" when each transaction is on one line.
    Returns rows with negative amount (credit to card).
    """
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    in_section = False
    for line in lines:
        if not line:
            continue
        if "Payments and Other Credits" in line and "TOTAL" not in line:
            in_section = True
            continue
        if "TOTAL PAYMENTS AND CREDITS" in line or "TOTAL PAYMENTS AND OTHER CREDITS" in line:
            in_section = False
            continue
        if not in_section:
            continue
        m = CC_SINGLE_LINE.match(line)
        if m:
            _post, _trans, desc, _ref, amt_str = m.groups()
            amt_clean = amt_str.replace(",", "").replace(" ", "").strip()
            is_neg = amt_clean.startswith("-") or amt_str.strip().startswith("-")
            if amt_clean.startswith("-"):
                amt_clean = amt_clean[1:]
            try:
                amount = float(amt_clean)
                if is_neg:
                    amount = -amount
                if amount < 0:
                    rows.append({"date": _post, "description": desc.strip(), "amount": amount})
            except ValueError:
                pass
    return rows


def _parse_cc_payments(text: str) -> list[dict]:
    """
    Parse BoA Credit Card "Payments and Other Credits". Returns rows with negative amount.
    """
    single_line_rows = _parse_cc_payments_single_line(text)
    if single_line_rows:
        return single_line_rows
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    in_section = False
    state = 0
    post_date = None
    trans_date = None
    description = None
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line:
            i += 1
            continue
        if "Payments and Other Credits" in line and "TOTAL" not in line:
            in_section = True
            state = 0
            i += 1
            continue
        if "TOTAL PAYMENTS AND CREDITS" in line or "TOTAL PAYMENTS AND OTHER CREDITS" in line:
            in_section = False
            state = 0
            i += 1
            continue
        if not in_section:
            i += 1
            continue
        if state == 0 and CC_DATE.match(line):
            post_date = CC_DATE.match(line).group(1)
            state = 1
            i += 1
            continue
        if state == 1 and CC_DATE.match(line):
            trans_date = CC_DATE.match(line).group(1)
            state = 2
            i += 1
            continue
        if state == 2:
            description = line
            state = 3
            i += 1
            continue
        if state == 3 and CC_REF.match(line):
            state = 4
            i += 1
            continue
        if state == 4:
            amt_match = CC_AMOUNT.match(line)
            if amt_match:
                raw = amt_match.group(1).replace(",", "").replace(" ", "").strip()
                amount = float(raw) if not raw.startswith("-") else -float(raw[1:])
                if amount < 0:
                    rows.append({
                        "date": post_date or trans_date or "",
                        "description": description or "",
                        "amount": amount,
                    })
            state = 0
            post_date = trans_date = description = None
            i += 1
            continue
        state = 0
        i += 1
    return rows


def parse_bofa_text_to_rows(text: str) -> tuple[list[dict], str | None]:
    """
    Parse Bank of America statement text (eStatement bank or credit card).
    Returns (rows, "bank"|"credit_card"|None). Rows have keys: date, description, amount.
    Use this when you already have extracted PDF text (e.g. from vendor_match).
    """
    if "Withdrawals and other debits" in text:
        return _parse_bank_withdrawals(text), "bank"
    if "Purchases and Other Charges" in text:
        return _parse_cc_charges(text), "credit_card"
    return [], None


def parse_bofa_withdrawals_only(pdf_path: str) -> tuple[list[dict], str]:
    """
    Parse Bank of America PDF (bank eStatement or credit card). Auto-detects type.
    Returns (rows, statement_type) where statement_type is "bank" or "credit_card".
    Rows: [{"date", "description", "amount"}]. Bank amounts are negative; CC amounts are positive.
    """
    text = extract_text_from_pdf(pdf_path)
    if _is_bank_statement(text):
        return _parse_bank_withdrawals(text), "bank"
    return _parse_cc_charges(text), "credit_card"


def parse_bofa_bank_only(text: str) -> list[dict]:
    """Parse BoA bank eStatement text only (withdrawals/debits). Ignores credit card format."""
    return _parse_bank_withdrawals(text)


def parse_bofa_bank_credits_only(text: str) -> list[dict]:
    """Parse BoA bank eStatement text only (deposits and other credits)."""
    return _parse_bank_deposits(text)


def parse_bofa_bank_both(text: str) -> list[dict]:
    """Parse BoA bank eStatement text: both withdrawals (debits) and deposits (credits)."""
    return _parse_bank_withdrawals(text) + _parse_bank_deposits(text)


def parse_bofa_cc_only(text: str) -> list[dict]:
    """Parse BoA credit card statement text only (purchases and other charges). Ignores bank format."""
    return _parse_cc_charges(text)


def parse_bofa_cc_credits_only(text: str) -> list[dict]:
    """Parse BoA credit card statement text only (payments and other credits)."""
    return _parse_cc_payments(text)


def parse_bofa_cc_both(text: str) -> list[dict]:
    """Parse BoA credit card statement text: both purchases (debits) and payments (credits)."""
    return _parse_cc_charges(text) + _parse_cc_payments(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_bofa_debits.py <pdf_path> [--csv output.csv]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    out_csv = None
    if len(sys.argv) >= 4 and sys.argv[2] == "--csv":
        out_csv = sys.argv[3]

    rows, statement_type = parse_bofa_withdrawals_only(pdf_path)

    if not rows:
        print("No withdrawals/debits (bank) or purchases/charges (CC) found in the PDF.")
        sys.exit(0)

    label = "BoA Bank (withdrawals/debits)" if statement_type == "bank" else "BoA Credit Card (purchases/charges)"
    print(f"Detected: {label}")
    print()
    col_widths = (12, 50, 12)
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
