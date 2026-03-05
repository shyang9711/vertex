"""
Parse U.S. Bank Business Credit Card statement PDF and pasted text.
Output: list of dicts with date, description, amount (negative = purchases/fees, positive = credits), reference_number.

PDF format (Activity Summary): Post Date, Trans Date, Ref #, Transaction Description, Amount.
Sections: Payments and Other Credits, Purchases and Other Debits, Fees.
"""
import re
import sys
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


# Statement period: "09/19/2025 - 10/17/2025" or "Open Date: 09/19/2025 Closing Date: 10/17/2025"
RE_OPEN_CLOSING = re.compile(
    r"(?:Open\s+Date:\s*)?(\d{1,2})/(\d{1,2})/(\d{4}).*?(?:Closing\s+Date:\s*)?(\d{1,2})/(\d{1,2})/(\d{4})",
    re.I
)
RE_PERIOD_DASH = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})\s*-\s*(\d{1,2})/(\d{1,2})/(\d{4})")

# Amount line: $1,234.56 or $195.00
RE_AMOUNT_LINE = re.compile(r"^\s*\$?([\d,]+\.\d{2})\s*$")

# Post/Trans date: MM/DD
RE_MM_DD = re.compile(r"^(\d{1,2})/(\d{1,2})\s*$")

# Ref #: digits only (4+ digits)
RE_REF = re.compile(r"^\d{4,}\s*$")


def _extract_statement_year(text: str) -> int | None:
    """Get statement year from Open/Closing date or period range."""
    for pat in (RE_OPEN_CLOSING, RE_PERIOD_DASH):
        m = pat.search(text)
        if m:
            y = int(m.group(3)) if pat.groups >= 3 else int(m.group(6))
            if 1990 <= y <= 2030:
                return y
    for m in re.finditer(r"\b(20\d{2})\b", text[:3000]):
        y = int(m.group(1))
        if 1990 <= y <= 2030:
            return y
    return None


def _normalize_date(mm_dd: str, year: int | None) -> str:
    """Convert MM/DD to MM/DD/YYYY."""
    if not year or not mm_dd or not mm_dd.strip():
        return ""
    m = RE_MM_DD.match(mm_dd.strip())
    if not m:
        return ""
    mo, d = int(m.group(1)), int(m.group(2))
    if 1 <= mo <= 12 and 1 <= d <= 31:
        return f"{mo:02d}/{d:02d}/{year}"
    return ""


def _parse_us_bank_cc_text(text: str) -> list[dict]:
    """
    Parse US Bank credit card statement text.
    Sections: Payments and Other Credits (positive), Purchases and Other Debits (negative), Fees (negative).
    Transaction block: Post Date (MM/DD), Trans Date (MM/DD), Ref #, Description lines..., Amount ($X.XX).
    """
    lines = [ln.strip() for ln in text.split("\n")]
    year = _extract_statement_year(text)
    if not year:
        year = 2025  # fallback

    # Section boundaries: for each amount line, use the section whose header appears last before that line
    section_starts = []  # list of (index, "credits"|"purchases"|"fees")
    for idx, line in enumerate(lines):
        lower = (line or "").lower()
        if "payments and other credits" in lower:
            section_starts.append((idx, "credits"))
        if "purchases and other debits" in lower:
            section_starts.append((idx, "purchases"))
        if "total fees this period" in lower or ("fees" in lower and "post" in lower and "date" in lower and idx > 50):
            section_starts.append((idx, "fees"))

    def _section_at(line_index: int) -> str | None:
        best = None
        best_idx = -1
        for idx, sec in section_starts:
            if idx < line_index and idx > best_idx:
                best_idx = idx
                best = sec
        return best

    rows = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line or not line.strip():
            i += 1
            continue

        # Check for amount-only line (end of a transaction block). Skip total rows.
        prev_non_empty = next((lines[k] for k in range(i - 1, max(-1, i - 5), -1) if lines[k].strip()), "")
        if prev_non_empty and "total this period" in prev_non_empty.lower():
            i += 1
            continue
        if prev_non_empty and "total fees this period" in prev_non_empty.lower():
            i += 1
            continue
        amt_m = RE_AMOUNT_LINE.match(line.replace("$", "").replace(",", ""))
        current_section = _section_at(i)
        if amt_m and current_section:
            try:
                amount = float(amt_m.group(1).replace(",", ""))
            except ValueError:
                i += 1
                continue
            # Work backwards to find ref, trans date, post date, description
            j = i - 1
            desc_parts = []
            ref_num = ""
            trans_date = ""
            post_date = ""
            while j >= 0 and j >= i - 15:
                prev = lines[j].strip()
                if not prev:
                    j -= 1
                    continue
                if RE_AMOUNT_LINE.match(prev.replace("$", "").replace(",", "")):
                    break
                if RE_REF.match(prev):
                    if not ref_num:
                        ref_num = prev
                        j -= 1
                        continue
                if RE_MM_DD.match(prev):
                    if not trans_date:
                        trans_date = prev
                        j -= 1
                        continue
                    if not post_date:
                        post_date = prev
                        j -= 1
                        break
                if prev.upper() in ("TOTAL THIS PERIOD", "TOTAL FEES THIS PERIOD", "POST", "DATE", "TRANS", "REF #", "NOTATION"):
                    j -= 1
                    continue
                if "Transaction Description" in prev or "Amount" == prev:
                    j -= 1
                    continue
                desc_parts.insert(0, prev)
                j -= 1

            date_str = _normalize_date(trans_date or post_date, year)
            description = " ".join(desc_parts).strip() if desc_parts else "Unknown"
            # Sign: credits section or refund/return description = positive
            description_lower = description.lower()
            if current_section == "credits" or "return" in description_lower or "refund" in description_lower:
                amount_signed = abs(amount)
            else:
                amount_signed = -abs(amount)
            # Skip junk: need valid date; skip totals and interest table lines
            if not date_str:
                i += 1
                continue
            if any(x in description for x in ("Total ", "Total Interest", "Variable Interest", "Annual Percentage", "**BALANCE TRANSFER", "**PURCHASES", "**ADVANCES", "Expires with Statement")):
                i += 1
                continue
            if description == "Unknown" and not ref_num and amount_signed == 0:
                i += 1
                continue
            rows.append({
                "date": date_str,
                "description": description,
                "amount": amount_signed,
                "reference_number": ref_num
            })
        i += 1

    return rows


def parse_us_bank_cc_text(text: str) -> list[dict]:
    """Parse US Bank credit card statement text. Returns list of {date, description, amount, reference_number}."""
    return _parse_us_bank_cc_text(text)


def parse_us_bank_cc_pdf(pdf_path: str) -> list[dict]:
    """Parse US Bank credit card statement PDF."""
    text = extract_text_from_pdf(pdf_path)
    return parse_us_bank_cc_text(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_us_bank_cc.py <pdf_path> [--csv output.csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    rows = parse_us_bank_cc_pdf(pdf_path)
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
