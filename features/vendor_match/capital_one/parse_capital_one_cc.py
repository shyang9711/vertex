"""
Parse Capital One business credit card statement PDF and pasted text.
Output: list of dicts with date, description, amount (negative = purchases/fees/interest, positive = payments/credits).

PDF format: Trans Date, Post Date, Description, Amount per transaction.
Sections: Payments/Credits and Adjustments, Transactions (per cardholder), Fees, Interest Charged.
"""
import re
import sys
import os

MONTH_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
MONTH_NAMES = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

RE_BILLING_PERIOD = re.compile(
    r"([A-Z][a-z]{2})\s+(\d{1,2}),\s+(\d{4})\s*-\s*([A-Z][a-z]{2})\s+(\d{1,2}),\s+(\d{4})"
)
RE_MON_DAY = re.compile(r"^([A-Za-z]{3})\s+(\d{1,2})$")
RE_AMOUNT = re.compile(r"^\s*(-)?\s*\$([\d,]+\.\d{2})\s*$")
RE_CARDHOLDER_SECTION = re.compile(
    r"^.+?#\d{4}:\s*(Payments, Credits and Adjustments|Transactions)\s*$",
    re.I,
)

SKIP_DESCRIPTION_PREFIXES = (
    "tk#:", "orig:", "dest:", "psgr:", "carrier:", "svc:", "s/o:",
)

SKIP_LINES = {
    "trans date", "post date", "description", "amount",
    "transactions", "transactions (continued)",
    "visit capitalone.com to see detailed transactions.",
    "fees", "interest charged",
}


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


def _extract_statement_period(text: str) -> tuple[int, int, int, int] | None:
    """Return (start_month, start_year, end_month, end_year)."""
    m = RE_BILLING_PERIOD.search(text)
    if not m:
        return None
    sm = MONTH_ABBR.get(m.group(1).lower())
    sy = int(m.group(3))
    em = MONTH_ABBR.get(m.group(4).lower())
    ey = int(m.group(6))
    if not sm or not em:
        return None
    if not (1990 <= sy <= 2035 and 1990 <= ey <= 2035):
        return None
    return sm, sy, em, ey


def _closing_date(period: tuple[int, int, int, int] | None, text: str) -> str:
    if not period:
        return ""
    m = RE_BILLING_PERIOD.search(text)
    if not m:
        return ""
    end_month = MONTH_ABBR.get(m.group(4).lower())
    end_day = int(m.group(5))
    if not end_month:
        return ""
    return _normalize_mon_day(f"{MONTH_NAMES[end_month - 1]} {end_day}", period)


def _year_for_tx_month(tx_month: int, start_m: int, start_y: int, end_m: int, end_y: int) -> int:
    if end_y > start_y or (end_m < start_m and end_y == start_y):
        if tx_month >= start_m:
            return start_y
        return end_y
    return end_y


def _normalize_mon_day(mon_day: str, period: tuple[int, int, int, int] | None) -> str:
    m = RE_MON_DAY.match((mon_day or "").strip())
    if not m:
        return ""
    mo = MONTH_ABBR.get(m.group(1).lower())
    day = int(m.group(2))
    if not mo or not (1 <= day <= 31):
        return ""
    if period:
        start_m, start_y, end_m, end_y = period
        year = _year_for_tx_month(mo, start_m, start_y, end_m, end_y)
        return f"{mo:02d}/{day:02d}/{year}"
    return ""


def _parse_amount_line(line: str) -> tuple[float, bool] | None:
    """Return (absolute amount, is_credit) or None."""
    m = RE_AMOUNT.match((line or "").strip())
    if not m:
        return None
    try:
        amount = float(m.group(2).replace(",", ""))
    except ValueError:
        return None
    is_credit = bool(m.group(1))
    return amount, is_credit


def _is_skip_line(line: str) -> bool:
    lower = (line or "").strip().lower()
    if not lower:
        return True
    if lower in SKIP_LINES:
        return True
    if lower.startswith("total ") or lower.endswith(": total transactions"):
        return True
    if "total fees for this period" in lower or "total interest for this period" in lower:
        return True
    if "total transactions for this period" in lower:
        return True
    if lower.startswith("additional information on the next page"):
        return True
    if "page " in lower and " of " in lower:
        return True
    if "spark cash" in lower and "ending in" in lower:
        return True
    if "days in billing cycle" in lower:
        return True
    return False


def _is_continuation_line(line: str) -> bool:
    lower = (line or "").strip().lower()
    return any(lower.startswith(p) for p in SKIP_DESCRIPTION_PREFIXES)


def _section_type_from_header(line: str) -> str | None:
    m = RE_CARDHOLDER_SECTION.match((line or "").strip())
    if not m:
        return None
    kind = m.group(1).lower()
    if "payment" in kind or "credit" in kind:
        return "credits"
    return "purchases"


def _signed_amount(amount: float, is_credit: bool, section: str | None) -> float:
    if is_credit or section == "credits":
        return abs(amount)
    return -abs(amount)


def _append_interest_rows(lines: list[str], period: tuple[int, int, int, int] | None, text: str, rows: list[dict]) -> None:
    closing = _closing_date(period, text)
    for idx, line in enumerate(lines):
        if line.strip().lower() != "interest charged":
            continue
        j = idx + 1
        while j < len(lines):
            cur = lines[j].strip()
            if not cur:
                j += 1
                continue
            lower = cur.lower()
            if lower.startswith("total interest for this period"):
                break
            if lower.startswith("interest charge on"):
                parsed = _parse_amount_line(lines[j + 1] if j + 1 < len(lines) else "")
                if parsed:
                    amt, _ = parsed
                    if amt > 0 and not any(
                        r.get("description", "").lower() == lower and abs(r.get("amount", 0)) == amt
                        for r in rows
                    ):
                        rows.append({
                            "date": closing,
                            "description": cur,
                            "amount": -abs(amt),
                        })
            j += 1
        break


def _parse_capital_one_cc_text(text: str) -> list[dict]:
    """
    Parse Capital One credit card statement text.
    Uses post date when available, otherwise trans date.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    period = _extract_statement_period(text)

    start_idx = 0
    for idx, line in enumerate(lines):
        lower = line.lower()
        if "visit capitalone.com" in lower and "detailed transactions" in lower:
            start_idx = idx + 1
            break
        if ": payments, credits and adjustments" in lower:
            start_idx = idx
            break

    end_idx = len(lines)
    for idx, line in enumerate(lines):
        if "interest charge calculation" in line.lower():
            end_idx = idx
            break

    rows: list[dict] = []
    current_section: str | None = None
    i = start_idx

    while i < end_idx:
        line = lines[i]
        if not line:
            i += 1
            continue

        lower = line.lower()
        if lower == "fees":
            current_section = "fees"
            i += 1
            continue
        if lower == "interest charged":
            current_section = "interest"
            i += 1
            continue

        sec = _section_type_from_header(line)
        if sec:
            current_section = sec
            i += 1
            continue

        if _is_skip_line(line):
            i += 1
            continue

        if not RE_MON_DAY.match(line):
            i += 1
            continue

        trans_date = line
        j = i + 1
        while j < end_idx and not lines[j].strip():
            j += 1
        if j >= end_idx or not RE_MON_DAY.match(lines[j]):
            i += 1
            continue

        post_date = lines[j]
        k = j + 1
        desc_parts: list[str] = []
        amount_info: tuple[float, bool] | None = None

        while k < end_idx:
            cur = lines[k].strip()
            if not cur:
                k += 1
                continue
            if _is_skip_line(cur):
                break
            if _section_type_from_header(cur):
                break
            if cur.lower() in ("fees", "interest charged"):
                break
            if RE_MON_DAY.match(cur) and desc_parts:
                break

            parsed = _parse_amount_line(cur)
            if parsed:
                amount_info = parsed
                k += 1
                break

            if _is_continuation_line(cur):
                k += 1
                continue

            desc_parts.append(cur)
            k += 1

        if not amount_info or not desc_parts:
            i += 1
            continue

        amount, is_credit = amount_info
        if amount == 0:
            i = k
            continue

        description = " ".join(desc_parts).strip()
        date_str = _normalize_mon_day(post_date, period) or _normalize_mon_day(trans_date, period)
        if not date_str:
            i = k
            continue

        rows.append({
            "date": date_str,
            "description": description,
            "amount": _signed_amount(amount, is_credit, current_section),
        })
        i = k

    _append_interest_rows(lines, period, text, rows)
    return rows


def parse_capital_one_cc_text(text: str) -> list[dict]:
    """Parse Capital One credit card statement text."""
    return _parse_capital_one_cc_text(text)


def parse_capital_one_cc_pdf(pdf_path: str) -> list[dict]:
    """Parse Capital One credit card statement PDF."""
    text = extract_text_from_pdf(pdf_path)
    return parse_capital_one_cc_text(text)


def main():
    if len(sys.argv) < 2:
        print("Usage: python parse_capital_one_cc.py <pdf_path> [--csv output.csv]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    rows = parse_capital_one_cc_pdf(pdf_path)
    print(f"Parsed {len(rows)} transactions")
    for r in rows:
        print(r)
    if len(sys.argv) >= 4 and sys.argv[2] == "--csv":
        import csv
        with open(sys.argv[3], "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["date", "description", "amount"])
            w.writeheader()
            w.writerows(rows)
        print(f"Wrote {sys.argv[3]}")


if __name__ == "__main__":
    main()
