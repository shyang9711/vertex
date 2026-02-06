"""
Parse bank statement text where Debits and Credits are in separate columns
but both amounts are positive. We distinguish by description keywords:
- ELECTRONIC CREDIT, CREDIT, DEPOSIT -> credit (positive amount)
- DEBIT CARD PURCH, DEBIT, PAYMENT -> debit (negative amount)

Text format per transaction line:
  MM/DD  Description  Amount  Balance
e.g.:
  12/01 ELECTRONIC CREDIT 34.66 3,361.81
  Square Inc SQ251201 T37ZB12J9KS3YYP Dec 01   (continuation line)

Or for debit:
  12/01 DEBIT CARD PURCH Card Ending in 3965 ... 30.00 4,554.48
"""
import re

# One line: date, description (middle), amount, balance (last two numbers)
RE_TX_LINE = re.compile(
    r"^(\d{1,2}/\d{1,2})\s+(.+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)

# Keywords that indicate a credit (money in) — amount stays positive
CREDIT_PHRASES = (
    "ELECTRONIC CREDIT",
    "CREDIT",
    "DEPOSIT",
    "MISC CREDIT",
    "ATM CREDIT",
    "TRANSFER CREDIT",
)

# Keywords that indicate a debit (money out) — amount will be stored negative
DEBIT_PHRASES = (
    "DEBIT CARD PURCH",
    "DEBIT CARD",
    "DEBIT ",
    "PAYMENT",
    "WITHDRAWAL",
    "PURCHASE",
)


def _is_credit(description: str) -> bool:
    """True if transaction is a credit (money in)."""
    u = description.upper()
    for phrase in CREDIT_PHRASES:
        if phrase.upper() in u:
            return True
    return False


def _is_debit(description: str) -> bool:
    """True if transaction is a debit (money out)."""
    u = description.upper()
    for phrase in DEBIT_PHRASES:
        if phrase.upper() in u:
            return True
    return False


def _extract_statement_year(text: str) -> int | None:
    """Get statement year from header if present."""
    lines = text.split("\n")
    head = " ".join(lines[:50])
    m = re.search(r"\d{1,2}/\d{1,2}/(\d{4})", head)
    if m:
        y = int(m.group(1))
        if 1990 <= y <= 2030:
            return y
    for line in lines[:30]:
        for m in re.finditer(r"(\d{4})", line):
            y = int(m.group(1))
            if 1990 <= y <= 2030:
                return y
    return None


def _normalize_date_to_dd_mm_yyyy(date_str: str, statement_year: int | None) -> str:
    """Normalize MM/DD or MM/DD/YY to dd/mm/yyyy."""
    if not date_str or not date_str.strip():
        return ""
    date_str = date_str.strip()
    parts = date_str.split("/")
    if len(parts) == 3:
        try:
            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
            year = 2000 + y if y < 100 and y < 50 else (1900 + y if y < 100 else y)
            return f"{d:02d}/{m:02d}/{year}"
        except (ValueError, IndexError):
            return date_str
    if len(parts) == 2:
        try:
            m, d = int(parts[0]), int(parts[1])
            year = statement_year if statement_year and 1990 <= statement_year <= 2030 else 2000
            return f"{d:02d}/{m:02d}/{year}"
        except (ValueError, IndexError):
            return date_str
    return date_str


def parse_debit_credit_columns_text(text: str) -> list[dict]:
    """
    Parse statement text where each transaction has:
      MM/DD  Description  Amount  Balance
    with optional continuation lines (no date/amount) appended to description.
    Both debits and credits are positive in the source; we set amount sign by description:
    - Credit phrases (ELECTRONIC CREDIT, etc.) -> amount positive
    - Debit phrases (DEBIT CARD PURCH, etc.) -> amount negative
    Returns list of {"date": "dd/mm/yyyy", "description": "...", "amount": float}.
    """
    lines = [ln.strip() for ln in text.split("\n")]
    rows = []
    current = None
    statement_year = _extract_statement_year(text)

    for line in lines:
        if not line:
            continue
        m = RE_TX_LINE.match(line)
        if m:
            if current is not None:
                # Save previous transaction
                amt = float(current["amount_str"].replace(",", ""))
                if _is_credit(current["description"]):
                    current["amount"] = amt
                elif _is_debit(current["description"]):
                    current["amount"] = -amt
                else:
                    # Default: treat as debit if unknown
                    current["amount"] = -amt
                current["date"] = _normalize_date_to_dd_mm_yyyy(current["date_raw"], statement_year)
                rows.append({
                    "date": current["date"],
                    "description": current["description"].strip(),
                    "amount": current["amount"],
                })
            date_raw, desc, amount_str, balance = m.groups()
            current = {
                "date_raw": date_raw,
                "description": desc,
                "amount_str": amount_str,
                "amount": 0.0,
            }
            continue
        if current is not None:
            # Continuation line (e.g. "Square Inc SQ251201 ...")
            current["description"] += " " + line
    if current is not None:
        amt = float(current["amount_str"].replace(",", ""))
        if _is_credit(current["description"]):
            current["amount"] = amt
        elif _is_debit(current["description"]):
            current["amount"] = -amt
        else:
            current["amount"] = -amt
        current["date"] = _normalize_date_to_dd_mm_yyyy(current["date_raw"], statement_year)
        rows.append({
            "date": current["date"],
            "description": current["description"].strip(),
            "amount": current["amount"],
        })
    return rows
