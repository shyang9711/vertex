import pandas as pd
import fitz  # PyMuPDF
import warnings
import re
from io import StringIO
from collections import Counter
from datetime import datetime, date
from tkinter import Tk, Toplevel, Text, Scrollbar, Button, END, RIGHT, Y, LEFT, BOTH, messagebox, filedialog, StringVar, OptionMenu, Label
try:
    from colorama import init, Fore, Style
except Exception:
    class _Dummy:
        def __getattr__(self, _): return ""
    Fore = Style = _Dummy()

init(autoreset=True)  # Reset color after each print

# Auto-install required packages
import subprocess
import sys
import os

# Best-effort: try to make stdout UTF-8 on Windows terminals
try:
    if os.name == "nt":
        sys.stdout.reconfigure(encoding="cp1252", errors="replace")
        sys.stderr.reconfigure(encoding="cp1252", errors="replace")
except Exception:
    pass

BAD = "X"
OK = "OK"

required_packages = [
    "pandas",
    "PyMuPDF",    # for fitz
    "colorama"
]

for package in required_packages:
    try:
        __import__(package if package != "PyMuPDF" else "fitz")
    except ImportError:
        print(f"Installing missing package: {package}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])


# Define colors
GREEN = Fore.GREEN
RED = Fore.RED
RESET = Style.RESET_ALL

def get_previous_quarter_and_year():
    today = datetime.today()
    current_month = today.month
    current_year = today.year

    # Determine current quarter (1–4)
    current_quarter = (current_month - 1) // 3 + 1

    # Previous quarter logic
    if current_quarter == 1:
        return str(current_year - 1), "Q4"
    else:
        return str(current_year), f"Q{current_quarter - 1}"
    
def quarter_date_range(tax_year: str, tax_quarter: str):
    y = int(tax_year)
    q = tax_quarter.upper().strip()

    if q == "Q1":
        start = date(y, 1, 1);  end = date(y, 3, 31)
    elif q == "Q2":
        start = date(y, 4, 1);  end = date(y, 6, 30)
    elif q == "Q3":
        start = date(y, 7, 1);  end = date(y, 9, 30)
    elif q == "Q4":
        start = date(y, 10, 1); end = date(y, 12, 31)
    else:
        raise ValueError(f"Invalid quarter: {tax_quarter}")

    return pd.Timestamp(start), pd.Timestamp(end)

def is_date_like(val: str) -> bool:
    try:
        pd.to_datetime(val, errors="raise")
        return True
    except Exception:
        return False


# Prompt user for tax year and quarter
def get_tax_period_input():

    # Get default values
    default_year, default_quarter = get_previous_quarter_and_year()


    period_root = Toplevel()
    period_root.title("Tax Period Selection")
    period_root.geometry("300x180")

    year_var = StringVar(value=default_year)
    quarter_var = StringVar(value=default_quarter)

    Label(period_root, text="Enter Tax Year:").pack(pady=(10, 2))
    year_entry = Text(period_root, height=1, width=10)
    year_entry.insert("1.0", default_year)
    year_entry.pack()

    Label(period_root, text="Select Quarter:").pack(pady=(10, 2))
    quarter_dropdown = OptionMenu(period_root, quarter_var, "Q1", "Q2", "Q3", "Q4")
    quarter_dropdown.pack()

    def submit():
        year = year_entry.get("1.0", "end").strip()
        if not year.isdigit():
            messagebox.showerror("Invalid Input", "Year must be numeric.")
            return
        period_root.result = (year, quarter_var.get())
        period_root.destroy()

    def cancel():
        period_root.result = None
        period_root.destroy()

    period_root.protocol("WM_DELETE_WINDOW", cancel)


    btn_row = Toplevel(period_root)
    btn_row.overrideredirect(True)
    btn_row.destroy()

    Button(period_root, text="Submit", command=submit).pack(pady=(10, 2))
    Button(period_root, text="Cancel", command=cancel).pack()

    period_root.result = None
    period_root.grab_set()
    period_root.wait_window()
    return period_root.result

# --- Step 1: Prompt user for Excel-style text input
def get_excel_input():
    def on_submit():
        content = text.get("1.0", END).strip()
        if content:
            window.input = content
            window.destroy()
        else:
            messagebox.showwarning("Empty Input", "Please paste the Excel data.")
    
    window = Toplevel()
    window.title("Paste Excel Data")
    window.geometry("700x400")  # width x height

    scrollbar = Scrollbar(window)
    scrollbar.pack(side=RIGHT, fill=Y)

    text = Text(window, wrap="none", yscrollcommand=scrollbar.set)
    text.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.config(command=text.yview)

    submit_btn = Button(window, text="Submit", command=on_submit)
    submit_btn.pack()

    window.input = None
    window.grab_set()
    window.wait_window()
    return window.input

# Call this instead of simpledialog
root = Tk()
root.withdraw()

# Get user-defined tax year and quarter
res = get_tax_period_input()
if not res:
    print("Tax period selection canceled. Exiting.")
    root.destroy()
    sys.exit(0)
tax_year, tax_quarter = res

current_year = datetime.now().year
if int(tax_year) > current_year:
    print(f"{RED}{BAD} Tax year {tax_year} is in the future. Exiting...{RESET}")
    root.destroy()
    sys.exit(1)


excel_text = get_excel_input()

if not excel_text:
    print("No input provided. Exiting.")
    sys.exit(0)

# --- Step 2: Load EFTPS and EDD PDFs using a file dialog ---
print("Select the EFTPS PDF file...")
eftps_path = filedialog.askopenfilename(title="Select EFTPS PDF")
print("Select the EDD PDF file...")
edd_path = filedialog.askopenfilename(title="Select EDD PDF")

# --- Step 3: Parse Excel text ---
def parse_excel_dates(series: pd.Series) -> pd.Series:
    """
    Parse mixed Excel-like date strings without triggering Pandas 'Could not infer format' warnings.
    Tries several common formats first; falls back to dateutil (warning-suppressed) only for leftovers.
    """
    s = series.astype(str).str.strip()

    fmts = [
        "%Y-%m-%d",   # 2025-10-08
        "%m/%d/%Y",   # 10/08/2025
        "%m/%d/%y",   # 10/08/25
        "%m-%d-%Y",   # 10-08-2025
        "%m-%d-%y",   # 10-08-25
        "%m.%d.%Y",   # 10.08.2025
        "%m.%d.%y",   # 10.08.25
        "%d-%b-%Y",   # 08-Oct-2025
        "%d-%b-%y",   # 08-Oct-25
    ]

    out = pd.Series(pd.NaT, index=s.index)

    for fmt in fmts:
        m = out.isna()
        if not m.any():
            break
        parsed = pd.to_datetime(s[m], format=fmt, errors="coerce")
        out.loc[m] = parsed

    # Fallback for any remaining weird cases (suppress the warning you’re seeing)
    m = out.isna()
    if m.any():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            out.loc[m] = pd.to_datetime(s[m], errors="coerce")

    return out

rows = [re.split(r"\s+", line.strip()) for line in excel_text.strip().split("\n") if line.strip()]
col_count = len(rows[0])

first_row = rows[0]
date_cols = [i for i, v in enumerate(first_row) if is_date_like(v)]

if col_count == 10:
    # Two date columns: take the later one
    excel_df = pd.DataFrame(rows, columns=["Date1", "Date2", "Total", "UI", "ETT", "UI+ETT", "SDI", "PIT", "P+I", "EDD_Total"])
    excel_df["Date1"] = parse_excel_dates(excel_df["Date1"])
    excel_df["Date2"] = parse_excel_dates(excel_df["Date2"])
    excel_df["Date"] = excel_df[["Date1", "Date2"]].max(axis=1)
    excel_df.drop(columns=["Date1", "Date2"], inplace=True)

elif col_count == 9:
    excel_df = pd.DataFrame(rows, columns=["Date", "Total", "UI", "ETT", "UI+ETT", "SDI", "PIT", "P+I", "EDD_Total"])
    excel_df["Date"] = parse_excel_dates(excel_df["Date"])
    
elif col_count == 8:
    if len(date_cols) == 2:
        # Case 1: two dates → NO UI+ETT
        excel_df = pd.DataFrame(
            rows,
            columns=["Date1", "Date2", "Total", "UI", "ETT", "SDI", "PIT", "EDD_Total"]
        )
        excel_df["Date1"] = parse_excel_dates(excel_df["Date1"])
        excel_df["Date2"] = parse_excel_dates(excel_df["Date2"])
        excel_df["Date"] = excel_df[["Date1", "Date2"]].max(axis=1)
        excel_df.drop(columns=["Date1", "Date2"], inplace=True)

    elif len(date_cols) == 1:
        # Case 2: one date → HAS UI+ETT
        excel_df = pd.DataFrame(
            rows,
            columns=["Date", "Total", "UI", "ETT", "UI+ETT", "SDI", "PIT", "EDD_Total"]
        )
        excel_df["Date"] = parse_excel_dates(excel_df["Date"])

    else:
        raise ValueError("8-column input but could not determine date columns")
    
elif col_count == 7:
    excel_df = pd.DataFrame(rows, columns=["Date", "Total", "UI", "ETT", "SDI", "PIT", "EDD_Total"])
    excel_df["Date"] = parse_excel_dates(excel_df["Date"])
else:
    raise ValueError(f"Unexpected column count: {col_count}")

# Clean numeric columns
for col in excel_df.columns:
    if col == "Date":
        continue
    excel_df[col] = (
        excel_df[col]
        .astype(str)
        .str.replace("-", "0") # Replace "-" with "0"
        .str.replace(",", "", regex=False)  # Remove commas
        .astype(float)
    )

if excel_df["Date"].isnull().any():
    print(f"{BAD} Some dates couldn't be parsed. Please check your input.")
    print(excel_df[excel_df["Date"].isnull()])
    sys.exit(1)

# --- Step 4: Parse EFTPS PDF ---
def parse_eftps(pdf_path, tax_year, tax_quarter):
    doc = fitz.open(pdf_path)
    text = "\n".join(page.get_text() for page in doc)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    records = []
    for i in range(len(lines) - 5):
        try:
            settlement_date = lines[i]
            initiation_date = lines[i + 1]
            tax_form = lines[i + 2]
            tax_period = lines[i + 3]
            amount = lines[i + 4]
            status = lines[i + 5].strip()
            if status == "Return":
                status = "Returned"

            if (
                re.match(r"\d{4}-\d{2}-\d{2}", settlement_date)
                and tax_period == f"{tax_year}/{tax_quarter}"
                and status in ("Settled", "Scheduled", "Returned")
            ):
                records.append({
                    "SettlementDate": pd.to_datetime(settlement_date),
                    "InitiationDate": pd.to_datetime(initiation_date, errors="coerce"),
                    "Amount": float(amount.replace(",", "")),
                    "Status": status
                })
        except Exception:
            continue

    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values(["SettlementDate", "InitiationDate"], ascending=[True, True]).reset_index(drop=True)
    return df

def reconcile_eftps_returns(eftps_df: pd.DataFrame):
    """
    For each Returned transaction:
      - find ONE prior successful payment (Settled/Scheduled) with same amount
        and SettlementDate <= return date
      - remove it once (most recent prior one)
      - add info to returned bucket

    Also checks whether each return amount was repaid later by a successful payment.
    """
    if eftps_df is None or eftps_df.empty:
        return eftps_df, pd.DataFrame(), []

    df = eftps_df.copy().sort_values(["SettlementDate", "InitiationDate"], ascending=[True, True]).reset_index(drop=True)

    # Separate
    success_mask = df["Status"].isin(["Settled", "Scheduled"])
    return_mask  = df["Status"].eq("Returned")

    success = df[success_mask].copy().reset_index(drop=False)  # keep original row index
    returns = df[return_mask].copy().reset_index(drop=False)

    # Track which successful rows get canceled by returns
    canceled_success_row_ids = set()
    returned_bucket = []

    # For matching: we want most recent successful BEFORE/EQUAL return date, same amount
    # We'll scan returns in chronological order (already sorted)
    for _, r in returns.iterrows():
        r_date = r["SettlementDate"]
        r_amt  = round(float(r["Amount"]), 2)

        # Candidate successful payments not already canceled
        candidates = success[
            (~success["index"].isin(canceled_success_row_ids)) &
            (success["SettlementDate"] <= r_date) &
            (success["Amount"].round(2) == r_amt)
        ]

        if not candidates.empty:
            # pick most recent prior (max SettlementDate, then max InitiationDate)
            pick = candidates.sort_values(
                ["SettlementDate", "InitiationDate", "index"],
                ascending=[False, False, False]
            ).iloc[0]

            canceled_success_row_ids.add(int(pick["index"]))
            returned_bucket.append({
                "ReturnedDate": r_date,
                "Amount": r_amt,

                # This is the key your matching logic needs
                "CanceledPaymentDate": pick["SettlementDate"],

                # show "Returned" (not the original status)
                "CanceledPaymentStatus": "Returned",
            })
        else:
            # Return exists but we couldn't find a payment to cancel (still important)
            returned_bucket.append({
                "ReturnedDate": r_date,
                "Amount": r_amt,
                "CanceledPaymentDate": pd.NaT,
                "CanceledPaymentStatus": "NOT_FOUND",
                "CanceledPaymentOriginalStatus": "",
            })

    returned_df = pd.DataFrame(returned_bucket)

    # Effective success = all success minus canceled ones
    effective_success = success[~success["index"].isin(canceled_success_row_ids)].copy()
    effective_success = effective_success.drop(columns=["index"]).reset_index(drop=True)

    # Now: Return health check
    # A return is "repaid" if there exists a later successful payment (effective) of same amount AFTER return date
    return_flags = []
    if not returned_df.empty:
        for _, rr in returned_df.iterrows():
            r_date = rr["ReturnedDate"]
            r_amt = round(float(rr["Amount"]), 2)

            repaid = not effective_success[
                (effective_success["SettlementDate"] > r_date) &
                (effective_success["Amount"].round(2) == r_amt)
            ].empty

            if not repaid:
                return_flags.append(
                    f"{BAD} Returned payment {r_amt} on {pd.to_datetime(r_date).date()} was not repaid later."
                )

    return effective_success, returned_df, return_flags


# --- Step 5: Parse EDD PDF ---
def parse_edd(pdf_path, tax_year, tax_quarter):
    """
    Parse EDD e-Services 'My Payments' export and return DE88 payment amounts
    whose PAY DATE is within the selected quarter (e.g., Q4 = Oct 1–Dec 31).
    """
    doc = fitz.open(pdf_path)
    text = "\n".join(page.get_text() for page in doc)

    q_start, q_end = quarter_date_range(tax_year, tax_quarter)

    # Typical row in extracted text looks like:
    # 31-Dec-2025 DE88 Payment
    # 12-Dec-2025 $46.76 $46.76 $2.74 $0.18 $42.97 $0.87 $0.00
    #
    # We capture:
    #  - Period (date)
    #  - Pay Date (date)
    #  - Orig Payment ($)
    #  - Applied Amount ($)  (if present)
    row_pat = re.compile(
        r'(?P<period>\d{2}-[A-Za-z]{3}-\d{4})\s+DE88\s+Payment\s+'
        r'(?P<paydate>\d{2}-[A-Za-z]{3}-\d{4})\s+'
        r'\$(?P<orig>[\d,]+\.\d{2})'
        r'(?:\s+\$(?P<applied>[\d,]+\.\d{2}))?',
        re.MULTILINE
    )

    payments = []
    for m in row_pat.finditer(text):
        pay_dt = pd.to_datetime(m.group("paydate"), format="%d-%b-%Y", errors="coerce")
        if pd.isna(pay_dt):
            continue

        # Quarter filter (inclusive)
        if not (q_start <= pay_dt <= q_end):
            continue

        # Prefer Applied Amount; fallback to Orig
        amt_str = m.group("applied") or m.group("orig")
        amt = float(amt_str.replace(",", ""))
        payments.append(amt)

    return payments


eftps_raw_df = parse_eftps(eftps_path, tax_year, tax_quarter)
if eftps_raw_df.empty:
    print(f"\n{RED}{BAD} No EFTPS records found for {tax_year}/{tax_quarter}.{RESET}")
    sys.exit(1)

eftps_df, eftps_returned_df, eftps_return_flags = reconcile_eftps_returns(eftps_raw_df)

if eftps_df.empty:
    print(f"\n{RED}{BAD} No EFTPS records found for {tax_year}/{tax_quarter}.{RESET}")
    sys.exit(1)

edd_payments = parse_edd(edd_path, tax_year, tax_quarter)
if not edd_payments:
    print(f"\n{RED}{BAD} No EDD records found for {tax_year} {tax_quarter}.{RESET}")
    sys.exit(1)

# --- Step 6: Validation ---
eftps_flags = []

# Identify Excel rows that correspond to a payment that was later RETURNED
# Match rule: Excel Date == CanceledPaymentDate AND Excel Total == Returned Amount
excel_is_returned = pd.Series(False, index=excel_df.index)

if eftps_returned_df is not None and not eftps_returned_df.empty:
    r = eftps_returned_df.copy()

    # --- Robust column detection (handles small naming differences) ---
    possible_canceled_cols = [
        "CanceledPaymentDate",
        "Canceled Payment Date",
        "Canceled_Payment_Date",
        "CanceledPayment Date",
        "CanceledPaymentDate ",
        "CanceledPaymentDt",
        "CanceledDate",
    ]
    possible_amount_cols = [
        "Amount",
        "ReturnedAmount",
        "ReturnAmount",
        "Returned Amount",
        "Amount ",
    ]

    canceled_col = next((c for c in possible_canceled_cols if c in r.columns), None)
    amount_col   = next((c for c in possible_amount_cols if c in r.columns), None)

    # If we cannot find the columns, treat as "no returns" (don’t crash the program)
    if canceled_col is None or amount_col is None:
        print(f"[WARN] Returned bucket columns found: {list(r.columns)}")
        print("[WARN] Could not detect canceled-date/amount columns in returned bucket; skipping returned matching.")
        r = pd.DataFrame()  # force "no returns" behavior
    else:
        # Normalize to canonical names used by the rest of the script
        if canceled_col != "CanceledPaymentDate":
            r["CanceledPaymentDate"] = r[canceled_col]
        r["CanceledPaymentDate"] = pd.to_datetime(r["CanceledPaymentDate"], errors="coerce")

        if amount_col != "Amount":
            r["Amount"] = r[amount_col]
        r["Amount"] = pd.to_numeric(r["Amount"], errors="coerce").astype(float).round(2)
else:
    r = pd.DataFrame()

# --- Mark Excel rows that correspond to returned/canceled EFTPS payments ---
if r is not None and not r.empty and "CanceledPaymentDate" in r.columns and "Amount" in r.columns:
    for i, row in excel_df.iterrows():
        d = row["Date"]
        amt = round(float(row["Total"]), 2)

        hit = r[
            (r["CanceledPaymentDate"] == d) &
            (r["Amount"].round(2) == amt)
        ]
        if not hit.empty:
            excel_is_returned.loc[i] = True

# --- EFTPS per-date multiset validation (supports multiple payments same date) ---

# Build Excel paid amounts grouped by date (exclude returned/canceled rows)
excel_paid = excel_df.loc[~excel_is_returned, ["Date", "Total"]].copy()
excel_paid["Total"] = excel_paid["Total"].astype(float).round(2)

# Build EFTPS paid amounts grouped by date
eftps_paid = eftps_df[["SettlementDate", "Amount"]].copy()
eftps_paid["Amount"] = eftps_paid["Amount"].astype(float).round(2)

# Compare Counters per date
all_dates = sorted(set(excel_paid["Date"].unique()) | set(eftps_paid["SettlementDate"].unique()))

for d in all_dates:
    ex_list = excel_paid.loc[excel_paid["Date"] == d, "Total"].tolist()
    ef_list = eftps_paid.loc[eftps_paid["SettlementDate"] == d, "Amount"].tolist()

    ex_c = Counter(ex_list)
    ef_c = Counter(ef_list)

    if ex_c == ef_c:
        continue

    # Nice diagnostics: what is missing / extra
    missing = []
    extra = []

    # Missing in EFTPS means Excel has more of that amount than EFTPS
    for amt, cnt in (ex_c - ef_c).items():
        missing.append((amt, cnt))

    # Extra in EFTPS means EFTPS has more of that amount than Excel
    for amt, cnt in (ef_c - ex_c).items():
        extra.append((amt, cnt))

    msg_parts = [f"{BAD} EFTPS paid mismatch on {pd.to_datetime(d).date()}"]
    if missing:
        msg_parts.append(f"Missing in EFTPS: {missing}")
    if extra:
        msg_parts.append(f"Extra in EFTPS: {extra}")

    # Also show full lists for quick eyeballing
    msg_parts.append(f"ExcelPaid={sorted(ex_list)}")
    msg_parts.append(f"EFTPSPaid={sorted(ef_list)}")

    eftps_flags.append(" — ".join(msg_parts))


# --- EFTPS Sum Checks (split into Paid vs Returned) ---

# Paid = Excel excluding rows that match canceled-by-return payments
sum_excel_paid = round(float(excel_df.loc[~excel_is_returned, "Total"].sum()), 2)

# Returned = Excel rows that match canceled-by-return payments
sum_excel_returned = round(float(excel_df.loc[excel_is_returned, "Total"].sum()), 2)

# EFTPS effective paid (already excludes canceled payments)
sum_eftps_paid = round(float(eftps_df["Amount"].sum()), 2)

# EFTPS returned total (from returned bucket)
sum_eftps_returned = 0.00
if eftps_returned_df is not None and not eftps_returned_df.empty:
    sum_eftps_returned = round(float(eftps_returned_df["Amount"].astype(float).sum()), 2)

# Compare Paid totals
if sum_excel_paid != sum_eftps_paid:
    eftps_flags.append(
        f"{BAD} EFTPS PAID sum mismatch: ExcelPaid({sum_excel_paid}) vs EFTPSPaid({sum_eftps_paid})"
    )

# Compare Returned totals (only if either side has returns)
if (sum_excel_returned != 0.00) or (sum_eftps_returned != 0.00):
    if sum_excel_returned != sum_eftps_returned:
        eftps_flags.append(
            f"{BAD} EFTPS RETURNED sum mismatch: ExcelReturned({sum_excel_returned}) vs EFTPSReturned({sum_eftps_returned})"
        )
    else:
        # Optional informational line (no error), if you want it visible in console:
        # print(f"{GREEN}{OK} Returned totals match: {sum_excel_returned}{RESET}")
        pass


edd_flags = []

# Compare sums ONLY (line mismatches are allowed if sums match)
excel_edd_totals = excel_df["EDD_Total"].astype(float).round(2).tolist()
pdf_edd_totals   = [round(float(x), 2) for x in edd_payments]

sum_edd_excel = round(sum(excel_edd_totals), 2)
sum_edd_pdf   = round(sum(pdf_edd_totals), 2)

# Tolerance if needed change
SUM_TOL = 0.00

if abs(sum_edd_excel - sum_edd_pdf) <= SUM_TOL:
    # PASS: sums match, do not flag line-item mismatches
    pass
else:
    # FAIL: sums don't match -> now show helpful diagnostics
    excel_edd_counter = Counter(excel_edd_totals)
    edd_pdf_counter   = Counter(pdf_edd_totals)

    # Show which values have different counts
    all_vals = sorted(set(excel_edd_counter.keys()) | set(edd_pdf_counter.keys()))
    for val in all_vals:
        excel_count = excel_edd_counter.get(val, 0)
        pdf_count   = edd_pdf_counter.get(val, 0)
        if excel_count != pdf_count:
            edd_flags.append(
                f"{BAD} EDD Total {val} count mismatch — Excel: {excel_count}, EDD PDF: {pdf_count}"
            )

    edd_flags.append(f"{BAD} EDD sum mismatch — Excel: {sum_edd_excel}, EDD PDF: {sum_edd_pdf}")



# Optional: Uncomment to print tables
print("\nEFTPS Data:")
print(eftps_df)
print("\nEFTPS Returned Bucket:")
print(eftps_returned_df if not eftps_returned_df.empty else "[]")
print("\nEDD Data:")
print(edd_payments)

# --- Step 7: Print Results ---
print("\n--- EFTPS Validation ---")
eftps_flags.extend(eftps_return_flags)
if eftps_flags:
    print(RED + f"{BAD} " + ("\n" + f"{BAD} ").join(eftps_flags) + RESET)
else:
    print(GREEN + f"{OK} All EFTPS records match." + RESET)

print("\n--- EDD Validation ---")
if edd_flags:
    print(RED + f"{BAD} " + ("\n" + f"{BAD} ").join(edd_flags) + RESET)
else:
    print(GREEN + f"{OK} All EDD payments match." + RESET)
