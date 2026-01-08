import pandas as pd
import fitz  # PyMuPDF
import re
from io import StringIO
from collections import Counter
from datetime import datetime
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

def _can_print(s: str) -> bool:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        s.encode(enc, errors="strict")
        return True
    except Exception:
        return False

# Best-effort: try to make stdout UTF-8 on Windows terminals
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Symbols (emoji when possible, ASCII fallback when not)
BAD = "❌" if _can_print("❌") else "X"
OK = "✔" if _can_print("✔") else "OK"

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
    year_entry.insert("1.0", "2025")
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
rows = [re.split(r"\s+", line.strip()) for line in excel_text.strip().split("\n") if line.strip()]
col_count = len(rows[0])

if col_count == 10:
    # Two date columns: take the later one
    excel_df = pd.DataFrame(rows, columns=["Date1", "Date2", "Total", "UI", "ETT", "UI+ETT", "SDI", "PIT", "P+I", "EDD_Total"])
    excel_df["Date1"] = pd.to_datetime(excel_df["Date1"], errors='coerce')
    excel_df["Date2"] = pd.to_datetime(excel_df["Date2"], errors='coerce')
    excel_df["Date"] = excel_df[["Date1", "Date2"]].max(axis=1)
    excel_df.drop(columns=["Date1", "Date2"], inplace=True)
elif col_count == 9:
    excel_df = pd.DataFrame(rows, columns=["Date", "Total", "UI", "ETT", "UI+ETT", "SDI", "PIT", "P+I", "EDD_Total"])
    excel_df["Date"] = pd.to_datetime(excel_df["Date"], errors='coerce')
elif col_count == 7:
    excel_df = pd.DataFrame(rows, columns=["Date", "Total", "UI", "ETT", "SDI", "PIT", "EDD_Total"])
    excel_df["Date"] = pd.to_datetime(excel_df["Date"], errors='coerce')
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

# Auto-detect date formats like 4/9/25 or 04.09.2025
excel_df["Date"] = pd.to_datetime(excel_df["Date"], errors='coerce')

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
            status = lines[i + 5]

            if (
                re.match(r"\d{4}-\d{2}-\d{2}", settlement_date)
                and tax_period == f"{tax_year}/{tax_quarter}"
                and (status == "Settled" or status == "Scheduled")
            ):
                records.append({
                    "SettlementDate": pd.to_datetime(settlement_date),
                    "Amount": float(amount.replace(",", ""))
                })
        except Exception:
            continue

    return pd.DataFrame(records)

# --- Step 5: Parse EDD PDF ---
def parse_edd(pdf_path):
    doc = fitz.open(pdf_path)
    text = "\n".join(page.get_text() for page in doc)
    quarter_end_month = {
        "Q1": "Mar",
        "Q2": "Jun",
        "Q3": "Sep",
        "Q4": "Dec"
    }[tax_quarter]

    edd_pattern = rf'\d{{2}}-{quarter_end_month}-{tax_year}\s+DE88\s+Payment\s+\d{{2}}-[A-Za-z]{{3}}-{tax_year}\s+\$([\d,]+\.\d{{2}})'
    matches = re.findall(edd_pattern, text)
    return [float(p.replace(",", "")) for p in matches]

eftps_df = parse_eftps(eftps_path, tax_year, tax_quarter)
if eftps_df.empty:
    print(f"\n{RED}{BAD} No EFTPS records found for {tax_year}/{tax_quarter}.{RESET}")
    sys.exit(1)

edd_payments = parse_edd(edd_path)
if not edd_payments:
    print(f"\n{RED}{BAD} No EDD records found for {tax_year} {tax_quarter}.{RESET}")
    sys.exit(1)


# --- Step 6: Validation ---
eftps_flags = []
for _, row in excel_df.iterrows():
    date = row["Date"]
    total = round(row["Total"], 2)
    matches = eftps_df[eftps_df["SettlementDate"] == date]
    if matches.empty:
        eftps_flags.append(f"{BAD} No EFTPS record for {date.date()}")
    elif total not in matches["Amount"].round(2).values:
        eftps_flags.append(f"{BAD} Amount mismatch on {date.date()} — Excel: {total}, EFTPS: {[float(a) for a in matches['Amount']]}")

if len(eftps_df) != len(excel_df):
    eftps_flags.append(f"{BAD} Row count mismatch: Excel({len(excel_df)}) vs EFTPS({len(eftps_df)})")


excel_edd_counter = Counter(excel_df["EDD_Total"].round(2))
edd_pdf_counter = Counter([round(x, 2) for x in edd_payments])

edd_flags = []
excel_edd_totals = list(excel_df["EDD_Total"].round(2))

sum_edd_excel = round(sum(excel_edd_totals), 2)
sum_edd_pdf   = round(sum(round(x, 2) for x in edd_payments), 2)

if sum_edd_excel != sum_edd_pdf:
    # Excel amounts missing or wrong counts
    for val, excel_count in excel_edd_counter.items():
        edd_count = edd_pdf_counter.get(val, 0)
        if excel_count != edd_count:
            edd_flags.append(
                f"{BAD} EDD Total {val} count mismatch — Excel: {excel_count}, EDD PDF: {edd_count}"
            )

    # PDF contains extra amounts not in Excel
    for val, edd_count in edd_pdf_counter.items():
        excel_count = excel_edd_counter.get(val, 0)
        if edd_count != excel_count:
            edd_flags.append(
                f"{BAD} EDD Total {val} count mismatch — Excel: {excel_count}, EDD PDF: {edd_count}"
            )

    # Optional: helpful totals line
    edd_flags.append(f"{BAD} EDD sum mismatch — Excel: {sum_edd_excel}, EDD PDF: {sum_edd_pdf}")


# Optional: Uncomment to print tables
print("\nExcel Data:")
print(excel_df)
print("\nEFTPS Data:")
print(eftps_df)
print("\nEDD Data:")
print(edd_payments)

# --- Step 7: Print Results ---
print("\n--- EFTPS Validation ---")
if eftps_flags:
    print(RED + f"{BAD} " + ("\n" + f"{BAD} ").join(eftps_flags) + RESET)
else:
    print(GREEN + f"{OK} All EFTPS records match." + RESET)

print("\n--- EDD Validation ---")
if edd_flags:
    print(RED + f"{BAD} " + ("\n" + f"{BAD} ").join(edd_flags) + RESET)
else:
    print(GREEN + f"{OK} All EDD payments match." + RESET)
