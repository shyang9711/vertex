import fitz  # PyMuPDF
import pandas as pd
import re
import os
from datetime import datetime
from tkinter import Tk, filedialog, simpledialog, messagebox

# Hide Tkinter root
root = Tk()
root.withdraw()

# Ask whether to calculate split payment now
include_split = messagebox.askyesno(
    "Split Payment",
    "Calculate split payment count now?"
)

mode = None
threshold_hours = None

# Ask for mode
if include_split:
    mode = simpledialog.askinteger(
        "Select Mode",
        "Enter 1 for Split Payment mode\nEnter 2 for Daily Total Threshold mode"
    )
    if mode not in (1, 2):
        messagebox.showerror("Invalid Choice", "You must enter 1 or 2.")
        exit()

    if mode == 2:
        threshold_hours = simpledialog.askfloat(
            "Threshold",
            "Enter the daily hours threshold for split payment (inclusive)"
        )
        if threshold_hours is None or threshold_hours <= 0:
            messagebox.showerror("Invalid Threshold", "Must be greater than 0.")
            exit()

pdf_path = filedialog.askopenfilename(
    title="Select PDF File",
    filetypes=[("PDF Files", "*.pdf")]
)
if not pdf_path:
    print("❌ No file selected. Exiting...")
    exit()

# Load PDF
doc = fitz.open(pdf_path)

employee_header_pattern = re.compile(r'\[(.*?)\] (.+)')
numeric_line_pattern   = re.compile(r'^\d+(?:\.\d{2})?$')

# Match optional weekday + DATE TIME  (e.g., "Mon 08/11/2025 10:07 AM")
date_time_pattern = re.compile(
    r'(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?\s*'
    r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+'
    r'(\d{1,2}:\d{2}\s?(?:AM|PM))',
    re.IGNORECASE
)

time_only_pattern = re.compile(
    r'(\d{1,2}:\d{2}\s?(?:AM|PM))',
    re.IGNORECASE
)

weekday_token      = re.compile(r'^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)$', re.IGNORECASE)
date_only_pattern  = re.compile(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$')
time_only_pattern  = re.compile(r'^\d{1,2}:\d{2}\s?(?:AM|PM)$', re.IGNORECASE)

def parse_time_entries(block_lines):
    from collections import defaultdict
    if not include_split:
        return 0

    # -------- 1) Build datetime stamps from vertical layout --------
    stamps = []
    i = 0

    def _mkdt(d, t):
        s = f"{d} {t.replace(' ', '')}"
        for fmt in ("%m/%d/%Y %I:%M%p", "%m-%d-%Y %I:%M%p",
                    "%m/%d/%y %I:%M%p",  "%m-%d-%y %I:%M%p"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None

    skip_tokens = {"IN","OUT","REG","OVER","DOUBLE","TOTAL","BREAK"}
    n = len(block_lines)

    while i < n:
        tok = block_lines[i].strip()
        if not tok or tok in skip_tokens:
            i += 1
            continue

        # Case A: Weekday, then date, then time  (e.g., Mon / 08/11/2025 / 10:07 AM)
        if weekday_token.match(tok):
            if i+2 < n and date_only_pattern.match(block_lines[i+1].strip()) \
                       and time_only_pattern.match(block_lines[i+2].strip()):
                d = block_lines[i+1].strip()
                t = block_lines[i+2].strip()
                dt = _mkdt(d, t)
                if dt:
                    stamps.append(dt)
                i += 3
                continue

        # Case B: Date, then time on next line  (no weekday label)
        if date_only_pattern.match(tok):
            if i+1 < n and time_only_pattern.match(block_lines[i+1].strip()):
                d = tok
                t = block_lines[i+1].strip()
                dt = _mkdt(d, t)
                if dt:
                    stamps.append(dt)
                i += 2
                continue

        i += 1

    # -------- 2) Pair IN/OUT within each calendar day --------
    by_day = defaultdict(list)
    for dt in sorted(stamps):
        by_day[dt.date()].append(dt)

    daily_shifts = defaultdict(list)
    for d, ts in by_day.items():
        if len(ts) % 2 == 1:  # drop dangling stamp if odd
            ts = ts[:-1]
        for k in range(0, len(ts), 2):
            start, end = ts[k], ts[k+1]
            hrs = (end - start).total_seconds() / 3600.0
            if hrs > 1e-6:
                daily_shifts[d].append(hrs)

    # -------- 3) Modes --------
    if mode == 1:
        # Split day = at least two shifts of >= 0.5h
        return sum(1 for hs in daily_shifts.values()
                   if sum(1 for h in hs if h >= 0.5) >= 2)

    # ------- Mode 2: your existing printed-total scan -------
    numeric = []
    for l in block_lines:
        l = l.strip()
        if numeric_line_pattern.fullmatch(l):
            try:
                numeric.append(float(l))
            except ValueError:
                pass
        else:
            numeric.append(None)

    day_totals, chunk = [], []
    for token in numeric + [None]:  # flush tail
        if token is None:
            if len(chunk) >= 4:
                for k in range(0, len(chunk) - 3):
                    a, b, c, d = chunk[k:k+4]
                    if a >= 0 and b >= 0 and c >= 0 and d >= max(a, b, c):
                        day_totals.append(d); break
            chunk = []
        else:
            chunk.append(token)

    if mode == 2:
        eps = 1e-9
        return sum(1 for total in day_totals if total + eps >= threshold_hours)

    return 0



# Extract lines from PDF
all_lines = []
for page in doc:
    all_lines.extend(page.get_text().splitlines())

data = []
employee_name = None
employee_block = []

def flush_employee_block():
    global employee_block, employee_name
    if employee_name and employee_block:
        numeric_values = [float(l) for l in employee_block if numeric_line_pattern.fullmatch(l)]
        split_count = parse_time_entries(employee_block) if include_split else None
        if len(numeric_values) >= 6:
            last6 = numeric_values[-6:]
            reg, over, double, total = last6[2], last6[3], last6[4], last6[5]
            row = {
                "Employee": employee_name,
                "REG": reg,
                "OVER": over,
                "DOUBLE": double,
                "TOTAL": total,
                "VALID SUM": abs((reg + over + double) - total) < 0.01,
            }
            if include_split:
                row["SPLIT COUNT"] = split_count
            data.append(row)
    employee_block = []

# Main loop
for line in all_lines:
    line = line.strip()
    m = employee_header_pattern.match(line)
    if m:
        if employee_name:
            flush_employee_block()
        employee_name = m.group(2).strip()
        employee_block = []
    elif employee_name:
        employee_block.append(line)

# Last employee
flush_employee_block()

df = pd.DataFrame(data)

output_folder = os.path.dirname(pdf_path)
csv_path = os.path.join(output_folder, "employee_hours_summary.csv")
df.to_csv(csv_path, index=False)

print("\n✅ Extraction complete!")
print(f"CSV saved to: {csv_path}\n")
print(df)
