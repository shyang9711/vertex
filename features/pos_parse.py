import fitz
import pandas as pd
import re
import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import Tk, filedialog

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def ask_split_payment_options(parent):
    """One dialog for split-payment settings. Returns (include_split, mode, threshold) or None if cancelled."""
    dlg = tk.Toplevel(parent)
    dlg.title("Split Payment")
    dlg.resizable(False, False)
    dlg.transient(parent)
    dlg.grab_set()

    result = {"ok": False}

    BG = "#f4f6f9"
    CARD = "#ffffff"
    FG = "#1a202c"
    MUTED = "#718096"
    ACCENT = "#3182ce"
    ACCENT_HOVER = "#2b6cb0"
    BORDER = "#e2e8f0"

    dlg.configure(bg=BG)

    outer = tk.Frame(dlg, bg=BG, padx=20, pady=18)
    outer.pack(fill="both", expand=True)

    tk.Label(
        outer, text="Split Payment", font=("Segoe UI", 14, "bold"),
        fg=FG, bg=BG, anchor="w",
    ).pack(fill="x")
    tk.Label(
        outer,
        text="Optionally count split-payment days while extracting payroll hours.",
        font=("Segoe UI", 9), fg=MUTED, bg=BG, wraplength=360, justify="left",
        anchor="w",
    ).pack(fill="x", pady=(2, 14))

    card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1, padx=14, pady=12)
    card.pack(fill="x")

    choice = tk.StringVar(value="skip")
    threshold_var = tk.StringVar(value="8")

    thresh_row = tk.Frame(card, bg=CARD)
    err_lbl = tk.Label(outer, text="", font=("Segoe UI", 8), fg="#c53030", bg=BG, anchor="w")

    thresh_entry = tk.Entry(
        thresh_row, textvariable=threshold_var, width=8,
        font=("Segoe UI", 10), relief="solid", bd=1,
    )

    def _sync(*_):
        enabled = choice.get() == "2"
        thresh_entry.configure(state="normal" if enabled else "disabled")
        err_lbl.configure(text="")

    def _radio(text, value, desc):
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=4)
        tk.Radiobutton(
            row, text=text, variable=choice, value=value,
            font=("Segoe UI", 10), fg=FG, bg=CARD, activebackground=CARD,
            selectcolor=CARD, anchor="w", command=_sync,
        ).pack(anchor="w")
        tk.Label(
            row, text=desc, font=("Segoe UI", 8), fg=MUTED, bg=CARD,
            wraplength=320, justify="left", anchor="w",
        ).pack(anchor="w", padx=(22, 0))

    _radio(
        "Skip split payment",
        "skip",
        "Extract hours only — no SPLIT COUNT column.",
    )
    _radio(
        "Split Payment mode",
        "1",
        "Count days with at least two shifts of 0.5 hours or more.",
    )
    _radio(
        "Daily Total Threshold mode",
        "2",
        "Count days whose printed daily total meets or exceeds a threshold.",
    )

    thresh_row.pack(fill="x", pady=(8, 0), padx=(22, 0))
    tk.Label(
        thresh_row, text="Daily hours threshold", font=("Segoe UI", 9),
        fg=FG, bg=CARD,
    ).pack(side="left")
    thresh_entry.pack(side="left", padx=(10, 0))
    err_lbl.pack(fill="x", pady=(8, 0))

    def _ok():
        c = choice.get()
        if c == "skip":
            result.update(ok=True, include_split=False, mode=None, threshold=None)
            dlg.destroy()
            return
        if c == "1":
            result.update(ok=True, include_split=True, mode=1, threshold=None)
            dlg.destroy()
            return
        try:
            thr = float(threshold_var.get().strip())
        except ValueError:
            err_lbl.configure(text="Enter a valid number for the threshold.")
            return
        if thr <= 0:
            err_lbl.configure(text="Threshold must be greater than 0.")
            return
        result.update(ok=True, include_split=True, mode=2, threshold=thr)
        dlg.destroy()

    def _cancel():
        result["ok"] = False
        dlg.destroy()

    btns = tk.Frame(outer, bg=BG)
    btns.pack(fill="x", pady=(16, 0))
    tk.Button(
        btns, text="Cancel", command=_cancel, font=("Segoe UI", 9),
        bg=CARD, fg=FG, relief="flat", padx=14, pady=6, cursor="hand2",
        highlightthickness=1, highlightbackground=BORDER,
    ).pack(side="right")
    tk.Button(
        btns, text="Continue", command=_ok, font=("Segoe UI", 9, "bold"),
        bg=ACCENT, fg="white", activebackground=ACCENT_HOVER, activeforeground="white",
        relief="flat", padx=16, pady=6, cursor="hand2",
    ).pack(side="right", padx=(0, 8))

    _sync()
    dlg.protocol("WM_DELETE_WINDOW", _cancel)
    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    x = (dlg.winfo_screenwidth() - w) // 2
    y = (dlg.winfo_screenheight() - h) // 3
    dlg.geometry(f"+{x}+{y}")
    dlg.wait_window()

    if not result.get("ok"):
        return None
    return result["include_split"], result["mode"], result["threshold"]


# Hide Tkinter root
root = Tk()
root.withdraw()

split_opts = ask_split_payment_options(root)
if split_opts is None:
    print("[X] Cancelled. Exiting...")
    sys.exit(0)

include_split, mode, threshold_hours = split_opts

pdf_path = filedialog.askopenfilename(
    title="Select PDF File",
    filetypes=[("PDF Files", "*.pdf")]
)
if not pdf_path:
    print("[X] No file selected. Exiting...")
    sys.exit(0)

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
