# -*- coding: utf-8 -*-
"""
FS x Credit-Card matcher (Meals classification)
- Auto-installs dependencies if missing
"""

# ---------- Auto-install missing dependencies ----------
import sys
import subprocess

def ensure_package(pkg, import_name=None):
    try:
        __import__(import_name or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
        __import__(import_name or pkg)

# Required packages
ensure_package("pandas")
ensure_package("openpyxl")
ensure_package("numpy")

import pandas as pd
import numpy as np
import re
from datetime import datetime
from tkinter import Tk, filedialog, messagebox
import os

# --------------- helpers ----------------

def pick_fs_and_cards():
    root = Tk()
    root.withdraw()
    messagebox.showinfo("Select FS CSV", "Pick the FS CSV file to annotate.")
    fs_path = filedialog.askopenfilename(
        title="Select FS CSV",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    if not fs_path:
        raise SystemExit("No FS CSV selected.")

    messagebox.showinfo("Select Credit Card Excel", "Pick one or more Credit Card Excel files.")
    cc_paths = filedialog.askopenfilenames(
        title="Select Credit Card Excel files",
        filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
    )
    if not cc_paths:
        raise SystemExit("No Credit Card Excel files selected.")

    save_path = filedialog.asksaveasfilename(
        title="Save output CSV as...",
        defaultextension=".csv",
        initialfile="fs_matched_classified.csv",
        filetypes=[("CSV Files", "*.csv")]
    )
    if not save_path:
        raise SystemExit("No output path selected.")
    return fs_path, cc_paths, save_path


def find_first_col(df, candidates, normalize=True):
    cols = list(df.columns)
    if normalize:
        norm = {re.sub(r"\s+", "", c).strip().lower(): c for c in cols}
        for cand in candidates:
            key = re.sub(r"\s+", "", cand).strip().lower()
            if key in norm:
                return norm[key]
    else:
        for cand in candidates:
            if cand in cols:
                return cand
    for c in cols:
        compact = re.sub(r"\s+", "", c).lower()
        for cand in candidates:
            if re.sub(r"\s+", "", cand).lower() in compact:
                return c
    return None


def extract_last4_from_sheet(df):
    pattern = re.compile(r"(credit\s*card\s*no|card\s*no|card\s*number)\s*[:#-]?\s*(\d{4})", re.I)
    sample_area = df.head(6).iloc[:, : min(6, df.shape[1])]
    for val in sample_area.values.flatten():
        if isinstance(val, str):
            m = pattern.search(val)
            if m:
                return m.group(2)
    return None


def extract_last4_from_filename(path):
    name_no_ext = os.path.splitext(os.path.basename(path))[0]
    tokens = re.split(r"[_\-\s]+", name_no_ext.lower())
    all4 = re.findall(r"\b(\d{4})\b", name_no_ext)

    def likely_year(s):
        n = int(s)
        return 1900 <= n <= 2099

    cardish = {"card", "credit", "visa", "amex", "chase", "master", "mastercard", "company"}
    best, best_score = None, -1
    for m in all4:
        if likely_year(m):
            continue
        score = 0
        for t in tokens:
            if m in t:
                score += 2
            if any(k in t for k in cardish):
                score += 1
        if score > best_score:
            best, best_score = m, score
    return best


def parse_amount(x):
    if isinstance(x, (int, float, np.number)):
        return float(x)
    if isinstance(x, str):
        s = re.sub(r"[^\d\.\-\+]", "", x.strip().replace(",", ""))
        if s in ("", "-", "+"):
            return np.nan
        try:
            return float(s)
        except ValueError:
            return np.nan
    return np.nan


def parse_date(x):
    try:
        return pd.to_datetime(x, errors="coerce").date()
    except Exception:
        pass
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%b %d, %Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(str(x), fmt).date()
        except Exception:
            continue
    return pd.NaT


def normalize_food_mapping(cat_text):
    if not isinstance(cat_text, str):
        return None
    s = cat_text.lower().strip()
    if "employee benefits -meals" in s or "staff meal" in s:
        return "Office Expense:Meal"
    if "business trip" in s:
        return "Travel Expense:Meals "
    if "client meeting" in s or "customer meeting" in s:
        return "Meals and Entertainment "
    return None


def is_food_and_drink(row, food_cat_cols):
    for c in food_cat_cols:
        if c and c in row and isinstance(row[c], str):
            if re.search(r"\b(food|drink)\b", row[c], re.I):
                return True
    return False


def main():
    fs_path, cc_paths, save_path = pick_fs_and_cards()

    fs = pd.read_csv(fs_path, dtype=str)
    fs_card_col = find_first_col(fs, ["Card", "Card Last4", "Card Number (Last 4)", "Card No", "Last4"])
    fs_date_col = find_first_col(fs, ["Date", "Transaction Date", "Post Date"])
    fs_amt_col  = find_first_col(fs, ["Amount", "Cost", "Debit", "Charge Amount", "Total"])

    if not (fs_card_col and fs_date_col and fs_amt_col):
        messagebox.showerror("FS CSV columns missing", "Could not find needed columns in FS.")
        return

    fs["_card_last4"] = fs[fs_card_col].astype(str).str.extract(r"(\d{4})", expand=False)
    fs["_date"] = fs[fs_date_col].apply(parse_date)
    fs["_amount"] = fs[fs_amt_col].apply(parse_amount)
    fs["_amt_r"] = fs["_amount"].round(2)

    cc_rows = []

    for path in cc_paths:
        df = pd.read_excel(path, dtype=str)
        last4 = extract_last4_from_sheet(df) or extract_last4_from_filename(path)
        if not last4:
            cand_last4_col = find_first_col(df, ["Card", "Card Last4", "Card No", "Last4"])
            if cand_last4_col is not None:
                tmp = df[cand_last4_col].astype(str).str.extract(r"(\d{4})", expand=False)
                last4 = tmp.dropna().iloc[0] if not tmp.dropna().empty else None

        date_col = find_first_col(df, ["Date", "Transaction Date", "Posted Date", "Post Date"])
        amt_col  = find_first_col(df, ["Amount", "Cost", "Debit", "Charge Amount", "Total", "Amount (USD)"])
        food_scope_cols = [
            find_first_col(df, ["Category", "Main Category", "Top Category", "Group"]),
            find_first_col(df, ["Type", "Subcategory", "Sub Category", "Memo", "Notes"])
        ]
        subcat_col = find_first_col(df, ["Category", "Type", "Subcategory", "Notes", "Memo"])

        if not (date_col and amt_col):
            continue

        tmp = pd.DataFrame({
            "_date": df[date_col].apply(parse_date),
            "_amount": df[amt_col].apply(parse_amount),
        })
        tmp["_amt_r"] = tmp["_amount"].round(2)
        tmp["_last4"] = last4

        for c in set([c for c in food_scope_cols if c] + ([subcat_col] if subcat_col else [])):
            tmp[c] = df[c]

        mask_food = tmp.apply(lambda r: is_food_and_drink(r, [c for c in food_scope_cols if c]), axis=1)
        tmp = tmp[mask_food]
        tmp["__classification"] = tmp[subcat_col].apply(normalize_food_mapping) if subcat_col else None
        tmp = tmp[tmp["__classification"].notna()]

        cc_rows.append(tmp)

    if not cc_rows:
        fs["Classification"] = ""
        fs.to_csv(save_path, index=False)
        messagebox.showinfo("Saved", f"Saved: {save_path}")
        return

    cc_all = pd.concat(cc_rows, ignore_index=True)
    fs["__key_signed"] = fs["_card_last4"].fillna("") + "|" + fs["_date"].astype(str) + "|" + fs["_amt_r"].astype(str)
    cc_all["__key_signed"] = cc_all["_last4"].fillna("") + "|" + cc_all["_date"].astype(str) + "|" + cc_all["_amt_r"].astype(str)
    fs["__key_abs"] = fs["_card_last4"].fillna("") + "|" + fs["_date"].astype(str) + "|" + fs["_amt_r"].abs().astype(str)
    cc_all["__key_abs"] = cc_all["_last4"].fillna("") + "|" + cc_all["_date"].astype(str) + "|" + cc_all["_amt_r"].abs().astype(str)

    left = fs.merge(cc_all[["__key_signed", "__classification"]], how="left", on="__key_signed")
    missing_mask = left["__classification"].isna()
    if missing_mask.any():
        fill = fs.loc[missing_mask, ["__key_abs"]].merge(
            cc_all[["__key_abs", "__classification"]],
            how="left",
            on="__key_abs"
        )["__classification"]
        left.loc[missing_mask, "__classification"] = fill.values

    left["Classification"] = left["__classification"].fillna("")
    out_cols = list(fs.columns.drop(["__key_signed", "__key_abs"], errors="ignore")) + ["Classification"]
    out = left[out_cols]
    out.to_csv(save_path, index=False)
    messagebox.showinfo("Done", f"Annotated FS saved to:\n{save_path}")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        print(e)
    except Exception as ex:
        import traceback
        traceback.print_exc()
