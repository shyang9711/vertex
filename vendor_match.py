import os
import re
import json
import difflib
import pandas as pd
from tkinter import (
    Tk, Button, Label, filedialog, messagebox, StringVar,
    DISABLED, NORMAL, Toplevel, Entry, Listbox, SINGLE, END
)
import sys
import subprocess
import importlib

required_packages = [
    "pandas",
    "openpyxl",
    "xlrd",
]

def install_if_missing(pkg):
    try:
        importlib.import_module(pkg.split("==")[0])
    except ImportError:
        print(f"Installing missing dependency: {pkg} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

for package in required_packages:
    install_if_missing(package)


# ========= Paths & persistence helpers =========
def _script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()

def _rules_dir() -> str:
    p = os.path.join(APP_DIR, "data", "match_rules")
    os.makedirs(p, exist_ok=True)  # ensure folder exists
    return p

def _vendor_lists_dir() -> str:
    return os.path.join(APP_DIR, "data", "vendor_lists")

APP_DIR = _script_dir()
COMPANY_LIST_FILENAME = "company_list.json"
COMPANY_LIST_PATH = os.path.join(APP_DIR, COMPANY_LIST_FILENAME)

def slugify(name: str) -> str:
    s = re.sub(r"\s+", "_", name.strip())
    s = re.sub(r"[^A-Za-z0-9_]+", "", s)
    return s.lower()

def company_rules_path(company_name: str) -> str:
    slug = slugify(company_name)
    return os.path.join(_rules_dir(), f"vendor_rules_{slug}.json")

def load_company_list():
    if os.path.exists(COMPANY_LIST_PATH):
        try:
            with open(COMPANY_LIST_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and "companies" in data and isinstance(data["companies"], list):
                return sorted({str(x).strip() for x in data["companies"] if str(x).strip()})
            if isinstance(data, list):
                return sorted({str(x).strip() for x in data if str(x).strip()})
        except Exception as e:
            messagebox.showwarning("Company List", f"Could not load {COMPANY_LIST_FILENAME}:\n{e}")
    return []

def save_company_list(companies):
    try:
        with open(COMPANY_LIST_PATH, "w", encoding="utf-8") as f:
            json.dump({"companies": companies}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("Company List", f"Could not save {COMPANY_LIST_FILENAME}:\n{e}")

def _safe_initial_filename(company_name: str | None) -> str:
    base = (company_name or "Matched Vendors").strip()
    # replace Windows-illegal characters
    base = re.sub(r'[\\/:*?"<>|]', "_", base)
    return f"{base}.csv"

def load_rules_from_disk(rules_path):
    if os.path.exists(rules_path):
        try:
            with open(rules_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            rules = []
            for item in data:
                phrase = str(item.get("phrase", "")).strip()
                vendor = str(item.get("vendor", "")).strip()
                if phrase and vendor:
                    rules.append({"phrase": phrase, "vendor": vendor})
            return rules
        except Exception as e:
            messagebox.showwarning("Rules Load", f"Could not load rules:\n{e}")
    return []

def save_rules_to_disk(rules_path, rules):
    try:
        os.makedirs(os.path.dirname(rules_path) or APP_DIR, exist_ok=True)
        with open(rules_path, "w", encoding="utf-8") as f:
            json.dump(rules, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("Rules Save", f"Could not save rules:\n{e}")

def _safe_initial_filename(company_name: str | None) -> str:
    base = (company_name or "Matched Vendors").strip()
    base = re.sub(r'[\\/:*?"<>|]', "_", base)
    return f"{base}.csv"

def _compact(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())

def _candidate_vendor_filenames(company_name: str) -> list[str]:
    """
    Generate likely filenames for the vendor list.
    """
    slug = slugify(company_name)
    base = company_name.strip()
    return [
        f"{base}.csv",
        f"{slug}.csv",
        f"{slug}_vendor_list.csv",
        f"{base}.xlsx",
        f"{slug}.xlsx",
        f"{slug}_vendor_list.xlsx",
    ]

def _find_auto_vendor_list(company_name: str) -> str | None:
    """
    Look in APP_DIR/vendor_lists for a file that matches the company.
    Priority:
      1) Exact candidate names (csv/xlsx)
      2) Any file whose *compact* name contains the compact company name
    """
    vdir = _vendor_lists_dir()
    if not os.path.isdir(vdir):
        return None

    # 1) exact/slug candidates first
    cands = _candidate_vendor_filenames(company_name)
    for fn in cands:
        p = os.path.join(vdir, fn)
        if os.path.isfile(p):
            return p

    # 2) fuzzy "contains company" match
    want = _compact(company_name)
    for name in sorted(os.listdir(vdir)):
        if os.path.splitext(name)[1].lower() not in {".csv", ".xlsx", ".xls"}:
            continue
        if want and (_compact(name).find(want) >= 0):
            return os.path.join(vdir, name)

    return None


# ========= Normalization & trailing-ID helpers =========
CORP_SUFFIXES = {
    "INC", "INCORPORATED", "CO", "COMPANY", "CORP", "CORPORATION",
    "LLC", "L.L.C", "LTD", "LIMITED", "PLC", "GMBH", "SAS", "SA", "BV", "PTY", "PTE"
}

def _is_id_like_token(tok: str) -> bool:
    if not tok:
        return False
    if re.fullmatch(r"\d{4,}", tok):
        return True
    if re.search(r"\d{4,}", tok):
        return True
    return False

def _remove_trailing_id_token(raw: str) -> str:
    if raw is None:
        return ""
    s = str(raw).upper()
    m = re.search(r'([A-Z0-9]+)[\W_]*$', s)
    if not m:
        return s
    last_tok = m.group(1)
    if _is_id_like_token(last_tok):
        s = s[:m.start(1)]
    return s

def normalize_text(s: str) -> str:
    s = _remove_trailing_id_token(s)
    s = s.replace("&", " AND ")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    parts = [p for p in s.split() if p and p not in CORP_SUFFIXES]
    return " ".join(parts)

# ========= Vendor list & matching =========
def parse_vendor_list(vendors_df: pd.DataFrame):
    if "Vendor" not in vendors_df.columns:
        raise ValueError("Vendor file must have a 'Vendor' column.")

    entries = []
    for _, row in vendors_df.iterrows():
        base = str(row["Vendor"]).strip()
        if not base or base.lower() == "nan":
            continue

        synonyms = {base}
        for alt_col in ("Aliases", "Alias", "Keywords"):
            if alt_col in vendors_df.columns and pd.notna(row.get(alt_col, None)):
                raw = str(row[alt_col])
                for piece in re.split(r"[|;]", raw):
                    piece = piece.strip()
                    if piece:
                        synonyms.add(piece)

        for syn in synonyms:
            norm = normalize_text(syn)
            if norm:
                entries.append((base, norm, syn))

    entries.sort(key=lambda t: len(t[1]), reverse=True)
    return entries

def detect_description_column(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    lows = [c.strip().lower() for c in cols]
    lowset = set(lows)

    # --- US Bank SkyPass format: use Name (ignore Memo) ---
    if {"date", "transaction", "name", "memo", "amount"}.issubset(lowset):
        return cols[lows.index("name")]

    # 1) Exact 'Description'
    for c in cols:
        if c.strip().lower() == "description":
            return c

    # 2) Heuristics
    heuristics = ["description", "desc", "details", "name", "memo", "narrative"]
    for h in heuristics:
        for c in cols:
            if h in c.strip().lower():
                return c

    # 3) Fallback: first column
    return cols[0]

def detect_amount_column(df: pd.DataFrame) -> str:
    for c in df.columns:
        if c.strip().lower() in ["amount", "amt"]:
            return c
    for h in ["amount", "amt", "payment", "value"]:
        for c in df.columns:
            if h in c.strip().lower():
                return c
    return None

def load_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".csv", ".txt"]:
        try:
            return pd.read_csv(path)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin-1")
    else:
        return pd.read_excel(path)

def find_best_vendor(description_norm: str, vendor_entries) -> str:
    if not description_norm:
        return ""
    best_vendor = ""
    best_ratio = 0.0
    for canonical, syn_norm, _orig in vendor_entries:
        if not syn_norm:
            continue
        if syn_norm in description_norm:
            return canonical
        ratio = difflib.SequenceMatcher(None, syn_norm, description_norm).ratio()
        if ratio >= 0.8 and ratio > best_ratio:
            best_vendor = canonical
            best_ratio = ratio
    return best_vendor

# ===== Accounts (Vendor -> Account) persistence =====
def company_accounts_path(company_name: str) -> str:
    slug = slugify(company_name)
    return os.path.join(_rules_dir(), f"accounts_{slug}.json")

def load_accounts_from_disk(accounts_path):
    if os.path.exists(accounts_path):
        try:
            with open(accounts_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {str(k): str(v) for k, v in data.items()}
        except Exception as e:
            messagebox.showwarning("Accounts Load", f"Could not load accounts:\n{e}")
    return {}

def save_accounts_to_disk(accounts_path, mapping: dict):
    try:
        os.makedirs(os.path.dirname(accounts_path) or APP_DIR, exist_ok=True)
        with open(accounts_path, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("Accounts Save", f"Could not save accounts:\n{e}")

# ========= Main Tkinter App =========
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Vendor Matcher")

        # --- state & StringVars ---
        self.tx_path = None
        self.vendor_path = None
        self.vendor_entries = None
        self.vendor_names = []
        self.manual_rules = []
        self.company_list = load_company_list()
        self.current_company = None
        self.account_map = {}


        self.tx_label_var = StringVar(value="No transactions file selected")
        self.vendor_label_var = StringVar(value="No vendor file selected")
        self.status_var = StringVar(value="")
        self.rules_label_var = StringVar(value="Manual rules: (none)")
        self.company_label_var = StringVar(value="Company: (none)")
        self.rules_file_hint = StringVar(value="Rules file: (select company)")

        # --- Build the main UI FIRST ---
        Button(root, text="Select Transactions (Excel/CSV)", command=self.pick_transactions).grid(row=0, column=0, padx=10, pady=(12,6), sticky="w")
        Label(root, textvariable=self.tx_label_var, wraplength=540, fg="#333").grid(row=1, column=0, padx=12, pady=(0,8), sticky="w")

        Button(root, text="Select Vendor List", command=self.pick_vendor).grid(row=2, column=0, padx=10, pady=(6,6), sticky="w")
        Label(root, textvariable=self.vendor_label_var, wraplength=540, fg="#333").grid(row=3, column=0, padx=12, pady=(0,8), sticky="w")

        Button(root, text="Company…", command=self.select_company_dialog).grid(row=4, column=0, padx=10, pady=(6,0), sticky="w")
        Label(root, textvariable=self.company_label_var, wraplength=540, fg="#333").grid(row=5, column=0, padx=12, pady=(2,2), sticky="w")
        Label(root, textvariable=self.rules_file_hint, wraplength=540, fg="#888").grid(row=6, column=0, padx=12, pady=(0,8), sticky="w")

        Button(root, text="Manage Rules…", command=self.manage_rules_dialog).grid(row=7, column=0, padx=10, pady=(2,0), sticky="w")
        Label(root, textvariable=self.rules_label_var, wraplength=540, fg="#555").grid(row=8, column=0, padx=12, pady=(2,10), sticky="w")
        Button(root, text="Manage Accounts…", command=self.manage_accounts_dialog).grid(row=7, column=1, padx=10, pady=(2,0), sticky="w")

        self.run_btn = Button(root, text="Run and Save", command=self.run_and_save, state=DISABLED)
        self.run_btn.grid(row=9, column=0, padx=10, pady=(6,12), sticky="w")

        Label(root, textvariable=self.status_var, fg="#006400").grid(row=10, column=0, padx=12, pady=(0,12), sticky="w")

        self.maybe_enable_run()

        self.select_company_dialog(initial=True)


    # ===== Company selection & management =====
    def select_company_dialog(self, initial=False):
        dlg = Toplevel(self.root)
        dlg.title("Select Company")

        dlg.transient(self.root)
        dlg.grab_set()
        dlg.lift()
        dlg.focus_force()
        dlg.attributes("-topmost", True)
        dlg.after(100, lambda: dlg.attributes("-topmost", False))

        # ---- Lists and controls ----
        lb = Listbox(dlg, selectmode=SINGLE, width=50, height=10)
        lb.grid(row=0, column=0, columnspan=3, padx=10, pady=(10,6), sticky="w")

        def refresh_lb():
            lb.delete(0, END)
            for name in self.company_list:
                lb.insert(END, name)

        refresh_lb()

        def set_company_and_close(name):
            self.set_company(name)
            if initial:
                self.root.deiconify()
            dlg.destroy()

            # --- Auto-load vendor list for this company if present ---
            try:
                auto_v = _find_auto_vendor_list(name)
                if auto_v:
                    vendors_df = load_table(auto_v)
    
                    # Normalize common header spellings to the canonical "Vendor"
                    norm_map = {c: re.sub(r"\s+", "", str(c)).strip().lower() for c in vendors_df.columns}
                    inv = {v: k for k, v in norm_map.items()}
                    if "vendor" in inv and inv["vendor"] != "Vendor":
                        vendors_df.rename(columns={inv["vendor"]: "Vendor"}, inplace=True)
                    for alt in ("aliases", "alias", "keywords"):
                        if alt in inv and inv[alt] != alt.capitalize():
                            vendors_df.rename(columns={inv[alt]: alt.capitalize()}, inplace=True)
    
                    self.vendor_entries = parse_vendor_list(vendors_df)
                    if "Vendor" in vendors_df.columns:
                        names = vendors_df["Vendor"].dropna().astype(str).map(str.strip)
                        self.vendor_names = sorted({n for n in names if n})
                    else:
                        self.vendor_names = []
    
                    self.vendor_path = auto_v
                    self.vendor_label_var.set(auto_v)
                    self.maybe_enable_run()
                else:
                    # Clear previous vendor info so user knows to pick one
                    self.vendor_entries = None
                    self.vendor_names = []
                    self.vendor_path = None
                    self.vendor_label_var.set("No vendor file selected")
                    self.maybe_enable_run()
            except Exception as e:
                messagebox.showwarning("Vendor Auto-Load",
                                       f"Found vendor list but failed to load:\n{e}")


        def do_use():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Select Company", "Pick a company or add one.")
                return
            set_company_and_close(self.company_list[sel[0]])

        def do_add():
            sub = Toplevel(dlg)
            sub.title("Add Company")
            Label(sub, text="Company name:").grid(row=0, column=0, padx=10, pady=(10,4), sticky="w")
            name_e = Entry(sub, width=40); name_e.grid(row=1, column=0, padx=10, pady=(0,8), sticky="w")

            def add_and_close():
                name = name_e.get().strip()
                if not name:
                    messagebox.showerror("Add Company", "Name required.")
                    return
                if name in self.company_list:
                    messagebox.showerror("Add Company", "Company already exists.")
                    return
                self.company_list.append(name)
                self.company_list = sorted(set(self.company_list))
                save_company_list(self.company_list)
                
                save_rules_to_disk(company_rules_path(name), [])
                refresh_lb()
                sub.destroy()

            Button(sub, text="Add", command=add_and_close).grid(row=2, column=0, padx=10, pady=(2,10), sticky="w")
            sub.grab_set()
            name_e.focus_set()

        def do_delete():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Delete Company", "Select a company to delete.")
                return
            name = self.company_list[sel[0]]
            if not messagebox.askyesno("Delete Company", f"Delete '{name}' and its rules file?"):
                return
                
            rp = company_rules_path(name)
            try:
                if os.path.exists(rp):
                    os.remove(rp)
            except Exception as e:
                messagebox.showwarning("Delete Rules", f"Could not delete rules file:\n{e}")
                
            self.company_list.remove(name)
            save_company_list(self.company_list)
            refresh_lb()
            
            if self.current_company == name:
                self.current_company = None
                self.manual_rules = []
                self.company_label_var.set("Company: (none)")
                self.rules_file_hint.set("Rules file: (select company)")
                self._refresh_rules_label()

        Button(dlg, text="Use", command=do_use).grid(row=1, column=0, padx=10, pady=(4,10), sticky="w")
        Button(dlg, text="Add", command=do_add).grid(row=1, column=1, padx=6, pady=(4,10), sticky="w")
        Button(dlg, text="Delete", command=do_delete).grid(row=1, column=2, padx=6, pady=(4,10), sticky="w")

        if initial and len(self.company_list) == 1:
            set_company_and_close(self.company_list[0])
            return

        def on_close():
            if initial and not self.current_company:
                try:
                    dlg.destroy()
                finally:
                    self.root.destroy()
            else:
                dlg.destroy()

        dlg.protocol("WM_DELETE_WINDOW", on_close)

        dlg.wait_window(dlg)


    def set_company(self, name):
        self.current_company = name
        self.company_label_var.set(f"Company: {name}")
        rp = company_rules_path(name)
        ap = company_accounts_path(name)
        self.rules_file_hint.set(f"Rules file: {os.path.basename(rp)}")
        self.manual_rules = load_rules_from_disk(rp)
        self.account_map  = load_accounts_from_disk(ap)
        self._refresh_rules_label()
        self.maybe_enable_run()

    # ===== File pickers =====
    def pick_transactions(self):
        path = filedialog.askopenfilename(
            title="Select Transactions file (Excel or CSV)",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv *.txt"), ("All files", "*.*")]
        )
        if path:
            self.tx_path = path
            self.tx_label_var.set(path)
            self.maybe_enable_run()

    def pick_vendor(self):
        # Default start directory
        vendor_lists_dir = _vendor_lists_dir()
        initial_dir = vendor_lists_dir if os.path.isdir(vendor_lists_dir) else APP_DIR
    
        path = filedialog.askopenfilename(
            title="Select Vendor List (CSV or Excel)",
            filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if path:
            self.vendor_path = path
            self.vendor_label_var.set(path)
            self.maybe_enable_run()
            try:
                vendors_df = load_table(self.vendor_path)
    
                # normalize headers so parse_vendor_list() sees "Vendor"
                norm_map = {c: re.sub(r"\s+", "", str(c)).strip().lower() for c in vendors_df.columns}
                inv = {v: k for k, v in norm_map.items()}
                if "vendor" in inv and inv["vendor"] != "Vendor":
                    vendors_df.rename(columns={inv["vendor"]: "Vendor"}, inplace=True)
                for alt in ("aliases", "alias", "keywords"):
                    if alt in inv and inv[alt] != alt.capitalize():
                        vendors_df.rename(columns={inv[alt]: alt.capitalize()}, inplace=True)
    
                self.vendor_entries = parse_vendor_list(vendors_df)
                if "Vendor" in vendors_df.columns:
                    names = vendors_df["Vendor"].dropna().astype(str).map(str.strip)
                    self.vendor_names = sorted({n for n in names if n})
                else:
                    self.vendor_names = []
            except Exception as e:
                messagebox.showerror("Vendor File", f"Failed to parse vendor file:\n{e}")
                self.vendor_entries = None
                self.vendor_names = []


    def maybe_enable_run(self):
        if not hasattr(self, "run_btn"):
            return
        if self.tx_path and self.vendor_path and self.current_company:
            self.run_btn.config(state=NORMAL)
        else:
            self.run_btn.config(state=DISABLED)

    # ===== Manual rules =====
    def _refresh_rules_label(self):
        if not self.manual_rules:
            self.rules_label_var.set("Manual rules: (none)")
            return
        preview = "; ".join([f"'{r['phrase']}'→{r['vendor']}" for r in self.manual_rules[:5]])
        if len(self.manual_rules) > 5:
            preview += f"; (+{len(self.manual_rules)-5} more)"
        self.rules_label_var.set(f"Manual rules: {preview}")

    def _apply_manual_rules(self, raw_desc: str) -> str:
        txt = "" if raw_desc is None else str(raw_desc).upper()
        for r in self.manual_rules:
            if r["phrase"].upper() in txt:
                return r["vendor"]
        return ""

    def manage_rules_dialog(self):
        if not self.current_company:
            messagebox.showinfo("Manage Rules", "Please select a company first.")
            return
        rp = company_rules_path(self.current_company)

        win = Toplevel(self.root)
        win.title(f"Manage Rules — {self.current_company}")

        lb = Listbox(win, selectmode=SINGLE, width=64, height=10)
        lb.grid(row=0, column=0, columnspan=3, padx=10, pady=(10,6), sticky="w")

        def refresh_lb():
            lb.delete(0, END)
            for r in self.manual_rules:
                lb.insert(END, f"{r['phrase']}  →  {r['vendor']}")
        refresh_lb()

        # ---- Add
        def do_add():
            sub = Toplevel(win)
            sub.title("Add Rule")

            Label(sub, text="If description contains (word/phrase):").grid(row=0, column=0, padx=10, pady=(10,4), sticky="w")
            phrase_e = Entry(sub, width=48); phrase_e.grid(row=1, column=0, padx=10, pady=(0,10), sticky="w")

            Label(sub, text="Search vendor:").grid(row=2, column=0, padx=10, pady=(0,2), sticky="w")
            search_e = Entry(sub, width=48); search_e.grid(row=3, column=0, padx=10, pady=(0,6), sticky="w")

            vendor_lb = Listbox(sub, selectmode=SINGLE, width=48, height=8)
            vendor_lb.grid(row=4, column=0, padx=10, pady=(0,6), sticky="w")

            Label(sub, text="Chosen vendor (or type):").grid(row=5, column=0, padx=10, pady=(2,2), sticky="w")
            vendor_e = Entry(sub, width=48); vendor_e.grid(row=6, column=0, padx=10, pady=(0,8), sticky="w")

            def update_vendor_list(*_):
                q = search_e.get().strip().lower()
                vendor_lb.delete(0, END)
                for name in self.vendor_names:
                    if q in name.lower():
                        vendor_lb.insert(END, name)

            def use_selected(_=None):
                sel = vendor_lb.curselection()
                if not sel:
                    return
                vendor_e.delete(0, END)
                vendor_e.insert(0, vendor_lb.get(sel[0]))

            search_e.bind("<KeyRelease>", update_vendor_list)
            vendor_lb.bind("<Double-Button-1>", use_selected)
            update_vendor_list()

            def add_and_close():
                phrase = phrase_e.get().strip()
                vendor = vendor_e.get().strip()
                if not phrase or not vendor:
                    messagebox.showerror("Invalid Rule", "Both fields are required.")
                    return
                self.manual_rules.append({"phrase": phrase, "vendor": vendor})
                save_rules_to_disk(rp, self.manual_rules)  # immediate save
                self._refresh_rules_label()
                refresh_lb()
                sub.destroy()

            Button(sub, text="Use Selected", command=use_selected).grid(row=7, column=0, padx=10, pady=(0,6), sticky="w")
            Button(sub, text="Add", command=add_and_close).grid(row=8, column=0, padx=10, pady=(0,10), sticky="w")
            sub.grab_set(); phrase_e.focus_set()

        def do_edit():
            idxs = lb.curselection()
            if not idxs:
                messagebox.showinfo("Edit Rule", "Select a rule to edit.")
                return
            i = idxs[0]
            rule = self.manual_rules[i]

            sub = Toplevel(win)
            sub.title("Edit Rule")

            Label(sub, text="If description contains (word/phrase):").grid(row=0, column=0, padx=10, pady=(10,4), sticky="w")
            phrase_e = Entry(sub, width=48); phrase_e.insert(0, rule["phrase"]); phrase_e.grid(row=1, column=0, padx=10, pady=(0,10), sticky="w")

            Label(sub, text="Search vendor:").grid(row=2, column=0, padx=10, pady=(0,2), sticky="w")
            search_e = Entry(sub, width=48); search_e.grid(row=3, column=0, padx=10, pady=(0,6), sticky="w")

            vendor_lb = Listbox(sub, selectmode=SINGLE, width=48, height=8)
            vendor_lb.grid(row=4, column=0, padx=10, pady=(0,6), sticky="w")

            Label(sub, text="Chosen vendor (or type):").grid(row=5, column=0, padx=10, pady=(2,2), sticky="w")
            vendor_e = Entry(sub, width=48); vendor_e.insert(0, rule["vendor"]); vendor_e.grid(row=6, column=0, padx=10, pady=(0,8), sticky="w")

            def update_vendor_list(*_):
                q = search_e.get().strip().lower()
                vendor_lb.delete(0, END)
                for name in self.vendor_names:
                    if q in name.lower():
                        vendor_lb.insert(END, name)

            def use_selected(_=None):
                sel = vendor_lb.curselection()
                if not sel:
                    return
                vendor_e.delete(0, END)
                vendor_e.insert(0, vendor_lb.get(sel[0]))

            search_e.bind("<KeyRelease>", update_vendor_list)
            vendor_lb.bind("<Double-Button-1>", use_selected)
            update_vendor_list()

            def save_and_close():
                phrase = phrase_e.get().strip()
                vendor = vendor_e.get().strip()
                if not phrase or not vendor:
                    messagebox.showerror("Invalid Rule", "Both fields are required.")
                    return
                self.manual_rules[i] = {"phrase": phrase, "vendor": vendor}
                save_rules_to_disk(rp, self.manual_rules)
                self._refresh_rules_label()
                refresh_lb()
                sub.destroy()

            Button(sub, text="Use Selected", command=use_selected).grid(row=7, column=0, padx=10, pady=(0,6), sticky="w")
            Button(sub, text="Save", command=save_and_close).grid(row=8, column=0, padx=10, pady=(0,10), sticky="w")
            sub.grab_set(); phrase_e.focus_set()

        def do_delete():
            idxs = lb.curselection()
            if not idxs:
                messagebox.showinfo("Delete Rule", "Select a rule to delete.")
                return
            i = idxs[0]
            confirm = messagebox.askyesno("Delete Rule", f"Delete this rule?\n\n{self.manual_rules[i]['phrase']} → {self.manual_rules[i]['vendor']}")
            if confirm:
                del self.manual_rules[i]
                save_rules_to_disk(rp, self.manual_rules)
                self._refresh_rules_label()
                refresh_lb()

        Button(win, text="Add", command=do_add).grid(row=1, column=0, padx=10, pady=(4,10), sticky="w")
        Button(win, text="Edit", command=do_edit).grid(row=1, column=1, padx=6, pady=(4,10), sticky="w")
        Button(win, text="Delete", command=do_delete).grid(row=1, column=2, padx=6, pady=(4,10), sticky="w")

        win.grab_set()

    def manage_accounts_dialog(self):
        if not self.current_company:
            messagebox.showinfo("Manage Accounts", "Please select a company first.")
            return
        ap = company_accounts_path(self.current_company)

        win = Toplevel(self.root)
        win.title(f"Manage Accounts — {self.current_company}")

        lb = Listbox(win, selectmode=SINGLE, width=64, height=12)
        lb.grid(row=0, column=0, columnspan=3, padx=10, pady=(10,6), sticky="w")

        def refresh_lb():
            lb.delete(0, END)
            for vendor in sorted(self.account_map.keys(), key=lambda x: x.lower()):
                lb.insert(END, f"{vendor}  →  {self.account_map[vendor]}")

        refresh_lb()

        def do_add():
            sub = Toplevel(win)
            sub.title("Add Vendor → Account")
            Label(sub, text="Search vendor:").grid(row=0, column=0, padx=10, pady=(10,2), sticky="w")
            search_e = Entry(sub, width=48); search_e.grid(row=1, column=0, padx=10, pady=(0,6), sticky="w")

            vendor_lb = Listbox(sub, selectmode=SINGLE, width=48, height=8)
            vendor_lb.grid(row=2, column=0, padx=10, pady=(0,6), sticky="w")

            Label(sub, text="Chosen vendor (or type):").grid(row=3, column=0, padx=10, pady=(2,2), sticky="w")
            vendor_e = Entry(sub, width=48); vendor_e.grid(row=4, column=0, padx=10, pady=(0,8), sticky="w")

            Label(sub, text="Account:").grid(row=5, column=0, padx=10, pady=(2,2), sticky="w")
            account_e = Entry(sub, width=48); account_e.grid(row=6, column=0, padx=10, pady=(0,8), sticky="w")

            def update_vendor_list(*_):
                q = search_e.get().strip().lower()
                vendor_lb.delete(0, END)
                for name in self.vendor_names:
                    if q in name.lower():
                        vendor_lb.insert(END, name)

            def use_selected(_=None):
                sel = vendor_lb.curselection()
                if not sel: return
                vendor_e.delete(0, END)
                vendor_e.insert(0, vendor_lb.get(sel[0]))

            search_e.bind("<KeyRelease>", update_vendor_list)
            vendor_lb.bind("<Double-Button-1>", use_selected)
            update_vendor_list()

            def add_and_close():
                vendor = vendor_e.get().strip()
                account = account_e.get().strip()
                if not vendor:
                    messagebox.showerror("Invalid", "Vendor is required.")
                    return
                    
                self.account_map[vendor] = account
                save_accounts_to_disk(ap, self.account_map)
                refresh_lb()
                sub.destroy()

            Button(sub, text="Use Selected", command=use_selected).grid(row=7, column=0, padx=10, pady=(0,6), sticky="w")
            Button(sub, text="Add", command=add_and_close).grid(row=8, column=0, padx=10, pady=(0,10), sticky="w")
            sub.grab_set(); search_e.focus_set()

        def do_edit():
            idxs = lb.curselection()
            if not idxs:
                messagebox.showinfo("Edit", "Select a mapping to edit.")
                return
                
            line = lb.get(idxs[0])
            if "  →  " not in line:
                return
            vendor_sel, account_sel = [x.strip() for x in line.split("  →  ", 1)]

            sub = Toplevel(win)
            sub.title("Edit Vendor → Account")
            Label(sub, text="Vendor:").grid(row=0, column=0, padx=10, pady=(10,2), sticky="w")
            vendor_e = Entry(sub, width=48); vendor_e.insert(0, vendor_sel); vendor_e.grid(row=1, column=0, padx=10, pady=(0,8), sticky="w")

            Label(sub, text="Account:").grid(row=2, column=0, padx=10, pady=(2,2), sticky="w")
            account_e = Entry(sub, width=48); account_e.insert(0, account_sel); account_e.grid(row=3, column=0, padx=10, pady=(0,8), sticky="w")

            def save_and_close():
                new_vendor = vendor_e.get().strip()
                new_account = account_e.get().strip()
                if not new_vendor:
                    messagebox.showerror("Invalid", "Vendor is required.")
                    return
                if new_vendor != vendor_sel and vendor_sel in self.account_map:
                    del self.account_map[vendor_sel]
                self.account_map[new_vendor] = new_account
                save_accounts_to_disk(ap, self.account_map)
                refresh_lb()
                sub.destroy()

            Button(sub, text="Save", command=save_and_close).grid(row=4, column=0, padx=10, pady=(0,10), sticky="w")
            sub.grab_set(); account_e.focus_set()

        def do_delete():
            idxs = lb.curselection()
            if not idxs:
                messagebox.showinfo("Delete", "Select a mapping to delete.")
                return
            line = lb.get(idxs[0])
            if "  →  " not in line:
                return
            vendor_sel = line.split("  →  ", 1)[0].strip()
            if messagebox.askyesno("Delete", f"Delete mapping for '{vendor_sel}'?"):
                if vendor_sel in self.account_map:
                    del self.account_map[vendor_sel]
                    save_accounts_to_disk(ap, self.account_map)
                    refresh_lb()

        Button(win, text="Add", command=do_add).grid(row=1, column=0, padx=10, pady=(4,10), sticky="w")
        Button(win, text="Edit", command=do_edit).grid(row=1, column=1, padx=6, pady=(4,10), sticky="w")
        Button(win, text="Delete", command=do_delete).grid(row=1, column=2, padx=6, pady=(4,10), sticky="w")

        win.grab_set()


    # ===== Run pipeline =====
    def run_and_save(self):
        try:
            self.status_var.set("Loading files...")
            self.root.update_idletasks()

            if not (self.tx_path and self.vendor_path and self.current_company):
                messagebox.showerror("Missing", "Select company, transactions file, and vendor list.")
                return

            tx_df = load_table(self.tx_path)
            vendors_df = load_table(self.vendor_path)

            desc_col = detect_description_column(tx_df)
            amount_col = detect_amount_column(tx_df)
            vendor_entries = parse_vendor_list(vendors_df)
            if not vendor_entries:
                messagebox.showerror("Vendor List", "No usable vendor names found.")
                return
            if not amount_col:
                messagebox.showwarning("No Amount Column", "No amount column detected. Proceeding without amount flip.")

            self.status_var.set("Matching vendors...")
            self.root.update_idletasks()

            vendors_out = []
            for raw_desc in tx_df[desc_col].fillna("").astype(str):
                manual = self._apply_manual_rules(raw_desc)
                if manual:
                    vendors_out.append(manual)
                else:
                    norm = normalize_text(raw_desc)
                    vendors_out.append(find_best_vendor(norm, vendor_entries))

            tx_df["Vendor"] = vendors_out
            accounts = [self.account_map.get(v, "") if isinstance(v, str) and v.strip() else "" for v in vendors_out]
            vendor_idx = tx_df.columns.get_loc("Vendor")
            tx_df.insert(vendor_idx + 1, "Account", accounts)

            if amount_col:
                amt = pd.to_numeric(tx_df[amount_col], errors="coerce")
                flipped = amt * -1
                acc_idx = tx_df.columns.get_loc("Account")
                tx_df.insert(acc_idx + 1, "FlippedAmount", flipped)
            else:
                tx_df["FlippedAmount"] = pd.NA

            matched_vendor = sum(1 for v in vendors_out if isinstance(v, str) and v.strip())
            matched_vendor_account = sum(
                1
                for v, a in zip(vendors_out, accounts)
                if isinstance(v, str) and v.strip() and isinstance(a, str) and a.strip()
            )
            total = len(vendors_out)

            out_path = filedialog.asksaveasfilename(
                title="Save CSV as",
                defaultextension=".csv",
                initialfile=_safe_initial_filename(self.current_company),
                filetypes=[("CSV", "*.csv")]
            )
            if not out_path:
                return

            tx_df.to_csv(out_path, index=False, encoding="utf-8-sig")
            for col in tx_df.select_dtypes(include=["datetime64[ns]"]).columns:
                tx_df[col] = tx_df[col].dt.strftime("%#m/%#d/%Y")

            self.status_var.set(
                f"Saved: {out_path}  |  Matched vendor: {matched_vendor}/{total}  |  "
                f"Matched vendor+account: {matched_vendor_account}/{total}"
                )
            messagebox.showinfo(
                "Done", 
                f"Saved:\n"
                f"{out_path}\n"
                f"Rows: {total}\n"
                f"Matched vendor: {matched_vendor}\n"
                f"Matched vendor + account: {matched_vendor_account}"
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("")


def main():
    root = Tk()
    root.geometry("780x460")
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
