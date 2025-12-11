import os
import re
import csv
import math
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
from dataclasses import dataclass
from typing import List, Tuple, Optional, Iterable, Set

# ----------------- Tkinter pickers -----------------
import tkinter as tk
from tkinter import filedialog, messagebox

# --- logging & progress ---
import logging
import time
from tqdm import tqdm

try:
    import torch
    TORCH_OK = True
except Exception:
    TORCH_OK = False

# Configure logging (INFO to see progress; switch to DEBUG for more detail)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

# ----------------- Optional fuzzy matcher -----------------
FUZZY_OK = True
try:
    from rapidfuzz import fuzz, process as rfprocess
except Exception:
    FUZZY_OK = False

# ----------------- OCR (EasyOCR: free CNN OCR, MIT license) -----------------
try:
    import easyocr
    READER = easyocr.Reader(['en'], gpu=False)  # set gpu=True if you have CUDA
except Exception as e:
    READER = None
    print("WARNING: EasyOCR not available. Install with: pip install easyocr\n", e)

# ----------------- Heuristics / Patterns -----------------
DPI      = 300
DETECT_DPI = 260
READTEXT_BATCH = 32
LOG_EVERY_PAGE = True
MIN_PAGE_INDEX_HINT = 5  # checks usually start later in the statement
SCAN_KEYWORDS = ["PAY TO THE ORDER OF", "TO THE ORDER OF", "CHECK", "DOLLARS", "MEMO"]

RE_DATE   = re.compile(r"\b(0?[1-9]|1[0-2])[\/\-\.](0?[1-9]|[12][0-9]|3[01])[\/\-\.](\d{2,4})\b")
RE_AMOUNT = re.compile(r"(?<![A-Za-z])\$\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?(?!\S)")
RE_CHKNO  = re.compile(r"\b(?:Check\s*#|Chk\s*#|#\s*|No\.?\s*)(\d{2,})\b", re.IGNORECASE)

AMOUNT_MIN = 1000.00  # filter threshold

@dataclass
class OCRWord:
    text: str
    box: np.ndarray  # 4-point polygon [[x1,y1],[x2,y2]...]

@dataclass
class CheckHit:
    page_index: int
    left_half_bbox: Tuple[int,int,int,int]
    info_strip_bbox: Tuple[int,int,int,int]
    payee_line_bbox: Tuple[int,int,int,int]

@dataclass
class ExtractedCheck:
    date: str
    check_number: str
    payee: str
    amount: str

# ----------------- OCR utilities -----------------
def pix_to_npimg(pix: fitz.Pixmap) -> np.ndarray:
    if pix.alpha:
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 4)
        arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
    else:
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 1:
            arr = cv2.cvtColor(arr, cv2.COLOR_GRAY2BGR)
        else:
            arr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    return arr

def render_page(doc: fitz.Document, i: int, dpi: int=DPI) -> np.ndarray:
    page = doc.load_page(i)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return pix_to_npimg(pix)

# ---------- OCR helpers ----------
def _rgb(img_bgr):
    return cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)

if TORCH_OK:
    @torch.no_grad()
    def easyocr_full(img_bgr, desc: str = "page OCR"):
        """OCR with timing + debug logs, returns list[OCRWord]."""
        if READER is None:
            return []
        t0 = time.perf_counter()
        results = READER.readtext(
            _rgb(img_bgr),
            detail=1,
            paragraph=False,
            batch_size=READTEXT_BATCH  # helps recognizer throughput
        )
        dt = time.perf_counter() - t0
        logging.debug(f"EasyOCR {desc}: {len(results)} boxes in {dt:.2f}s")
        out = []
        for box, text, conf in results:
            pts = np.array(box, dtype=np.float32)
            out.append(OCRWord(text=text.strip(), box=pts))
        return out
else:
    def easyocr_full(img_bgr, desc: str = "page OCR"):
        if READER is None:
            return []
        t0 = time.perf_counter()
        results = READER.readtext(
            _rgb(img_bgr),
            detail=1,
            paragraph=False,
            batch_size=READTEXT_BATCH
        )
        dt = time.perf_counter() - t0
        logging.debug(f"EasyOCR {desc}: {len(results)} boxes in {dt:.2f}s")
        out = []
        for box, text, conf in results:
            pts = np.array(box, dtype=np.float32)
            out.append(OCRWord(text=text.strip(), box=pts))
        return out
    
def render_page_fast(doc: fitz.Document, i: int, dpi: int=DETECT_DPI) -> np.ndarray:
    """Slightly lower DPI for detection OCR to speed up page scanning."""
    page = doc.load_page(i)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return pix_to_npimg(pix)

def ocr_text(img: np.ndarray) -> str:
    words = easyocr_full(img)
    return " ".join(w.text for w in words)

# ----------------- Geometry helpers -----------------
def poly_to_rect(poly: np.ndarray, pad: int=0, wmax: Optional[int]=None, hmax: Optional[int]=None) -> Tuple[int,int,int,int]:
    xs = poly[:,0]; ys = poly[:,1]
    x1, y1 = int(max(0, math.floor(xs.min()) - pad)), int(max(0, math.floor(ys.min()) - pad))
    x2, y2 = int(math.ceil(xs.max()) + pad), int(math.ceil(ys.max()) + pad)
    if wmax is not None:
        x1 = max(0, min(x1, wmax)); x2 = max(0, min(x2, wmax))
    if hmax is not None:
        y1 = max(0, min(y1, hmax)); y2 = max(0, min(y2, hmax))
    return x1, y1, x2, y2

def crop_rect(img: np.ndarray, rect: Tuple[int,int,int,int]) -> np.ndarray:
    x1,y1,x2,y2 = rect
    h,w = img.shape[:2]
    x1,x2 = max(0,min(x1,w)), max(0,min(x2,w))
    y1,y2 = max(0,min(y1,h)), max(0,min(y2,h))
    if x2<=x1 or y2<=y1: 
        return img[0:0,0:0]
    return img[y1:y2, x1:x2]

# ----------------- Page detection -----------------
def page_has_checks(words: List[OCRWord]) -> bool:
    t = " ".join(w.text for w in words).lower()
    return any(k.lower() in t for k in SCAN_KEYWORDS)

def find_left_check_and_regions(page_img: np.ndarray, words: List[OCRWord]) -> Optional[CheckHit]:
    """
    Heuristic for Bank-of-Hope-style layouts:
      1) Find left-most "PAY TO THE ORDER OF".
      2) Define a left-half check region around that Y band.
      3) Define info strip below it.
      4) Define payee line region to the right of the keyword.
    """
    H, W = page_img.shape[:2]
    if not words:
        return None
    candidates = [w for w in words if "pay" in w.text.lower() and "order" in w.text.lower()]
    if not candidates:
        return None
    hit = sorted(candidates, key=lambda w: np.min(w.box[:,0]))[0]
    x1,y1,x2,y2 = poly_to_rect(hit.box, pad=6, wmax=W, hmax=H)

    # Left check area
    left_x2 = int(W * 0.48)     # left half width
    band_h  = int(H * 0.22)     # vertical band height around the pay-to line
    ly1     = max(0, y1 - band_h//2)
    ly2     = min(H, y1 + band_h//2)
    left_check = (0, ly1, left_x2, ly2)

    # Info strip below the check image
    strip_h = int(H * 0.08)
    sy1     = min(H-1, ly2 + int(H*0.01))
    sy2     = min(H, sy1 + strip_h)
    info_strip = (0, sy1, left_x2, sy2)

    # Payee line region (handwritten area)
    payee_pad_h = int(H * 0.05)
    px1  = min(W-1, x2 + int(W*0.01))
    px2  = min(W, int(W * 0.95))
    py1  = max(0, y1 - payee_pad_h//2)
    py2  = min(H, y1 + payee_pad_h//2)
    payee_line = (px1, py1, px2, py2)

    return CheckHit(page_index=-1, left_half_bbox=left_check, info_strip_bbox=info_strip, payee_line_bbox=payee_line)

# ----------------- Field parsing -----------------
def parse_fields_from_info_text(text: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    # Amount: like $1,234.56
    amt = None
    m_amt = RE_AMOUNT.search(text.replace(" ", ""))
    if m_amt:
        amt = re.sub(r"\s+", "", m_amt.group(0))

    # Date: mm/dd/yy(yy)
    dte = None
    m_date = RE_DATE.search(text)
    if m_date:
        mm, dd, yy = m_date.group(1), m_date.group(2), m_date.group(3)
        if len(yy) == 2:
            yy = ("20" if int(yy) < 50 else "19") + yy
        dte = f"{int(mm):02d}/{int(dd):02d}/{yy}"

    # Check number: "Check #1234" / "No. 1234" / bare 4–8 digits
    chk = None
    m_chk = RE_CHKNO.search(text)
    if m_chk:
        chk = m_chk.group(1)
    else:
        m2 = re.search(r"\b(\d{4,8})\b", text)
        if m2:
            chk = m2.group(1)

    return dte, chk, amt

def amount_to_float(amt: str) -> Optional[float]:
    if not amt:
        return None
    s = amt.strip().replace("$","").replace(",","")
    try:
        return float(s)
    except Exception:
        return None

# ----------------- Vendor helpers -----------------
def normalize_name(s: str) -> str:
    s0 = s.lower()
    s0 = re.sub(r'[^a-z0-9& ]+', ' ', s0)
    s0 = re.sub(r'\s+', ' ', s0).strip()
    s0 = re.sub(r'\b(memo|dollars?)\b', '', s0).strip()
    return s0

def load_vendors_from_excel(xlsx_path: str) -> Set[str]:
    """
    Load vendor names from ALL sheets of the workbook.
    Preference order (first present & non-empty wins):
      1) 'Print On Check As', 'Print on Check As'
      2) 'Vendor', 'Vendor Name'
      3) 'Company', 'Company Name', 'Name', 'Payee'
    If none are present, fall back to ALL string-like columns across sheets.
    Returns a set of normalized names.
    """
    import pandas as pd
    import re

    prefer_groups = [
        ["Print On Check As", "Print on Check As"],
        ["Vendor", "Vendor Name"],
        ["Company", "Company Name"],
        ["Name", "Payee"]
    ]

    xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
    gathered = []

    # read all sheets; ignore completely empty ones
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xlsx_path, sheet_name=sheet, engine="openpyxl")
        except Exception:
            continue
        if df is None or df.shape[0] == 0 or df.shape[1] == 0:
            continue

        # try preferred columns in order
        picked = None
        for group in prefer_groups:
            for col in group:
                if col in df.columns:
                    series = df[col].dropna().astype(str).str.strip()
                    series = series[series.ne("")]
                    if len(series) > 0:
                        picked = series
                        break
            if picked is not None:
                break

        if picked is None:
            # fallback: any object dtype columns
            txt_cols = [c for c in df.columns if df[c].dtype == object]
            if txt_cols:
                tmp = []
                for c in txt_cols:
                    s = df[c].dropna().astype(str).str.strip()
                    s = s[s.ne("")]
                    tmp.append(s)
                if tmp:
                    picked = pd.concat(tmp, ignore_index=True)

        if picked is not None and len(picked) > 0:
            gathered.append(picked)

    if not gathered:
        return set()

    all_names = pd.concat(gathered, ignore_index=True).astype(str)

    # split aliases on common separators
    names = []
    for raw in all_names:
        for part in re.split(r"[;/,\n]+", raw):
            part = part.strip()
            if part:
                names.append(part)

    # normalize
    def _norm(s: str) -> str:
        s0 = s.lower()
        s0 = re.sub(r"[^a-z0-9& ]+", " ", s0)
        s0 = re.sub(r"\s+", " ", s0).strip()
        s0 = re.sub(r"\b(memo|dollars?)\b", "", s0).strip()
        return s0

    return { _norm(x) for x in names if _norm(x) }


def best_vendor_match(payee: str, vendor_norms: Set[str]) -> Tuple[str, int]:
    """Return (best_name, score 0-100)."""
    if not payee:
        return "", 0
    p = normalize_name(payee)
    if not p:
        return payee, 0
    # quick exact/substring checks
    if p in vendor_norms:
        return p, 100
    for v in vendor_norms:
        if p == v or p.startswith(v) or v.startswith(p):
            return v, 96
    # fuzzy (optional)
    if not FUZZY_OK or not vendor_norms:
        return payee, 0
    best = rfprocess.extractOne(p, vendor_norms, scorer=fuzz.token_set_ratio)
    if best:
        name, score, _ = best
        return name, int(score)
    return payee, 0

# ----------------- Per-PDF pipeline -----------------
def process_pdf(pdf_path: str) -> List[ExtractedCheck]:
    rows: List[ExtractedCheck] = []
    with fitz.open(pdf_path) as doc:
        n = doc.page_count
        page_order = list(range(max(0, MIN_PAGE_INDEX_HINT), n)) + list(range(0, max(0, MIN_PAGE_INDEX_HINT)))

        logging.info(f"Scanning '{os.path.basename(pdf_path)}' | pages: {n} | order: {page_order[:3]}...")

        for i in tqdm(page_order, desc=f"{os.path.basename(pdf_path)}", unit="pg"):
            try:
                t_page = time.perf_counter()

                # Fast pass (lower DPI) to detect check presence + locate payee y-band
                img_fast = render_page_fast(doc, i, dpi=DETECT_DPI)
                words_fast = easyocr_full(img_fast, desc=f"page {i} (fast)")
                has_checks = page_has_checks(words_fast)

                if LOG_EVERY_PAGE:
                    logging.info(f"Page {i}: has_checks={has_checks} (fast pass)")

                if not has_checks:
                    continue

                # Re-render at full DPI for accurate crops
                img_full = render_page(doc, i, dpi=DPI)
                # Recompute keyword positions at full DPI (optional but safer geometrically)
                words_full = easyocr_full(img_full, desc=f"page {i} (full)")
                hit = find_left_check_and_regions(img_full, words_full)
                if not hit:
                    logging.debug(f"Page {i}: no hit after region finding")
                    continue
                hit.page_index = i

                # Crops
                strip_crop = crop_rect(img_full, hit.info_strip_bbox)
                payee_crop = crop_rect(img_full, hit.payee_line_bbox)

                # OCR info strip
                info_text = ocr_text(strip_crop)
                date, chk, amt = parse_fields_from_info_text(info_text)

                # Fallback lower-left quadrant if needed
                if not any([date, chk, amt]):
                    llq = img_full[img_full.shape[0]//2:, :img_full.shape[1]//2]
                    d2, c2, a2 = parse_fields_from_info_text(ocr_text(llq))
                    date = date or d2
                    chk  = chk  or c2
                    amt  = amt  or a2

                # Payee OCR
                payee_text = ocr_text(payee_crop)
                payee = re.sub(r"(?i)pay\s*to\s*the\s*order\s*of[:\s]*", "", payee_text).strip()
                payee = re.sub(r"\s+\$\d[\d,\.]*$", "", payee).strip()

                rows.append(ExtractedCheck(
                    date=date or "",
                    check_number=chk or "",
                    payee=payee or "",
                    amount=(amt or "").replace("US$", "$")
                ))

                dt_page = time.perf_counter() - t_page
                logging.info(f"Page {i}: extracted fields in {dt_page:.2f}s: date='{date}' chk='{chk}' amt='{amt}' payee='{payee[:40]}'")

            except KeyboardInterrupt:
                logging.warning("Interrupted by user (Ctrl+C). Returning partial results.")
                return rows
            except Exception as e:
                logging.exception(f"ERROR on page {i}: {e}")
                continue

    return rows


# ----------------- CSV writer -----------------
def write_csv(rows: List[ExtractedCheck], csv_out: str):
    with open(csv_out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["date", "check_number", "payee", "", "amount"])
        for r in rows:
            w.writerow([r.date, r.check_number, r.payee, "", r.amount])

# ----------------- Tk flow -----------------
def tk_choose_files_and_vendor() -> Tuple[List[str], Optional[str], Optional[str]]:
    root = tk.Tk(); root.withdraw(); root.update()

    pdf_paths = filedialog.askopenfilenames(
        title="Select Bank of Hope PDF statement(s)",
        filetypes=[("PDF files","*.pdf")],
    )
    if not pdf_paths:
        messagebox.showerror("No PDF selected", "You must select at least one PDF.")
        root.destroy(); return [], None, None

    vendor_xlsx = filedialog.askopenfilename(
        title="Select Excel file of Vendors (optional)",
        filetypes=[("Excel files","*.xlsx *.xls")],
    )

    initial_dir = os.path.dirname(pdf_paths[0])
    save_csv = filedialog.asksaveasfilename(
        title="Save CSV as",
        defaultextension=".csv",
        initialdir=initial_dir,
        initialfile="checks_extracted.csv",
        filetypes=[("CSV","*.csv")]
    )
    if not save_csv:
        messagebox.showerror("No save path", "You must choose a CSV save path.")
        root.destroy(); return [], None, None

    root.destroy()
    return list(pdf_paths), (vendor_xlsx or None), save_csv

# ----------------- Main -----------------
def main():
    logging.info("Starting check extraction...")
    pdfs, vendor_excel, csv_out = tk_choose_files_and_vendor()
    if not pdfs or not csv_out:
        logging.error("No PDFs or CSV path selected. Exiting.")
        return

    # Vendors
    vendor_norms: Set[str] = set()
    if vendor_excel and os.path.exists(vendor_excel):
        try:
            vendor_norms = load_vendors_from_excel(vendor_excel)
            logging.info(f"Loaded {len(vendor_norms)} vendor names (normalized) from {os.path.basename(vendor_excel)}")
        except Exception as e:
            logging.exception(f"Failed to load vendors: {e}")

    # Process PDFs with progress bar
    all_rows: List[ExtractedCheck] = []
    for p in tqdm(pdfs, desc="PDFs", unit="pdf"):
        logging.info(f"Processing PDF: {p}")
        try:
            rows = process_pdf(p)
            all_rows.extend(rows)
            logging.info(f"PDF done: {os.path.basename(p)} -> {len(rows)} raw rows")
        except KeyboardInterrupt:
            logging.warning("Interrupted by user (Ctrl+C). Writing partial results...")
            break
        except Exception as e:
            logging.exception(f"ERROR processing {p}: {e}")

    # Filter by amount >= threshold (if you have AMOUNT_MIN in your script)
    filtered: List[ExtractedCheck] = []
    for r in all_rows:
        amt_float = amount_to_float(r.amount)
        if amt_float is not None and amt_float >= AMOUNT_MIN:
            filtered.append(r)

    # Standardize payees with vendor list
    if vendor_norms:
        std_rows: List[ExtractedCheck] = []
        for r in filtered:
            best, score = best_vendor_match(r.payee, vendor_norms)
            if score >= 90:
                r = ExtractedCheck(r.date, r.check_number, best, r.amount)
            std_rows.append(r)
        filtered = std_rows

    write_csv(filtered, csv_out)
    logging.info(f"All done. Wrote {len(filtered)} rows to {csv_out}")


if __name__ == "__main__":
    main()
