from __future__ import annotations

import sys, pathlib
if __package__ in (None, ""):
    _ROOT = pathlib.Path(__file__).resolve().parents[1]  # <Scripts/> root
    if str(_ROOT) not in sys.path:
        sys.path.insert(0, str(_ROOT))
        
import os, re, glob, shutil, datetime, sys, subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from functions.utils.app_logging import get_logger

ALLOWED_EXTS = {".pdf", ".jpg", ".jpeg", ".png"}
LOG = get_logger("documents")

DOC_TYPES_BY_ENTITY = {
    "S-Corporation": [
        "Articles of Incorporation", "Statement of Information", "EIN Confirmation Letter",
        "EFTPS Enrollments", "EFTPS PIN Letter", "EDD Registration", "Seller's Permit",
        "Board of Equalization Registration", "Fictitious Business Name Statement",
        "Stock Certificate", "Shareholder/Board Minutes", "Corporate Bylaws", "Insurance", "Misc",
    ],
    "Corporation (C-Corp)": [
        "Articles of Incorporation", "Statement of Information", "EIN Confirmation Letter",
        "EFTPS Enrollments", "EFTPS PIN Letter", "EDD Registration", "Seller's Permit",
        "Board of Equalization Registration", "Fictitious Business Name Statement",
        "Stock Certificate", "Shareholder/Board Minutes", "Corporate Bylaws", "Insurance", "Misc",
    ],
    "LLC": [
        "Articles of Organization", "Statement of Information", "Operating Agreement",
        "EIN Confirmation Letter", "EFTPS Enrollments", "EFTPS PIN Letter", "EDD Registration",
        "Seller's Permit", "Board of Equalization Registration", "Fictitious Business Name Statement",
        "Member Certificate", "Insurance", "Misc",
    ],
    "Individual / Sole Proprietor": [
        "Business License", "EIN Confirmation Letter", "Seller's Permit",
        "Board of Equalization Registration", "Fictitious Business Name Statement",
        "EFTPS Enrollments", "EFTPS PIN Letter", "EDD Registration", "Insurance", "Misc",
    ],
}

DOC_CLUSTERS_BY_ENTITY = {
    "S-Corporation": {
        "Corporate / Entity": [
            "Articles of Incorporation","Statement of Information","Stock Certificate",
            "Corporate Bylaws","Shareholder/Board Minutes","Fictitious Business Name Statement",
        ],
        "Taxes (Federal/State Payroll & EFTPS)": [
            "EIN Confirmation Letter","EFTPS Enrollments","EFTPS PIN Letter","EDD Registration",
        ],
        "Sales Tax / Permits (CDTFA/BOE)": [
            "Seller's Permit","Board of Equalization Registration",
        ],
    },
    "Corporation (C-Corp)": {
        "Corporate / Entity": [
            "Articles of Incorporation","Statement of Information","Stock Certificate",
            "Corporate Bylaws","Shareholder/Board Minutes","Fictitious Business Name Statement",
        ],
        "Taxes (Federal/State Payroll & EFTPS)": [
            "EIN Confirmation Letter","EFTPS Enrollments","EFTPS PIN Letter","EDD Registration",
        ],
        "Sales Tax / Permits (CDTFA/BOE)": [
            "Seller's Permit","Board of Equalization Registration",
        ],
    },
    "LLC": {
        "Corporate / Entity": [
            "Articles of Organization","Operating Agreement","Statement of Information",
            "Member Certificate","Fictitious Business Name Statement",
        ],
        "Taxes (Federal/State Payroll & EFTPS)": [
            "EIN Confirmation Letter","EFTPS Enrollments","EFTPS PIN Letter","EDD Registration",
        ],
        "Sales Tax / Permits (CDTFA/BOE)": [
            "Seller's Permit","Board of Equalization Registration",
        ],
    },
    "Individual / Sole Proprietor": {
        "Business / Entity": [
            "Business License","Fictitious Business Name Statement",
        ],
        "Taxes (Federal/State Payroll & EFTPS)": [
            "EIN Confirmation Letter","EFTPS Enrollments","EFTPS PIN Letter","EDD Registration",
        ],
        "Sales Tax / Permits (CDTFA/BOE)": [
            "Seller's Permit","Board of Equalization Registration",
        ],
    },
}

RENEWAL_REQUIRED = {
    "Statement of Information",
    "Business License",
    "Seller's Permit",
    "Fictitious Business Name Statement",
    "Insurance",
}

def _now_ts() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

def _quarter_label(q: str) -> str:
    q = (q or "").strip().upper()
    return q if q in ("1Q","2Q","3Q","4Q") else ""

def _is_allowed(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in ALLOWED_EXTS

def _clients_root() -> str:
    """
    Try to anchor at the same folder that holds clients.json / company_list.json.
    We now prefer .../functions/data/clients, but keep fallbacks.
    """
    guesses = [
        Path(__file__).resolve().parent.parent,           # functions/
        Path(__file__).resolve().parent.parent.parent,    # project root
        Path.home() / "Documents" / "LineUpDocs"
    ]
    for g in guesses:
        data_dir    = g / "data" / "clients"
        legacy_dir  = g / "clients"

        # Prefer new data/clients structure
        for base in (data_dir, legacy_dir):
            for fname in ("clients.json", "company_list.json"):
                p = base / fname
                if p.exists():
                    return str(base)
            if base.exists():
                return str(base)

    return str(Path.home() / "Documents" / "LineUpDocs" / "clients")


def _company_dir(c: dict) -> str:
    nm = (c.get("name") or "Client").strip()
    ein = (c.get("ein") or "").strip().replace("-", "")
    safe = re.sub(r"[^\w\-\.\s]", "_", nm).strip().replace(" ", "_")
    leaf = f"{safe}_{ein}" if ein else safe
    base = os.path.join(_clients_root(), leaf)
    os.makedirs(base, exist_ok=True)
    return base

def _year_dir(c: dict, year: int) -> str:
    base = _company_dir(c)
    ydir = os.path.join(base, f"{year:04d}")
    os.makedirs(ydir, exist_ok=True)
    return ydir

def _open_path(path: str):
    LOG.info("Open path: %s", path)
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)

def get_types_and_clusters_for_entity(entity_label: str):
    et = (entity_label or "").strip()
    types = DOC_TYPES_BY_ENTITY.get(et)
    clusters = DOC_CLUSTERS_BY_ENTITY.get(et)
    if not types or not clusters:
        types = DOC_TYPES_BY_ENTITY["LLC"]
        clusters = DOC_CLUSTERS_BY_ENTITY["LLC"]
    return types, clusters

def _cluster_for(doc_type: str, clusters_map: dict[str, list[str]]) -> str:
    for k, v in (clusters_map or {}).items():
        if doc_type in v:
            return k
    return "Other"

def _category_subcluster_tags_for(doc_type: str, clusters_map: dict[str, list[str]]):
    dt = (doc_type or "").strip().lower()
    is_eftps = "eftps" in dt
    is_edd   = "edd"   in dt
    is_ein   = "ein"   in dt or "irs" in dt

    if is_eftps or is_edd or is_ein:
        category = "Taxes"
        tags = ["Taxes"]
        if is_eftps:
            sub = "EFTPS"; tags.append("EFTPS")
        elif is_edd:
            sub = "EDD"; tags.append("EDD")
        else:
            sub = "IRS/EIN"; tags.append("IRS/EIN")
        return category, sub, tags

    primary = _cluster_for(doc_type, clusters_map)
    category = primary
    sub = ""
    tags = [primary]
    return category, sub, tags

def _next_index_for(c: dict, year: int, q: str, doc_type: str) -> int:
    qtag = f"_{q}" if q else ""
    pattern = os.path.join(_year_dir(c, year), f"{year:04d}{qtag}_{doc_type}*")
    existing = glob.glob(pattern)
    if not existing:
        return 1
    mx = 1
    for p in existing:
        base = os.path.splitext(os.path.basename(p))[0]
        m = re.search(r"_(\d+)$", base)
        if m:
            mx = max(mx, int(m.group(1)))
        else:
            mx = max(mx, 1)
    return mx + 1

def _target_path(c: dict, year: int, q: str, doc_type: str, ext: str, index: int) -> str:
    qtag = f"_{q}" if q else ""
    suffix = "" if index == 1 else f"_{index}"
    fname = f"{year:04d}{qtag}_{doc_type}{suffix}{ext.lower()}"
    return os.path.join(_year_dir(c, year), fname)

def _most_recent_for(c: dict, doc_type: str, y: int|None=None, q: str|None=None):
    ordering = {"1Q":1, "2Q":2, "3Q":3, "4Q":4}
    docs = [d for d in (c.get("documents") or [])
            if (d.get("type")==doc_type)
            and (y is None or d.get("year")==y)
            and (q in (None,"") or (d.get("quarter") or "")==q)]
    def _key(d):
        yr = int(d.get("year") or 0)
        qv = ordering.get((d.get("quarter") or "").upper(), 0)
        mo = int(d.get("month") or 0)
        tstr = d.get("ts") or "1970-01-01 00:00"
        try:
            dt = datetime.datetime.strptime(tstr, "%Y-%m-%d %H:%M")
        except Exception:
            dt = datetime.datetime(1970,1,1)
        return (yr, qv, mo, dt)
    if not docs:
        return None
    return sorted(docs, key=_key, reverse=True)[0]

def _most_recent_for_at_or_before(c: dict, doc_type: str, y: int|None=None, q: str|None=None):
    """
    Return the most recent document of given type with (year,quarter) <= (y,q),
    where quarter ordering is 1Q<2Q<3Q<4Q. If y/q is None, behaves like _most_recent_for.
    """
    if not y and not q:
        return _most_recent_for(c, doc_type, None, None)

    ordering = {"1Q":1, "2Q":2, "3Q":3, "4Q":4}
    qsel = (q or "").upper()
    qselv = ordering.get(qsel, 4)  # if q empty, treat as end-of-year

    cand = []
    for d in (c.get("documents") or []):
        if d.get("type") != doc_type:
            continue
        dy = int(d.get("year") or 0)
        dq = (d.get("quarter") or "").upper()
        dqv = ordering.get(dq, 0)
        if y is None:
            cand.append(d)
        else:
            if dy < y or (dy == y and dqv <= qselv):
                cand.append(d)

    if not cand:
        return None

    def _key(d):
        yr = int(d.get("year") or 0)
        qv = ordering.get((d.get("quarter") or "").upper(), 0)
        mo = int(d.get("month") or 0)
        tstr = d.get("ts") or "1970-01-01 00:00"
        try:
            dt = datetime.datetime.strptime(tstr, "%Y-%m-%d %H:%M")
        except Exception:
            dt = datetime.datetime(1970,1,1)
        return (yr, qv, mo, dt)

    return sorted(cand, key=_key, reverse=True)[0]

def init_documents_tab(nb: ttk.Notebook, self_ref, c: dict, save_clients_cb):
    docs_tab = ttk.Frame(nb, padding=8)
    nb.add(docs_tab, text="Documents")

    # Tab filter
    _entity = (c.get("entity") or c.get("entity_type") or "").strip()
    TYPES, CLUSTERS = get_types_and_clusters_for_entity(_entity)
    DEFAULT_TYPES = DOC_TYPES_BY_ENTITY.get(_entity, DOC_TYPES_BY_ENTITY["LLC"])
    enabled_types = {t: tk.BooleanVar(master=docs_tab, value=True) for t in DEFAULT_TYPES}

    host = ttk.Frame(docs_tab)
    host.grid(row=0, column=0, sticky="nsew")
    docs_tab.grid_rowconfigure(0, weight=1)
    docs_tab.grid_columnconfigure(0, weight=1)

    canvas = tk.Canvas(host, highlightthickness=0)
    vbar = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vbar.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    vbar.grid(row=0, column=1, sticky="ns")
    host.grid_rowconfigure(0, weight=1)
    host.grid_columnconfigure(0, weight=1)

    def _resize(_evt=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfigure(body_id, width=canvas.winfo_width())

    def _mw(event):
        if not canvas.winfo_exists():
            return
        try:
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units"); return
            elif event.num == 5:
                canvas.yview_scroll(1, "units"); return
        
            delta = int(getattr(event, "delta", 0))
            if delta:
                steps = -1 if delta > 0 else 1  # normalize to ±1 step
                canvas.yview_scroll(steps, "units")
        except tk.TclError:
            return

    def _stop_wheel(e):
        return "break"

    def _unbind_wheel(_evt=None):
        try:
            docs_tab.unbind_all("<MouseWheel>")
            docs_tab.unbind_all("<Button-4>")
            docs_tab.unbind_all("<Button-5>")
        except Exception:
            pass

    def _attach_wheel(widget):
        def _on_enter(_):
            widget.bind_all("<MouseWheel>", _mw)
            widget.bind_all("<Button-4>", _mw)   # Linux up
            widget.bind_all("<Button-5>", _mw)   # Linux down
        def _on_leave(_):
            try:
                widget.unbind_all("<MouseWheel>")
                widget.unbind_all("<Button-4>")
                widget.unbind_all("<Button-5>")
            except Exception:
                pass
        widget.bind("<Enter>", _on_enter)
        widget.bind("<Leave>", _on_leave)


    def _no_page_scroll_for(widget):
        # Windows/macOS
        widget.bind("<MouseWheel>", _stop_wheel)
        widget.bind("<Shift-MouseWheel>", _stop_wheel)
        # Linux
        widget.bind("<Button-4>", _stop_wheel)
        widget.bind("<Button-5>", _stop_wheel)


    body = ttk.Frame(canvas)  # everything goes into "body"
    body_id = canvas.create_window((0, 0), window=body, anchor="nw")
    body.bind("<Configure>", _resize)
    canvas.after(0, lambda: (docs_tab.update_idletasks(), _resize(), canvas.yview_moveto(0)))
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(body_id, width=e.width))
    docs_tab.bind("<Map>", lambda e: _refresh_scroll())

    _attach_wheel(canvas)

    def _refresh_scroll():
        docs_tab.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.yview_moveto(0)

    c.setdefault("documents", [])

    # entity-derived config
    _entity = (c.get("entity") or c.get("entity_type") or "").strip()
    TYPES, CLUSTERS = get_types_and_clusters_for_entity(_entity)
    DEFAULT_TYPES = DOC_TYPES_BY_ENTITY.get(_entity, DOC_TYPES_BY_ENTITY["LLC"])

    # ---- top bar ----
    top = ttk.Frame(body)
    top.grid(row=0, column=0, sticky="we", pady=(0,8))
    body.grid_columnconfigure(0, weight=1)

    top.grid_columnconfigure(9, weight=1)
    toolbar = ttk.Frame(top)
    toolbar.grid(row=0, column=10, sticky="e")

    hide_empty_var = tk.BooleanVar(value=True)

    def _apply_on_change(*_):
        _rebuild_boxes()

    hide_empty_var.trace_add("write", _apply_on_change)
    for v in enabled_types.values():
        v.trace_add("write", _apply_on_change)

    ttk.Label(top, text=f"Entity: {_entity or 'Unknown'}").grid(row=0, column=0, padx=(0,12), sticky="w")

    year_var    = tk.IntVar(value=int(_now_ts()[:4]))
    quarter_var = tk.StringVar(value="")
    def _topbar_changed(*_):
        _rebuild_boxes()

    year_var.trace_add("write", _topbar_changed)
    quarter_var.trace_add("write", _topbar_changed)


    ttk.Label(top, text="Year").grid(row=0, column=1, sticky="e")
    year_box = ttk.Spinbox(top, from_=2000, to=2100, width=6, textvariable=year_var)
    year_box.grid(row=0, column=2, padx=(4,12), sticky="w")

    ttk.Label(top, text="Quarter").grid(row=0, column=3, sticky="e")
    q_box = ttk.Combobox(top, width=4, state="readonly",
                         values=["", "1Q","2Q","3Q","4Q"], textvariable=quarter_var)
    q_box.grid(row=0, column=4, sticky="w")

    # ---- Group by
    group_mode = tk.StringVar(value="Type")
    ttk.Label(top, text="Group by").grid(row=0, column=7, sticky="e")
    group_box = ttk.Combobox(
        top, width=14, state="readonly",
        values=["Type", "Tax subcluster"], textvariable=group_mode
    )
    group_box.grid(row=0, column=8, padx=(4,0), sticky="w")

    # ---- boxes area
    boxes = ttk.Frame(body)
    boxes.grid(row=1, column=0, sticky="we")
    for i in range(3):
        boxes.grid_columnconfigure(i, weight=1)

    # ---- filters    
    # put this alongside your other toolbar buttons
    def _open_filters_panel():
        # keep a single window alive
        if getattr(_open_filters_panel, "_win", None) and _open_filters_panel._win.winfo_exists():
            _open_filters_panel._win.deiconify()
            _open_filters_panel._win.lift()
            return
    
        win = tk.Toplevel(self_ref.winfo_toplevel())
        _open_filters_panel._win = win
        win.title("Filters")
        win.transient(self_ref)          # stays on top of this window, but not modal
        # win.attributes("-topmost", True)  # uncomment if you want it always on top
        win.geometry("+{}+{}".format(self_ref.winfo_rootx()+80, self_ref.winfo_rooty()+80))
    
        frm = ttk.Frame(win, padding=10)
        frm.grid(sticky="nsew")
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)
    
        # Header buttons
        hdr = ttk.Frame(frm)
        hdr.grid(row=0, column=0, sticky="we", pady=(0,8))
        hdr.grid_columnconfigure(1, weight=1)
    
        def _select_all():
            for v in enabled_types.values(): v.set(True)
        def _select_none():
            for v in enabled_types.values(): v.set(False)
    
        ttk.Label(hdr, text="Show document types:").grid(row=0, column=0, sticky="w")
        ctls = ttk.Frame(hdr); ctls.grid(row=0, column=2, sticky="e")
        ttk.Button(ctls, text="Select all",  command=_select_all).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(ctls, text="Select none", command=_select_none).pack(side=tk.LEFT)
    
        # Hide empty toggle (auto-applies because of the trace)
        ttk.Checkbutton(frm, text="Hide types with no documents", variable=hide_empty_var)\
            .grid(row=1, column=0, sticky="w", pady=(0,8))
    
        # Scrollable checklist
        wrap = ttk.Frame(frm); wrap.grid(row=2, column=0, sticky="nsew")
        frm.grid_rowconfigure(2, weight=1); frm.grid_columnconfigure(0, weight=1)
    
        canvas2 = tk.Canvas(wrap, highlightthickness=0)
        vs2 = ttk.Scrollbar(wrap, orient="vertical", command=canvas2.yview)
        canvas2.configure(yscrollcommand=vs2.set)
        canvas2.pack(side=tk.LEFT, fill="both", expand=True)
        vs2.pack(side=tk.RIGHT, fill="y")
        _attach_wheel(canvas2)
    
        inner = ttk.Frame(canvas2)
        win_id = canvas2.create_window((0,0), window=inner, anchor="nw")
    
        def _resize_filters(_=None):
            canvas2.configure(scrollregion=canvas2.bbox("all"))
            canvas2.itemconfigure(win_id, width=canvas2.winfo_width())
        inner.bind("<Configure>", _resize_filters)
        canvas2.bind("<Configure>", lambda e: canvas2.itemconfigure(win_id, width=e.width))
    
        # lay out checks
        cols = 3
        for i, t in enumerate(DEFAULT_TYPES):
            ttk.Checkbutton(inner, text=t, variable=enabled_types[t]).grid(
                row=i//cols, column=i%cols, sticky="w", padx=8, pady=4
            )
        for ccol in range(cols):
            inner.grid_columnconfigure(ccol, weight=1)
    
        # gentle close (hide instead of destroy)
        def _on_close():
            win.withdraw()
        win.protocol("WM_DELETE_WINDOW", _on_close)
    
    # toolbar button
    ttk.Button(toolbar, text="Filters…", command=_open_filters_panel).pack(side=tk.LEFT, padx=(0,6))

    # ---- helpers that close over current tab state
    def _save():
        # mirror original call: save_clients(self.items)
        return save_clients_cb(self_ref.items)

    def _get_current_doc(c, doc_type: str):
         y = int(year_var.get()) if str(year_var.get()).isdigit() and int(year_var.get()) > 0 else None
         q = _quarter_label(quarter_var.get())
         if doc_type in RENEWAL_REQUIRED:
             return _most_recent_for_at_or_before(c, doc_type, y, q)

         return _most_recent_for(c, doc_type, None, None)

    def _rebuild_boxes():
        for child in boxes.winfo_children():
            child.destroy()

        mode = group_mode.get()
        if mode == "Tax subcluster":
            buckets = {"EFTPS": [], "EDD": [], "IRS/EIN": [], "Other Taxes": [], "Non-Tax": []}
            for t in DEFAULT_TYPES:
                category, sub, _tags = _category_subcluster_tags_for(t, CLUSTERS)
                if category == "Taxes":
                    if sub in ("EFTPS","EDD","IRS/EIN"):
                        buckets[sub].append(t)
                    else:
                        buckets["Other Taxes"].append(t)
                else:
                    buckets["Non-Tax"].append(t)
            grouped_types = [(k, v) for k, v in (
                ("EFTPS", buckets["EFTPS"]),
                ("EDD", buckets["EDD"]),
                ("IRS/EIN", buckets["IRS/EIN"]),
                ("Other Taxes", buckets["Other Taxes"]),
                ("Non-Tax", buckets["Non-Tax"]),
            ) if v]
        else:
            grouped_types = [("Types", DEFAULT_TYPES)]

        ncol = 3
        row_index = 0
        for title, types in grouped_types:
            hdr = ttk.Label(boxes, text=title, font=("", 10, "bold"))
            hdr.grid(row=row_index, column=0, columnspan=ncol, sticky="w", padx=6, pady=(6,0))
            row_index += 1

            for j, doc_type in enumerate(types):
                if not enabled_types.get(doc_type, tk.BooleanVar(value=True)).get():
                    continue
                if hide_empty_var.get():
                    any_doc = any((d.get("type") == doc_type) for d in (c.get("documents") or []))
                    if not any_doc:
                        continue

                r = row_index + (j // ncol)
                ccol = j % ncol

                card = ttk.Labelframe(boxes, text=doc_type)
                card.grid(row=r, column=ccol, padx=6, pady=6, sticky="we")
                card.grid_columnconfigure(1, weight=1)
                yrv = tk.IntVar(value=0)
                qrv = tk.StringVar(value="")

                created_controls = False
                ysel = None
                qsel = None

                name_row = 1 if created_controls else 0
                btn_row  = 2
                btns = ttk.Frame(card)
                btns.grid(row=btn_row, column = 0, columnspan = 4, sticky="w", pady=(6,0))

                name_lbl = ttk.Label(card, text="(no document)")
                name_lbl.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6,0))

                def _update_label(_evt=None, _dt=doc_type, _yrv=yrv, _qrv=qrv, _lbl=name_lbl):                
                    d = _get_current_doc(c, _dt)
                    if d:
                        meta = f"{d.get('year','')}"
                        if d.get('quarter'):
                            meta += f" ({d.get('quarter')})"
                        if d.get('month'):
                            meta += f", M{d.get('month')}"
                        _lbl.configure(text=f"{d.get('name','')} — {meta}")
                    else:
                        _lbl.configure(text="(no document)")

                if qsel is not None:
                    qsel.bind("<<ComboboxSelected>>", _update_label)
                if ysel is not None:
                    ysel.bind("<FocusOut>", _update_label)
                _update_label()

                def _view_current(_dt=doc_type, _yrv=yrv, _qrv=qrv):
                    d = _get_current_doc(c, _dt)
                    if d and d.get("path"):
                        _open_path(d["path"])
                                
                def _delete_current(_dt=doc_type, _yrv=yrv, _qrv=qrv):
                    d = _get_current_doc(c, _dt)
                    if not d:
                        messagebox.showinfo("Delete Document", "No document matched the current filters.")
                        return
                    choice = messagebox.askyesnocancel(
                        "Delete",
                        "Delete file from disk too?\n\nYes = delete from disk\nNo  = remove from list only\nCancel = do nothing"
                    )
                    if choice is None:
                        return
                    if choice:
                        try:
                            if d.get("path") and os.path.exists(d["path"]):
                                os.remove(d["path"])
                        except Exception as ex:
                            messagebox.showerror("Delete failed", str(ex))
                            return

                    docs = c.get("documents") or []
                    c["documents"] = [x for x in docs if not (x.get("name")==d.get("name") and str(x.get("ts"))==str(d.get("ts")))]
                    _save()
                    _update_label()        # refresh the card label
                    _rebuild_boxes()       # rebuild grid to reflect removal
                    refresh_csv_list()     # keep CSV in sync
                    _refresh_scroll()
                
                def _edit_current_type(_dt=doc_type, _yrv=yrv, _qrv=qrv):
                    d = _get_current_doc(c, _dt)
                    if not d:
                        messagebox.showinfo("Edit Type", "No document matched the current filters.")
                        return
                
                    win = tk.Toplevel(self_ref.winfo_toplevel()); win.title("Edit Document Type"); win.transient(self_ref)
                    frm = ttk.Frame(win, padding=10); frm.grid(sticky="nsew")
                    win.grid_columnconfigure(0, weight=1); win.grid_rowconfigure(0, weight=1)
                
                    default_types = DOC_TYPES_BY_ENTITY.get(_entity, DOC_TYPES_BY_ENTITY["LLC"])
                    ttk.Label(frm, text="Document Type").grid(row=0, column=0, sticky="w")
                    v = tk.StringVar(value=d.get("type","Misc"))
                    cb = ttk.Combobox(frm, textvariable=v, values=default_types, state="readonly")
                    cb.grid(row=1, column=0, sticky="we"); frm.grid_columnconfigure(0, weight=1)
                
                    def _ok():
                        dtyp = v.get().strip() or "Misc"
                        d["type"] = dtyp
                        # refresh tags/category/cluster:
                        category, subcluster, tags = _category_subcluster_tags_for(dtyp, CLUSTERS)
                        d["cluster"]    = _cluster_for(dtyp, CLUSTERS)
                        d["category"]   = category
                        d["subcluster"] = subcluster
                        d["tags"]       = tags
                        _save()
                        _update_label()
                        refresh_csv_list()
                        _refresh_scroll()
                        win.destroy()
                
                    btns2 = ttk.Frame(frm); btns2.grid(row=2, column=0, sticky="e", pady=(10,0))
                    ttk.Button(btns2, text="Cancel", command=win.destroy).grid(row=0, column=0, padx=(0,6))
                    ttk.Button(btns2, text="Save",   command=_ok).grid(row=0, column=1)
                    win.grab_set(); cb.focus_set()
                
                ttk.Button(btns, text="View",      command=_view_current).pack(side=tk.LEFT, padx=(0,6))
                ttk.Button(btns, text="Delete",    command=_delete_current).pack(side=tk.LEFT, padx=(0,6))
                ttk.Button(btns, text="Edit Type", command=_edit_current_type).pack(side=tk.LEFT)

            row_index += (len(types) + 2) // 3

    # first build cards
    _rebuild_boxes()
    _refresh_scroll()
    body.update_idletasks()
    _resize()
    canvas.yview_moveto(0)
    group_box.bind("<<ComboboxSelected>>", lambda e: _rebuild_boxes())


    # ---- CSV/Excel-style list (name, document type, year, quarter)
    csv_area = ttk.Frame(body)
    csv_area.grid(row=3, column=0, sticky="we", pady=(6, 0))
    body.grid_rowconfigure(3, weight=1)
    ttk.Label(csv_area, text="List").pack(anchor="w")
    csv_wrap = ttk.Frame(csv_area)
    csv_wrap.pack(fill="both", expand=True)

    # Treeview with 4 columns
    csv_cols = ("name", "type", "year", "quarter")
    csv_tree = ttk.Treeview(csv_wrap, columns=csv_cols, show="headings", selectmode="browse")
    for cid, label, w in (
        ("name","name",420),
        ("type","document type",220),
        ("year","year",80),
        ("quarter","quarter",90),
    ):
        csv_tree.heading(cid, text=label)
        csv_tree.column(cid, width=w, anchor="w")

    csv_scroll_y = ttk.Scrollbar(csv_wrap, orient="vertical", command=csv_tree.yview)
    csv_tree.configure(yscrollcommand=csv_scroll_y.set)
    csv_tree.pack(side=tk.LEFT, fill="both", expand=True)
    csv_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

    # Buttons
    csv_btns = ttk.Frame(csv_area)
    csv_btns.pack(anchor="e", pady=(6, 0))

    def _build_csv_lines() -> list[str]:
        lines = []
        # header (Excel/CSV friendly)
        lines.append("name,document type,year,quarter")
        for d in (c.get("documents") or []):
            name = (d.get("name") or "").replace(",", " ")
            dtype = (d.get("type") or "").replace(",", " ")
            year = str(d.get("year") or "")
            qtr = (d.get("quarter") or "").strip() or "-"
            lines.append(f"{name},{dtype},{year},{qtr}")
        return lines

    def refresh_csv_list():
        # rebuild rows in the csv_tree
        csv_tree.delete(*csv_tree.get_children())
        docs = (c.get("documents") or [])
        for d in docs:
            name = (d.get("name") or "").replace(",", " ")
            dtype = (d.get("type") or "").replace(",", " ")
            year = str(d.get("year") or "")
            qtr  = (d.get("quarter") or "").strip() or "-"
            csv_tree.insert("", "end", values=(name, dtype, year, qtr))

    def _copy_csv_to_clipboard():
        try:
            # reuse your existing _build_csv_lines() so it stays Excel-friendly
            txt = "\n".join(_build_csv_lines())
            csv_area.clipboard_clear()
            csv_area.clipboard_append(txt)
        except Exception:
            pass
    
    def _export_csv_file():
        import csv
        from tkinter import filedialog
        lines = _build_csv_lines()
        if not lines:
            return
        path = filedialog.asksaveasfilename(
            title="Save CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            for row in lines:
                w.writerow(row.split(","))

    ttk.Button(csv_btns, text="Copy CSV",   command=_copy_csv_to_clipboard).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(csv_btns, text="Export CSV…", command=_export_csv_file).pack(side=tk.LEFT)

    def _find_doc_by_row(values):
        # values = (name, type, year, quarter)
        n, t, y, q = values
        y_int = None
        try:
            y_int = int(y) if y else None
        except Exception:
            y_int = None
        # match by (name, type, year, quarter); if multiple, prefer most recent
        candidates = []
        for d in (c.get("documents") or []):
            if (d.get("name")==n and
                (d.get("type") or "") == t and
                (str(d.get("year") or "") == (y or "")) and
                ((d.get("quarter") or "").strip() or "-") == q):
                candidates.append(d)
        if not candidates:
            # fallback: match by name only if exact match not found
            candidates = [d for d in (c.get("documents") or []) if d.get("name")==n]
        if not candidates:
            return None
        # prefer most recent by timestamp if available
        def _key(d):
            ts = str(d.get("ts") or "1970-01-01 00:00")
            try:
                return datetime.datetime.strptime(ts, "%Y-%m-%d %H:%M")
            except Exception:
                return datetime.datetime(1970,1,1)
        return sorted(candidates, key=_key, reverse=True)[0]
    
    def _on_csv_select(_=None):
        sel = csv_tree.selection()
        if not sel:
            return
        vals = csv_tree.item(sel[0], "values")
        d = _find_doc_by_row(vals)
        # (optional) you can scroll/focus the matching card type if you want
        # For now we just keep it selected.
    
    def _on_csv_open(_=None):
        sel = csv_tree.selection()
        if not sel:
            return
        vals = csv_tree.item(sel[0], "values")
        d = _find_doc_by_row(vals)
        if d and d.get("path") and os.path.exists(d["path"]):
            _open_path(d["path"])
    
    csv_tree.bind("<<TreeviewSelect>>", _on_csv_select)
    csv_tree.bind("<Double-1>", _on_csv_open)

    docs_tab.update_idletasks()
    _resize()
    canvas.yview_moveto(0)
    refresh_csv_list()

    _entity_label = _entity  # capture for defaults

    def upload_doc():
        paths = filedialog.askopenfilenames(
            title="Select PDF/JPG/PNG",
            filetypes=[("Allowed Files", "*.pdf;*.jpg;*.jpeg;*.png"),
                       ("PDF", "*.pdf"), ("Images", "*.jpg;*.jpeg;*.png")]
        )
        if not paths: return

        win = tk.Toplevel(self_ref.winfo_toplevel()); win.title("Upload Document"); win.transient(self_ref)
        frm = ttk.Frame(win, padding=10); frm.grid(sticky="nsew")
        win.grid_columnconfigure(0, weight=1); win.grid_rowconfigure(0, weight=1)

        default_types = DOC_TYPES_BY_ENTITY.get(_entity_label, DOC_TYPES_BY_ENTITY["LLC"])
        ttk.Label(frm, text="Document Type").grid(row=0, column=0, sticky="w")
        type_var = tk.StringVar(value=default_types[0] if default_types else "Misc")
        cb = ttk.Combobox(frm, textvariable=type_var, values=default_types, state="readonly")
        cb.grid(row=1, column=0, sticky="we", pady=(2,8)); frm.grid_columnconfigure(0, weight=1)

        ttk.Label(frm, text="Year").grid(row=2, column=0, sticky="w")
        yv = tk.IntVar(value=year_var.get() or 2025)
        ys = ttk.Spinbox(frm, from_=2000, to=2100, width=6, textvariable=yv); ys.grid(row=3, column=0, sticky="w")

        ttk.Label(frm, text="Quarter").grid(row=4, column=0, sticky="w", pady=(8,0))
        qv = tk.StringVar(value=_quarter_label(quarter_var.get()))
        qs = ttk.Combobox(frm, width=4, state="readonly", values=["","1Q","2Q","3Q","4Q"], textvariable=qv)
        qs.grid(row=5, column=0, sticky="w")

        ttk.Label(frm, text="Month (optional)").grid(row=6, column=0, sticky="w", pady=(8,0))
        mv = tk.StringVar(value=str(""))
        ms = ttk.Combobox(frm, width=4, state="readonly",
                          values=["","1","2","3","4","5","6","7","8","9","10","11","12"], textvariable=mv)
        ms.grid(row=7, column=0, sticky="w")

        def _cancel(): win.destroy()

        def _do_upload():
            dtyp  = type_var.get().strip() or "Misc"
            year  = int(yv.get())
            q     = _quarter_label(qv.get())
            msel  = mv.get().strip()
            month = int(msel) if msel.isdigit() else 0

            added_any = False
            for src in paths:
                if not _is_allowed(src):
                    messagebox.showwarning("Not allowed", f"Skipped (not PDF/JPG/PNG):\n{src}")
                    continue
                ext = os.path.splitext(src)[1].lower()
                next_idx = _next_index_for(c, year, q, dtyp)
                tgt = _target_path(c, year, q, dtyp, ext, next_idx)

                if os.path.exists(tgt):
                    choice = messagebox.askyesnocancel(
                        "Duplicate",
                        f"File exists:\n{os.path.basename(tgt)}\n\nOverwrite? (Yes)\nAdd as new? (No)\nCancel to stop."
                    )
                    if choice is None:
                        continue
                    if choice is False:
                        k = next_idx + 1
                        while os.path.exists(_target_path(c, year, q, dtyp, ext, k)):
                            k += 1
                        tgt = _target_path(c, year, q, dtyp, ext, k)

                try:
                    shutil.copy2(src, tgt)
                    category, subcluster, tags = _category_subcluster_tags_for(dtyp, CLUSTERS)
                    c["documents"].append({
                        "name": os.path.basename(tgt),
                        "type": dtyp,
                        "cluster": _cluster_for(dtyp, CLUSTERS),
                        "category": category,
                        "subcluster": subcluster,
                        "tags": tags,
                        "ts": _now_ts(),
                        "year": year,
                        "quarter": q,
                        "month": month or "",
                        "path": tgt
                    })
                    added_any = True
                except Exception as ex:
                    messagebox.showerror("Copy failed", f"Could not add:\n{src}\n\n{ex}")
            if added_any:
                _save()
                refresh_csv_list()
                _rebuild_boxes()
                _refresh_scroll()
            win.destroy()

        btns = ttk.Frame(frm); btns.grid(row=8, column=0, sticky="e", pady=(10,0))
        ttk.Button(btns, text="Cancel", command=_cancel).grid(row=0, column=0, padx=(0,6))
        ttk.Button(btns, text="Upload", command=_do_upload).grid(row=0, column=1)
        win.grab_set(); cb.focus_set()



    # visibility state
    _cards_visible = True
    _list_visible  = True

    cards_toggle_text = tk.StringVar(value="Hide Cards")

    def _toggle_hide_empty():
        hide_empty_var.set(not hide_empty_var.get())
        _rebuild_boxes()
    
    def _toggle_cards():
        nonlocal _cards_visible
        if _cards_visible:
            boxes.grid_remove()
            _cards_visible = False
            cards_toggle_text.set("Show Cards")
        else:
            boxes.grid()
            _cards_visible = True
            cards_toggle_text.set("Hide Cards")
        _refresh_scroll()
    
    def _toggle_list():
        nonlocal _list_visible
        if _list_visible:
            csv_area.grid_remove()
            _list_visible = False
        else:
            csv_area.grid()
            _list_visible = True
        _refresh_scroll()

    ttk.Button(toolbar, text="Upload",    command=upload_doc).pack(side=tk.LEFT, padx=(0,6))
    ttk.Button(toolbar, textvariable=cards_toggle_text, command=_toggle_cards).pack(side=tk.LEFT, padx=(0,6))
    ttk.Button(toolbar, text="Toggle List",   command=_toggle_list).pack(side=tk.LEFT, padx=(0,6))

    return docs_tab
