import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any
from tkinter.scrolledtext import ScrolledText
from datetime import date, datetime
import traceback
import sys

try:
    from vertex.utils.helpers import (
        ensure_relation_dict, display_relation_name,
        ensure_relation_link, merge_relations,
        migrate_officer_business_links_to_relations,
    )
except ModuleNotFoundError:
    from utils.helpers import (
        ensure_relation_dict, display_relation_name,
        ensure_relation_link, merge_relations,
        migrate_officer_business_links_to_relations,
    )


try:
    from vertex.config import US_STATES, ENTITY_TYPES
except ModuleNotFoundError:
    from config import US_STATES, ENTITY_TYPES

try:
    from vertex.ui.dialogs.linkdialog import LinkDialog
except ModuleNotFoundError:
    from ui.dialogs.linkdialog import LinkDialog

try:
    from vertex.ui.components.scrollframe import ScrollFrame
except ModuleNotFoundError:
    from ui.components.scrollframe import ScrollFrame

try:
    from vertex.ui.components.autocomplete import AutocompletePopup
except ModuleNotFoundError:
    from ui.components.autocomplete import AutocompletePopup


class ClientDialog(tk.Toplevel):
    def __init__(self, master, title: str, initial: Optional[Dict[str, Any]] = None):
        super().__init__(master)
        self.title(title); self.resizable(True, True); self.result=None
        self.minsize(880, 600)
        self.maxsize(880, 1200)
        self.debug_links = True

        self._initial = dict(initial or {})
        init = self._initial
        self.v_name    = tk.StringVar(value=init.get("name",""))
        self.v_dba     = tk.StringVar(value=init.get("dba",""))
        self.v_first = tk.StringVar(value=init.get("first_name",""))
        self.v_mid   = tk.StringVar(value=init.get("middle_name",""))
        self.v_last  = tk.StringVar(value=init.get("last_name",""))
        self.v_nick  = tk.StringVar(value=init.get("nickname",""))
        self.v_entity  = tk.StringVar(value=init.get("entity_type",""))
        self.v_is_individual = tk.BooleanVar(value=False)
        self.v_ein     = tk.StringVar(value=init.get("ein",""))
        self.v_edd     = tk.StringVar(value=init.get("edd_number",""))
        self.v_sales   = tk.StringVar(value=init.get("sales_tax_account",""))
        self.v_acctmgr = tk.StringVar(value=init.get("acct_mgr",""))
        self.v_path    = tk.StringVar(value=init.get("file_location",""))
        self.v_addr1 = tk.StringVar(value=init.get("addr1",""))
        self.v_addr2 = tk.StringVar(value=init.get("addr2",""))
        self.v_city  = tk.StringVar(value=init.get("city",""))
        self.v_state = tk.StringVar(value=init.get("state",""))
        self.v_zip   = tk.StringVar(value=init.get("zip",""))
        self.v_ui_rate   = tk.StringVar(value=init.get("ui_rate",""))
        self.v_stx_rate  = tk.StringVar(value=init.get("sales_tax_rate",""))
        self.v_other_tax = tk.StringVar(value=init.get("other_tax_rates",""))

        self.memo_init = init.get("memo","")

        sf = ScrollFrame(self, height=520); sf.pack(fill="both", expand=True)
        frm = sf.inner
        for i in range(4):
            frm.grid_columnconfigure(i, weight=1, uniform="client_cols")

        # --- Determine initial individual state ---
        et = (init.get("entity_type") or "").strip().lower()
        is_ind = (et == "individual" or bool(init.get("is_individual")) or bool(str(init.get("ssn", "")).strip()))
        self.v_is_individual.set(bool(is_ind))
        if is_ind:
            self.v_entity.set("Individual")
        self._last_business_entity = self.v_entity.get() if self.v_entity.get().strip().lower() != "individual" else ""

        # Row offset: row 0 is reserved for the checkbox
        R0 = 1

        # Top toggle row (ALWAYS at the top)
        chk_ind = ttk.Checkbutton(
            frm,
            text="Individual client",
            variable=self.v_is_individual,
            onvalue=True,
            offvalue=False
        )
        chk_ind.grid(row=0, column=0, sticky="w", pady=(4, 6), columnspan=4)

        def row(lbl, var, r, c=0, colspan=1):
            ttk.Label(frm, text=lbl).grid(row=r, column=c, sticky="w", pady=(0, 2), columnspan=colspan)
            ent = ttk.Entry(frm, textvariable=var)
            ent.grid(row=r+1, column=c, sticky="we", pady=(0, 6), columnspan=colspan)
            return ent

        # --- Business name rows (R0+0 / R0+1) ---
        lbl_name = ttk.Label(frm, text="Name")
        lbl_name.grid(row=R0+0, column=0, sticky="w", pady=(0,2))
        ent_name = ttk.Entry(frm, textvariable=self.v_name)
        ent_name.grid(row=R0+1, column=0, sticky="we", pady=(0,6), columnspan=2)

        lbl_dba = ttk.Label(frm, text="DBA")
        lbl_dba.grid(row=R0+0, column=2, sticky="w", pady=(0,2))
        ent_dba = ttk.Entry(frm, textvariable=self.v_dba)
        ent_dba.grid(row=R0+1, column=2, sticky="we", pady=(0,6), columnspan=2)

        # --- Individual name widgets (same space as Name/DBA; we toggle) ---
        lbl_first = ttk.Label(frm, text="First Name")
        ent_first = ttk.Entry(frm, textvariable=self.v_first)

        lbl_mid = ttk.Label(frm, text="Middle Name")
        ent_mid = ttk.Entry(frm, textvariable=self.v_mid)

        lbl_last = ttk.Label(frm, text="Last Name")
        ent_last = ttk.Entry(frm, textvariable=self.v_last)

        lbl_nick = ttk.Label(frm, text="Nickname")
        ent_nick = ttk.Entry(frm, textvariable=self.v_nick)

        # --- Entity type + EIN/EDD/Sales (R0+2 / R0+3) ---
        ttk.Label(frm, text="Entity Type").grid(row=R0+2, column=0, sticky="w", pady=(0,2))
        cb_entity = ttk.Combobox(frm, values=ENTITY_TYPES, textvariable=self.v_entity, state="readonly")
        cb_entity.grid(row=R0+3, column=0, sticky="we", pady=(0,6))

        lbl_ein = ttk.Label(frm, text="EIN (e.g., 12-3456789)")
        lbl_ein.grid(row=R0+2, column=1, sticky="w", pady=(0,2))
        ent_ein = ttk.Entry(frm, textvariable=self.v_ein)
        ent_ein.grid(row=R0+3, column=1, sticky="we", pady=(0,6))

        lbl_edd = ttk.Label(frm, text="EDD Number")
        lbl_edd.grid(row=R0+2, column=2, sticky="w", pady=(0,2))
        ent_edd = ttk.Entry(frm, textvariable=self.v_edd)
        ent_edd.grid(row=R0+3, column=2, sticky="we", pady=(0,6))

        lbl_sales = ttk.Label(frm, text="Sales Tax Account")
        lbl_sales.grid(row=R0+2, column=3, sticky="w", pady=(0,2))
        ent_sales = ttk.Entry(frm, textvariable=self.v_sales)
        ent_sales.grid(row=R0+3, column=3, sticky="we", pady=(0,6))

        r = R0 + 4  # start after Entity/EIN/EDD/Sales rows

        # Account Manager
        ttk.Label(frm, text="Account Manager").grid(row=r, column=0, sticky="w", pady=(0, 2), columnspan=4)
        r += 1
        mgr_cb, self.v_acctmgr = self.master._make_mgr_combobox(frm, initial_name=self.v_acctmgr.get())
        mgr_cb.grid(row=r, column=0, sticky="we", pady=(0, 6), columnspan=4)
        r += 1

        # File Location
        row("File Location (folder path or URL)", self.v_path, r=r, c=0, colspan=4)
        r += 2

        # Client Address header + lines
        ttk.Label(frm, text="Client Address").grid(row=r, column=0, sticky="w", pady=(8, 2), columnspan=4)
        r += 1
        row("Address 1", self.v_addr1, r=r, c=0, colspan=4)
        r += 2
        row("Address 2", self.v_addr2, r=r, c=0, colspan=4)
        r += 2

        # City / State / Zip
        ttk.Label(frm, text="City").grid(row=r, column=0, sticky="w", pady=(0, 2))
        ttk.Label(frm, text="State").grid(row=r, column=2, sticky="w", pady=(0, 2))
        ttk.Label(frm, text="Zip").grid(row=r, column=3, sticky="w", pady=(0, 2))
        r += 1

        ent_city = ttk.Entry(frm, textvariable=self.v_city)
        ent_city.grid(row=r, column=0, sticky="we", pady=(0, 6), columnspan=2)

        cb_state = ttk.Combobox(frm, values=US_STATES, textvariable=self.v_state, state="readonly")
        cb_state.grid(row=r, column=2, sticky="we", pady=(0, 6))

        ent_zip = ttk.Entry(frm, textvariable=self.v_zip)
        ent_zip.grid(row=r, column=3, sticky="we", pady=(0, 6))
        r += 1

        # Tax frame starts right after City/State/Zip
        tax_frame = ttk.Frame(frm)
        tax_frame.grid(row=r, column=0, columnspan=4, sticky="we")



        for i in range(4):
            tax_frame.grid_columnconfigure(i, weight=1)

        lbl_tax = ttk.Label(tax_frame, text="Tax Rates")
        lbl_tax.grid(row=0, column=0, sticky="w", pady=(8,2))

        def tax_row(lbl, var, r, c=0, colspan=1):
            ttk.Label(tax_frame, text=lbl).grid(row=r, column=c, sticky="w", pady=(0,2), columnspan=colspan)
            ent = ttk.Entry(tax_frame, textvariable=var)
            ent.grid(row=r+1, column=c, sticky="we", pady=(0,6), columnspan=colspan)
            return ent


        tax_row("UI Rate (%)", self.v_ui_rate, 1, 0)
        tax_row("Sales Tax Rate (%)", self.v_stx_rate, 1, 1)
        tax_row("Other Tax Rates (text)", self.v_other_tax, 1, 2, colspan=2)
        
        PERSONNEL_R = r+1

        lbl_personnel = ttk.Label(frm, text="Personnel / Business")
        lbl_personnel.grid(row=PERSONNEL_R, column=0, sticky="w", pady=(6,2), columnspan=4)
        self.off_cols = (
            "name","first_name","middle_name","last_name","nickname",
            "email","phone","addr1","addr2","city","state","zip","dob",
            "role","id",""
        )


        self.off_tree = ttk.Treeview(frm, columns=self.off_cols, show="headings", height=4, selectmode="browse")

        for col, label, w in (
            ("name","Entity",220),
            ("role","Role",110),
            ("email","Email",220),
            ("phone","Phone",120),
        ):
            self.off_tree.heading(col, text=label)
            self.off_tree.column(col, width=w, anchor="w")

        for col in ("first_name","middle_name","last_name","nickname","addr1","addr2","city","state","zip","dob",
                    "id"):
            self.off_tree.heading(col, text=col)
            self.off_tree.column(col, width=0, stretch=False)

        
        self.off_tree.grid(row=PERSONNEL_R+1, column=0, columnspan=4, sticky="nsew")
        off_btns = ttk.Frame(frm)
        off_btns.grid(row=PERSONNEL_R+2, column=0, sticky="w", pady=(4,6), columnspan=4)
        ttk.Button(off_btns, text="Link", command=self._rel_link).pack(side=tk.LEFT)
        ttk.Button(off_btns, text="Add", command=self._rel_add).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(off_btns, text="Edit", command=self._rel_edit).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(off_btns, text="Remove", command=self._rel_remove).pack(side=tk.LEFT, padx=(6,0))


        ttk.Label(frm, text="Memo").grid(row=PERSONNEL_R+4, column=0, sticky="w", pady=(4,2), columnspan=4)
        self.memo_txt = ScrolledText(frm, width=72, height=6, wrap="word")
        self.memo_txt.grid(row=PERSONNEL_R+5, column=0, columnspan=4, sticky="we", pady=(0,6))
        if self.memo_init: self.memo_txt.insert("1.0", self.memo_init)
        
        # --- Legacy migration (if some business links were stored in officers) ---
        try:
            migrate_officer_business_links_to_relations(init)
        except Exception:
            pass

        # --- Build combined rows for the tree: all relations (people + entity links) ---
        combined_rows = []

        # Migrate any remaining officers to relations display
        for o in (init.get("officers", []) or []):
            if isinstance(o, dict):
                combined_rows.append(ensure_relation_dict(o))

        # All relations (both people without links and entity links)
        for r in (init.get("relations", []) or []):
            rr = ensure_relation_link(r)
            if rr.get("id"):
                # Entity relation - link to another client
                # Include all data fields from the relation
                combined_rows.append(ensure_relation_dict({
                    "name": rr.get("name", "") or "",
                    "first_name": rr.get("first_name", "") or "",
                    "middle_name": rr.get("middle_name", "") or "",
                    "last_name": rr.get("last_name", "") or "",
                    "nickname": rr.get("nickname", "") or "",
                    "email": rr.get("email", "") or "",
                    "phone": rr.get("phone", "") or "",
                    "addr1": rr.get("addr1", "") or "",
                    "addr2": rr.get("addr2", "") or "",
                    "city": rr.get("city", "") or "",
                    "state": rr.get("state", "") or "",
                    "zip": rr.get("zip", "") or "",
                    "dob": rr.get("dob", "") or "",
                    "role": rr.get("role", "") or "business",
                    "id": rr.get("id", ""),
                }))
            else:
                # Person relation (migrated from officers) - show as-is
                combined_rows.append(ensure_relation_dict(r))

        for o in combined_rows:
            o = ensure_relation_dict(o)
            vals = (
                display_relation_name(o), o["first_name"], o["middle_name"], o["last_name"], o["nickname"],
                o["email"], o["phone"], o["addr1"], o["addr2"], o["city"], o["state"], o["zip"], o["dob"],
                o.get("role","officer"), o.get("id",""), "",  # No longer using linked_client_label
            )
            self.off_tree.insert("", "end", values=vals)

        btns = ttk.Frame(self); btns.pack(fill="x", pady=(6,6))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(btns, text="Save", command=self._save).pack(side=tk.RIGHT)

        self.bind("<Return>", lambda _e: self._save()); self.bind("<Escape>", lambda _e: self.destroy())

        def _focus_first():
            # focus correct field depending on entity
            if self.v_is_individual.get():
                ent_first.focus_set()
            else:
                ent_name.focus_set()

        self.after(50, _focus_first)
        self.grab_set()
        self.transient(master)

        def _is_individual() -> bool:
            return bool(self.v_is_individual.get())
        
        def _apply_entity_ui():
            ind = _is_individual()

            # Toggle name fields
            if ind:
                # hide business name/DBA
                lbl_name.grid_remove(); ent_name.grid_remove()
                lbl_dba.grid_remove();  ent_dba.grid_remove()

                # show individual fields
                lbl_first.grid(row=R0+0, column=0, sticky="w", pady=(0,2))
                ent_first.grid(row=R0+1, column=0, sticky="we", pady=(0,6))

                lbl_mid.grid(row=R0+0, column=1, sticky="w", pady=(0,2))
                ent_mid.grid(row=R0+1, column=1, sticky="we", pady=(0,6))

                lbl_last.grid(row=R0+0, column=2, sticky="w", pady=(0,2))
                ent_last.grid(row=R0+1, column=2, sticky="we", pady=(0,6))

                lbl_nick.grid(row=R0+0, column=3, sticky="w", pady=(0,2))
                ent_nick.grid(row=R0+1, column=3, sticky="we", pady=(0,6))


                # clear business-only fields when switching to Individual
                self.v_edd.set("")
                self.v_sales.set("")
                self.v_ui_rate.set("")
                self.v_stx_rate.set("")
                self.v_other_tax.set("")

                # force entity type
                self.v_entity.set("Individual")
                cb_entity.configure(state="disabled")

                # EIN -> SSN and hide EDD / Sales
                lbl_ein.configure(text="SSN (e.g., 123-45-6789)")
                lbl_edd.grid_remove(); ent_edd.grid_remove()
                lbl_sales.grid_remove(); ent_sales.grid_remove()

                # Hide tax rates
                tax_frame.grid_remove()

            else:
                # show business name/DBA
                lbl_name.grid(row=R0+0, column=0, sticky="w", pady=(0,2))
                ent_name.grid(row=R0+1, column=0, sticky="we", pady=(0,6), columnspan=2)

                lbl_dba.grid(row=R0+0, column=2, sticky="w", pady=(0,2))
                ent_dba.grid(row=R0+1, column=2, sticky="we", pady=(0,6), columnspan=2)

                # hide individual fields
                lbl_first.grid_remove(); ent_first.grid_remove()
                lbl_mid.grid_remove();   ent_mid.grid_remove()
                lbl_last.grid_remove();  ent_last.grid_remove()
                lbl_nick.grid_remove();  ent_nick.grid_remove()

                # enable entity type
                cb_entity.configure(state="readonly")

                # restore EIN/EDD/Sales
                lbl_ein.configure(text="EIN (e.g., 12-3456789)")
                lbl_edd.grid(); ent_edd.grid()
                lbl_sales.grid(); ent_sales.grid()

                # show tax rates
                tax_frame.grid()

        def _on_ind_toggle(*_):
            if self.v_is_individual.get():
                # remember current business entity before forcing Individual
                cur = (self.v_entity.get() or "").strip()
                if cur and cur.lower() != "individual":
                    self._last_business_entity = cur
                self.v_entity.set("Individual")
            else:
                # restore last business entity or a sane default
                restore = self._last_business_entity or ("Business" if "Business" in ENTITY_TYPES else (ENTITY_TYPES[0] if ENTITY_TYPES else ""))
                self.v_entity.set(restore)
            _apply_entity_ui()

        self.v_is_individual.trace_add("write", _on_ind_toggle)

        # Apply once on load
        _apply_entity_ui()
        if _is_individual():
            self.v_entity.set("Individual")
            cb_entity.configure(state="disabled")


        # If user changes entity type to Individual, apply UI immediately
        def _on_entity_change(_e=None):
            _apply_entity_ui()

        cb_entity.bind("<<ComboboxSelected>>", _on_entity_change)
        
    def _dlog(self, msg: str):
        # Prints to console (and to any captured stderr in your app log).
        # Safe if running in packaged exe too.
        try:
            print(f"[ClientDialog][LINK] {msg}", file=sys.stderr)
        except Exception:
            pass

    def _rel_selected_iid_or_none(self) -> Optional[str]:
        sel = self.off_tree.selection()
        if not sel:
            return None
        return sel[0]

    def _rel_values_to_dict(self, values: list[str]) -> Dict[str, str]:
        # off_cols is your schema; values is the tree row values
        d: Dict[str, str] = {}
        for i, k in enumerate(self.off_cols):
            d[k] = values[i] if i < len(values) else ""
        # normalize common fields
        d.setdefault("role", "officer")
        return ensure_relation_dict(d)

    def _rel_dict_to_values(self, o: Dict[str, str]) -> tuple:
        o = ensure_relation_dict(o)
        return (
            display_relation_name(o), o["first_name"], o["middle_name"], o["last_name"], o["nickname"],
            o["email"], o["phone"], o["addr1"], o["addr2"], o["city"], o["state"], o["zip"], o["dob"],
            o.get("role", "officer"), o.get("id", ""), "",  # Use "id" instead of "linked_client_id"
        )

    def _this_client_link_id(self) -> str:
        """
        Returns canonical link id for THIS client:
        - Individual -> ssn:<9>
        - Business   -> ein:<9>
        Empty string if missing required id.
        """
        try:
            ent = (self.v_entity.get() or "").strip().lower()
        except Exception:
            ent = (str(getattr(self, "v_entity", "") or "")).strip().lower()

        if ent == "individual":
            ssn9 = "".join(ch for ch in str(self.v_ein.get() or "") if ch.isdigit())[-9:]
            return f"ssn:{ssn9}" if ssn9 else ""
        else:
            ein9 = "".join(ch for ch in str(self.v_ein.get() or "") if ch.isdigit())[-9:]
            return f"ein:{ein9}" if ein9 else ""


    def _linked_ids_in_tree(self) -> set[str]:
        s = set()
        for iid in self.off_tree.get_children():
            v = list(self.off_tree.item(iid, "values") or [])
            while len(v) < 16:
                v.append("")
            lid = str(v[14] or "").strip()
            if lid:
                s.add(lid)
        return s

    def _resolve_client_for_linkdialog(self, client_id: str):
        """
        Returns a client dict for a given canonical id (ein:######### or ssn:#########),
        or None if not resolvable.
        """
        m = self.master
        client_id = str(client_id or "").strip()
        if not client_id:
            return None

        # Preferred: master has a resolver function
        for fn_name in ("get_client_by_id", "resolve_client", "find_client_by_id"):
            if hasattr(m, fn_name):
                try:
                    return getattr(m, fn_name)(client_id)
                except Exception:
                    return None

        # Fallback: common dict caches
        for attr in ("clients_by_id", "client_by_id", "_clients_by_id"):
            try:
                d = getattr(m, attr, None)
                if isinstance(d, dict) and client_id in d:
                    return d.get(client_id)
            except Exception:
                pass

        return None

    def _rel_add(self):
        cands = []
        try:
            cands = self.master.build_link_candidates(exclude_client_id=self._this_client_link_id())
        except Exception:
            try:
                cands = self.master.build_link_candidates()
            except Exception:
                cands = []

        dlg = LinkDialog(
            self,
            title="Add entity",
            initial={},
            link_candidates=cands,
            this_client_id=self._compute_required_client_id(),
            already_linked_ids=self._linked_ids_in_tree(),
            resolve_client=self._resolve_client_for_linkdialog,
        )
        self.wait_window(dlg)

        if not dlg.result:
            return

        r = dlg.result

        # NOTE: For a NEW client, we allow “temporary link” even if SSN/EIN isn't set yet.
        # The real validation happens in _save() when we attempt to persist/link.
        # For an EXISTING client, you can optionally enforce it here (not required).
        if (not self._is_new_client()):
            cid = self._require_this_client_id_or_error(action_label="link entities")
            if not cid:
                return

        values = (
            r.get("name", ""),
            r.get("first_name", ""), r.get("middle_name", ""), r.get("last_name", ""), r.get("nickname", ""),
            r.get("email", ""), r.get("phone", ""),
            r.get("addr1", ""), r.get("addr2", ""), r.get("city", ""), r.get("state", ""), r.get("zip", ""),
            r.get("dob", ""),
            (r.get("role", "") or "officer").strip().lower(),
            r.get("id", ""),  # Use "id" instead of "linked_client_id"
            "",  # No longer using linked_client_label
        )
        self.off_tree.insert("", "end", values=values)

    def _rel_link(self):
        print("[ClientDialog][LINK] _rel_link: Starting link process")
        cands = []
        try:
            cands = self.master.build_link_candidates(exclude_client_id=self._this_client_link_id())
        except Exception:
            try:
                cands = self.master.build_link_candidates()
            except Exception:
                cands = []

        if not cands:
            messagebox.showinfo("Link", "No entities available to link.")
            return

        dlg = LinkDialog(
            self,
            title="Add entity",
            initial={},
            link_candidates=cands,
            this_client_id=self._compute_required_client_id(),
            already_linked_ids=self._linked_ids_in_tree(),
            resolve_client=self._resolve_client_for_linkdialog,
        )
        self.wait_window(dlg)

        if not dlg.result:
            print("[ClientDialog][LINK] _rel_link: LinkDialog returned no result")
            return

        print(f"[ClientDialog][LINK] _rel_link: LinkDialog result: {dlg.result}")

        # MUST be linked in this flow
        linked_id = (dlg.result.get("id") or "").strip()
        if not linked_id:
            print(f"[ClientDialog][LINK] _rel_link: No linked_id in result: {dlg.result}")
            messagebox.showerror("Link", "Please select an entity from the suggestions to link.")
            return

        print(f"[ClientDialog][LINK] _rel_link: Linked ID: {linked_id}")

        # For an EXISTING client you may enforce having EIN/SSN *before* adding linked row
        # (optional), but for NEW client we allow temporary row.
        if (not self._is_new_client()):
            cid = self._require_this_client_id_or_error(action_label="link entities")
            if not cid:
                print(f"[ClientDialog][LINK] _rel_link: Missing client ID for existing client")
                return
            print(f"[ClientDialog][LINK] _rel_link: This client ID: {cid}")

        r = dlg.result

        values = (
            r.get("name", ""),
            r.get("first_name", ""), r.get("middle_name", ""), r.get("last_name", ""), r.get("nickname", ""),
            r.get("email", ""), r.get("phone", ""),
            r.get("addr1", ""), r.get("addr2", ""), r.get("city", ""), r.get("state", ""), r.get("zip", ""),
            r.get("dob", ""),
            (r.get("role", "") or "officer").strip().lower(),
            r.get("id", ""),  # Use "id" instead of "linked_client_id"
            "",  # No longer using linked_client_label
        )
        print(f"[ClientDialog][LINK] _rel_link: Inserting into tree with values: {values}")
        self.off_tree.insert("", "end", values=values)
        print(f"[ClientDialog][LINK] _rel_link: Successfully added to tree")

    def _rel_edit(self):
        sel = self.off_tree.selection()
        if not sel:
            messagebox.showinfo("Edit", "Select a personnel row to edit.")
            return

        iid = sel[0]
        v = list(self.off_tree.item(iid, "values") or [])
        while len(v) < 16:
            v.append("")

        o = {
            "name":        v[0],
            "first_name":  v[1], "middle_name": v[2], "last_name": v[3], "nickname": v[4],
            "email":       v[5], "phone": v[6],
            "addr1":       v[7], "addr2": v[8], "city": v[9], "state": v[10], "zip": v[11],
            "dob":         v[12],
            "role":        (v[13] or "").strip().lower(),
            "id": v[14] or "",  # Use "id" instead of "linked_client_id"
        }

        this_id = self._compute_required_client_id()
        link_cands = []
        try:
            if hasattr(self.master, "build_link_candidates"):
                link_cands = self.master.build_link_candidates(exclude_client_id=this_id)
        except Exception:
            link_cands = []

        already = []
        for x_iid in self.off_tree.get_children():
            if x_iid == iid:
                continue
            vv = list(self.off_tree.item(x_iid, "values") or [])
            while len(vv) < 16:
                vv.append("")
            lid = (vv[14] or "").strip()
            if lid:
                already.append(lid)

        dlg = LinkDialog(
            self,
            title="Edit Link",
            initial=o,
            link_candidates=link_cands,
            this_client_id=this_id,
            already_linked_ids=already,
            resolve_client=(getattr(self.master, "get_client_by_id", None) if hasattr(self.master, "get_client_by_id") else None),
        )
        self.wait_window(dlg)
        if not getattr(dlg, "result", None):
            return

        new_o = dict(dlg.result)
        self.off_tree.item(iid, values=self._rel_dict_to_values(new_o))


    def _rel_remove(self):
        print(f"[ClientDialog][LINK] _rel_remove: Starting removal process")
        sel = self.off_tree.selection()
        if not sel:
            messagebox.showinfo("Remove", "Select a personnel row to remove.")
            return

        v = list(self.off_tree.item(sel[0], "values") or [])
        while len(v) < 16:
            v.append("")

        print(f"[ClientDialog][LINK] _rel_remove: Tree values: {v}")
        
        # Get the relation ID if it's a linked entity
        relation_id = (v[14] or "").strip()  # id field is at index 14
        print(f"[ClientDialog][LINK] _rel_remove: Extracted relation_id='{relation_id}' from index 14")
        
        # Check if this is a new client
        is_new = self._is_new_client()
        print(f"[ClientDialog][LINK] _rel_remove: Is new client: {is_new}")
        
        # Remove from tree
        print(f"[ClientDialog][LINK] _rel_remove: Removing from tree")
        self.off_tree.delete(sel[0])
        print(f"[ClientDialog][LINK] _rel_remove: Removed from tree")
        
        # If this is a linked entity (has id), we need to bidirectionally unlink
        # For existing clients, we can handle it immediately
        # For new clients, it will be handled in _save() when it detects the relation was removed
        if relation_id:
            if not is_new:
                this_id = self._compute_required_client_id()
                print(f"[ClientDialog][LINK] _rel_remove: Existing client - this_id='{this_id}'")
                if this_id and hasattr(self.master, "link_clients_relations"):
                    try:
                        print(f"[ClientDialog][LINK] _rel_remove: Bidirectionally unlinking this_id='{this_id}' from relation_id='{relation_id}'")
                        self.master.link_clients_relations(this_id, relation_id, link=False)
                        print(f"[ClientDialog][LINK] _rel_remove: Successfully unlinked")
                    except Exception as e:
                        print(f"[ClientDialog][LINK] _rel_remove: Error unlinking: {e}")
                        import traceback
                        traceback.print_exc()
                else:
                    print(f"[ClientDialog][LINK] _rel_remove: Cannot unlink - this_id='{this_id}', has link_clients_relations={hasattr(self.master, 'link_clients_relations')}")
            else:
                print(f"[ClientDialog][LINK] _rel_remove: New client - unlinking will be handled in _save()")
        else:
            print(f"[ClientDialog][LINK] _rel_remove: No relation_id - this is a person relation (not linked entity)")

    def _digits_only(self, s: str) -> str:
        return "".join(ch for ch in (s or "") if ch.isdigit())

    def _compute_required_client_id(self) -> str:
        """
        Individual  -> ssn:<9digits>
        Business    -> ein:<9digits>
        Returns "" if missing.
        """
        ent = (self.v_entity.get() or "").strip().lower()
        raw = (self.v_ein.get() or "").strip()  # reused for SSN when Individual
        digits = self._digits_only(raw)

        # Keep ONLY the last 9 digits (matches your other code’s expectations)
        if len(digits) > 9:
            digits = digits[-9:]

        if ent == "individual":
            return f"ssn:{digits}" if digits else ""
        else:
            return f"ein:{digits}" if digits else ""


    def _require_this_client_id_or_error(self, action_label: str = "link clients") -> str:
        cid = self._compute_required_client_id()
        if cid:
            return cid

        is_ind = (self.v_entity.get() or "").strip().lower() == "individual"
        if is_ind:
            messagebox.showerror(
                "Missing SSN",
                f"Cannot {action_label}.\n\nThis client is an Individual but has no SSN.\nAdd SSN first, then link."
            )
        else:
            messagebox.showerror(
                "Missing EIN",
                f"Cannot {action_label}.\n\nThis client is a Business but has no EIN.\nAdd EIN first, then link."
            )
        return ""

    def _this_client_id(self) -> str:
        # Prefer explicit id if present
        cid = str(self._initial.get("id") or self._initial.get("client_id") or "").strip()
        if cid:
            return cid

        # Fallback: if your system uses idx:<n> but it's not stored, try to ask master/app
        # (safe no-op if attributes don't exist)
        for attr in ("current_client_id", "_current_client_id", "selected_client_id"):
            try:
                v = getattr(self.master, attr, "")
                v = str(v or "").strip()
                if v:
                    return v
            except Exception:
                pass

        return ""

    def _is_new_client(self) -> bool:
        # New client = no persisted id in initial
        init_id = str(self._initial.get("id") or self._initial.get("client_id") or "").strip()
        return not bool(init_id)
    
    def _get_current_client_idx(self) -> int | None:
        """
        Returns the index of the current client being edited, or None if new client.
        Used to ignore the current client when checking for duplicates.
        """
        if self._is_new_client():
            return None
        
        # Find the client index by matching the id
        init_id = str(self._initial.get("id") or "").strip()
        if not init_id:
            return None
        
        if not hasattr(self.master, "items"):
            return None
        
        for i, c in enumerate(self.master.items):
            if not isinstance(c, dict):
                continue
            c_id = str(c.get("id") or "").strip()
            if c_id == init_id:
                return i
        
        return None

    def _persist_client_if_possible(self, client_payload: Dict[str, Any]) -> bool:
        """
        Best-effort: try to make sure the new client exists in the master/manager BEFORE linking.
        Returns True if we believe the client is now persisted/registered, else False.
        """
        m = self.master

        # If App provides a "save/upsert/add" method, call it.
        # (We try several names to match your codebase without breaking.)
        candidates = [
            "upsert_client",
            "save_client",
            "add_client",
            "create_client",
            "insert_client",
            "put_client",
        ]

        for fn in candidates:
            if hasattr(m, fn):
                try:
                    getattr(m, fn)(client_payload)
                    return True
                except Exception:
                    # don't crash the dialog; linking will be deferred to caller
                    if getattr(self, "debug_links", False):
                        self._dlog(f"_persist_client_if_possible: {fn} failed:\n{traceback.format_exc()}")
                    return False

        # If we don't know how to persist here, we let the caller persist using self.result.
        return False

    def _apply_symmetric_links_now_if_possible(self, this_id: str, relations: list[Dict[str, str]]):
        """
        After this client exists in manager, apply symmetric entity links using relations.

        Prefers master.link_clients_relations if present, else falls back to master.link_clients.
        Each relation item should look like: {"id": "...", "role": "..."}.
        """
        print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: this_id='{this_id}', relations count={len(relations or [])}")
        this_id = str(this_id or "").strip()
        if not this_id:
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: No this_id, returning")
            return

        m = self.master
        link_fn = None

        if hasattr(m, "link_clients_relations"):
            link_fn = getattr(m, "link_clients_relations")
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Using link_clients_relations")
        elif hasattr(m, "link_clients"):
            link_fn = getattr(m, "link_clients")
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Using link_clients")
        else:
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: No link function found")

        if not link_fn:
            if getattr(self, "debug_links", False):
                self._dlog("_apply_symmetric_links_now_if_possible: master has no link function -> skip")
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: No link function, returning")
            return

        for i, rel in enumerate(relations or []):
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Processing relation {i}: {rel}")
            rr = ensure_relation_link(rel)
            link_id = str(rr.get("id") or "").strip()
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Relation {i} - link_id: '{link_id}'")
            if not link_id:
                print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Relation {i} - No link_id, skipping")
                continue

            role = (str(rr.get("role") or "") or "linked_client").strip().lower()
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Linking this='{this_id}' <-> other='{link_id}' role='{role}'")

            # keep compatibility with older signatures
            try:
                link_fn(this_id, link_id, link=True, role=role)
                print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Successfully called link function")
            except TypeError as e:
                print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: TypeError, trying without role: {e}")
                link_fn(this_id, link_id, link=True)
            except Exception as e:
                print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Exception during linking: {e}")
                import traceback
                traceback.print_exc()

        if hasattr(m, "save_clients_data"):
            print(f"[ClientDialog][LINK] _apply_symmetric_links_now_if_possible: Calling save_clients_data")
            m.save_clients_data()



    def _sync_client_link(self, other_client_id: str, link: bool, role: str = ""):
        other_client_id = str(other_client_id or "").strip()
        if not other_client_id:
            if getattr(self, "debug_links", False):
                self._dlog("sync_client_link: other_client_id empty -> return")
            return

        # Compute this side's canonical ID
        this_id = self._compute_required_client_id()

        if getattr(self, "debug_links", False):
            self._dlog(f"sync_client_link: START link={link} role='{role}'")
            self._dlog(f"  this computed id = '{this_id}'")
            self._dlog(f"  other_client_id  = '{other_client_id}'")
            try:
                self._dlog(f"  entity_type(v_entity) = '{self.v_entity.get()}'")
            except Exception:
                self._dlog("  entity_type(v_entity) = <error reading>")
            try:
                self._dlog(f"  v_ein(raw) = '{self.v_ein.get()}'")
            except Exception:
                self._dlog("  v_ein(raw) = <error reading>")

            # Helpful: show initial stored id too (if any)
            try:
                init_id = str(self._initial.get("id") or self._initial.get("client_id") or "").strip()
                self._dlog(f"  _initial.id = '{init_id}'")
            except Exception:
                self._dlog("  _initial.id = <error reading>")

            self._dlog(f"  master has link_clients = {hasattr(self.master, 'link_clients')}")
            self._dlog(f"  master type = {type(self.master)}")

        if not this_id:
            # This is the most common silent cause: missing EIN/SSN digits
            if getattr(self, "debug_links", False):
                self._dlog("sync_client_link: FAIL because this_id is empty (missing EIN/SSN digits)")
            raise ValueError("Current entity has no ID (EIN/SSN digits). Save it first.")

        if not (hasattr(self.master, "link_clients_relations") or hasattr(self.master, "link_clients")):
            if getattr(self, "debug_links", False):
                self._dlog("sync_client_link: FAIL because master.link_clients does not exist")
            raise ValueError("Client manager has no link_clients().")

        # Call manager with full exception logging
        try:
            if getattr(self, "debug_links", False):
                self._dlog("sync_client_link: calling master.link_clients(...)")

            if hasattr(self.master, "link_clients_relations"):
                self.master.link_clients_relations(this_id, other_client_id, link=link, role=role)
            else:
                self.master.link_clients(this_id, other_client_id, link=link, role=role)

            if getattr(self, "debug_links", False):
                self._dlog("sync_client_link: master.link_clients OK")

            if hasattr(self.master, "save_clients_data"):
                if getattr(self, "debug_links", False):
                    self._dlog("sync_client_link: calling master.save_clients_data()")
                self.master.save_clients_data()
                if getattr(self, "debug_links", False):
                    self._dlog("sync_client_link: master.save_clients_data OK")

        except Exception as ex:
            tb = traceback.format_exc()
            if getattr(self, "debug_links", False):
                self._dlog("sync_client_link: EXCEPTION from master.link_clients / save_clients_data")
                self._dlog(f"  ex={ex!r}")
                self._dlog(tb)

            # Re-raise with extra context so your messagebox includes IDs
            raise RuntimeError(
                "Linking failed.\n\n"
                f"this_id='{this_id}'\n"
                f"other_id='{other_client_id}'\n"
                f"link={link} role='{role}'\n\n"
                f"{ex}"
            ) from ex

    def _gather_rows(self) -> list[Dict[str, str]]:
        rows: list[Dict[str, str]] = []
        for iid in self.off_tree.get_children():
            v = list(self.off_tree.item(iid, "values") or [])
            while len(v) < 16:
                v.append("")

            # Build dict with id field
            row_dict = {
                "name":        v[0],
                "first_name":  v[1], "middle_name": v[2], "last_name": v[3], "nickname": v[4],
                "email":       v[5], "phone": v[6],
                "addr1":       v[7], "addr2": v[8], "city": v[9], "state": v[10], "zip": v[11],
                "dob":         v[12],
                "role":        (v[13] or "officer").strip().lower(),
                "id": (v[14] or "").strip(),  # Use "id" instead of "linked_client_id"
            }
            print(f"[ClientDialog][LINK] _gather_rows: Row {len(rows)} - dict before ensure_relation_dict: {row_dict}")
            # ensure_relation_dict now preserves id field
            normalized = ensure_relation_dict(row_dict)
            print(f"[ClientDialog][LINK] _gather_rows: Row {len(rows)} - dict after ensure_relation_dict: {normalized}")
            rows.append(normalized)
        return rows


    def _split_officers_and_relations(self, rows: list[Dict[str, str]]):
        """
        All rows with id => entity relation (stored in relations).
        All rows without id are also stored in relations (migrated from officers).
        Officers field is deprecated - everything goes to relations.
        """
        print(f"[ClientDialog][LINK] _split_officers_and_relations: Processing {len(rows or [])} rows")
        officers: list[Dict[str, str]] = []  # Keep empty for backward compatibility
        rels: list[Dict[str, str]] = []

        for i, o in enumerate(rows or []):
            print(f"[ClientDialog][LINK] _split_officers_and_relations: Row {i} BEFORE ensure_relation_dict: {o}")
            # Don't call ensure_relation_dict again - it was already called in _gather_rows
            # Just ensure it's a dict
            if not isinstance(o, dict):
                o = ensure_relation_dict(o)
            else:
                # Make sure id is preserved
                if "id" not in o and o.get("linked_client_id"):
                    o["id"] = o.get("linked_client_id", "")
            print(f"[ClientDialog][LINK] _split_officers_and_relations: Row {i} AFTER: {o}")
            # Check for "id" first, then fallback to "linked_client_id" for backward compatibility
            lid = str(o.get("id") or o.get("linked_client_id") or "").strip()
            role = (str(o.get("role") or "") or "").strip().lower()
            print(f"[ClientDialog][LINK] _split_officers_and_relations: Row {i} - lid: '{lid}', role: '{role}'")

            if lid:
                # Entity relation - link to another client
                # Ensure lid is in proper format (ein: or ssn:)
                if ":" not in lid:
                    # If not in proper format, try to normalize it
                    # This shouldn't happen if LinkDialog is working correctly, but handle it
                    print(f"[ClientDialog][LINK] _split_officers_and_relations: WARNING - lid '{lid}' doesn't have ':' format")
                    pass
                # Build full relation with all data fields
                rel_dict = {
                    "id": lid,
                    "name": (o.get("name") or "").strip(),
                    "first_name": (o.get("first_name") or "").strip(),
                    "middle_name": (o.get("middle_name") or "").strip(),
                    "last_name": (o.get("last_name") or "").strip(),
                    "nickname": (o.get("nickname") or "").strip(),
                    "email": (o.get("email") or "").strip(),
                    "phone": (o.get("phone") or "").strip(),
                    "addr1": (o.get("addr1") or "").strip(),
                    "addr2": (o.get("addr2") or "").strip(),
                    "city": (o.get("city") or "").strip(),
                    "state": (o.get("state") or "").strip(),
                    "zip": (o.get("zip") or "").strip(),
                    "dob": (o.get("dob") or "").strip(),
                    "role": role or "business",
                }
                print(f"[ClientDialog][LINK] _split_officers_and_relations: Created relation dict: {rel_dict}")
                rel_link = ensure_relation_link(rel_dict)
                print(f"[ClientDialog][LINK] _split_officers_and_relations: After ensure_relation_link: {rel_link}")
                rels.append(rel_link)
            else:
                # Person without client link - also store in relations (migrated from officers)
                # Store as a relation dict with all person details
                print(f"[ClientDialog][LINK] _split_officers_and_relations: Row {i} - No id, treating as person relation")
                rels.append(ensure_relation_dict(o))

        print(f"[ClientDialog][LINK] _split_officers_and_relations: Returning {len(rels)} relations")
        return officers, rels


    def _save(self):
        # Check for duplicate ID before saving
        # This allows the user to continue editing if a duplicate is found
        if hasattr(self.master, "_dup_id_conflict"):
            # Determine if this is a new client or editing existing
            is_new = self._is_new_client()
            ignore_idx = None if is_new else self._get_current_client_idx()
            
            # Build candidate dict for duplicate check
            this_id = self._compute_required_client_id()
            candidate = {
                "is_individual": self.v_entity.get().strip().lower() == "individual",
                "entity_type": "Individual" if self.v_entity.get().strip().lower() == "individual" else self.v_entity.get().strip(),
                "ein": self.v_ein.get().strip(),
                "ssn": self.v_ein.get().strip() if self.v_entity.get().strip().lower() == "individual" else "",
            }
            
            conflict = self.master._dup_id_conflict(candidate, ignore_idx=ignore_idx)
            if conflict:
                kind, digits, other_idx, other_client = conflict
                other_name = (other_client.get("name") or "").strip() or f"Client #{other_idx}"
                if is_new:
                    messagebox.showerror(
                        "Duplicate ID",
                        f"A client with the same {kind} already exists.\n\n"
                        f"{kind}: {digits}\n"
                        f"Existing client: {other_name}\n\n"
                        f"Please change the {kind} or cancel."
                    )
                else:
                    messagebox.showerror(
                        "Duplicate ID",
                        f"Another client already has this {kind}.\n\n"
                        f"{kind}: {digits}\n"
                        f"Existing client: {other_name}\n\n"
                        f"Please change the {kind} or cancel."
                    )
                return  # Don't save, keep dialog open
        
        if self.v_entity.get().strip().lower() == "individual":
            first = self.v_first.get().strip()
            mid   = self.v_mid.get().strip()
            last  = self.v_last.get().strip()
            nick  = self.v_nick.get().strip()
            
            self.v_edd.set("")
            self.v_sales.set("")
            self.v_ui_rate.set("")
            self.v_stx_rate.set("")
            self.v_other_tax.set("")

            if not (first or last or nick):
                messagebox.showerror("Required", "First/Last (or Nickname) is required for Individual clients.")
                return

            # build display name (reuse your helper if you want, but ClientDialog doesn't import it)
            parts = [first, mid, last]
            base = " ".join([p for p in parts if p])
            name = f'{base} ("{nick}")' if nick and base else (nick or base)

            # keep legacy fields aligned
            self.v_name.set(name)
            self.v_dba.set(nick)  # nickname replaces DBA concept
        else:
            name = self.v_name.get().strip()
            if not name:
                messagebox.showerror("Required", "Name is required.")
                return


        print(f"[ClientDialog][LINK] _save: About to gather rows from tree")
        rows = self._gather_rows()
        print(f"[ClientDialog][LINK] _save: Gathered {len(rows)} rows from tree")
        print(f"[ClientDialog][LINK] _save: Rows data: {rows}")
        officers, new_relations = self._split_officers_and_relations(rows)
        print(f"[ClientDialog][LINK] _save: After _split_officers_and_relations: {len(new_relations)} relations")
        print(f"[ClientDialog][LINK] _save: new_relations data: {new_relations}")
        memo     = self.memo_txt.get("1.0", "end").strip()


        this_id = self._compute_required_client_id()

        self.result = {
            "name": name,
            "dba": self.v_dba.get().strip(),
            "first_name": self.v_first.get().strip(),
            "middle_name": self.v_mid.get().strip(),
            "last_name": self.v_last.get().strip(),
            "nickname": self.v_nick.get().strip(),

            "entity_type": "Individual" if self.v_entity.get().strip().lower() == "individual" else self.v_entity.get().strip(),
            "ein": self.v_ein.get().strip(),

            "is_individual": self.v_entity.get().strip().lower() == "individual",
            "ssn": self.v_ein.get().strip() if self.v_entity.get().strip().lower() == "individual" else "",

            "edd_number": self.v_edd.get().strip(),
            "sales_tax_account": self.v_sales.get().strip(),
            "acct_mgr": self.v_acctmgr.get().strip(),
            "file_location": self.v_path.get().strip(),

            "addr1": self.v_addr1.get().strip(),
            "addr2": self.v_addr2.get().strip(),
            "city":  self.v_city.get().strip(),
            "state": self.v_state.get().strip(),
            "zip":   self.v_zip.get().strip(),

            "ui_rate": self.v_ui_rate.get().strip(),
            "sales_tax_rate": self.v_stx_rate.get().strip(),
            "other_tax_rates": self.v_other_tax.get().strip(),
            "tax_rates_last_checked": "" if self.v_entity.get().strip().lower() == "individual" else date.today().isoformat(),

            "officers": [],  # Deprecated - all data moved to relations
            "employees": self._initial.get("employees", []),

            # IMPORTANT: id is the canonical resolver id used for linking
            "id": this_id,
        }
        
        # Use ONLY the relations from the tree (new_relations), not from _initial
        # This ensures that removed relations are actually removed
        print(f"[ClientDialog][LINK] _save: Processing relations")
        print(f"[ClientDialog][LINK] _save: Initial relations: {len(self._initial.get('relations', []) or [])}")
        print(f"[ClientDialog][LINK] _save: New relations from tree: {len(new_relations)}")
        print(f"[ClientDialog][LINK] _save: New relations data: {new_relations}")
        
        # IMPORTANT: Use only new_relations from the tree, not merge with _initial
        # This ensures that when a relation is removed from the tree, it's actually removed
        # The tree is the source of truth for what relations should exist
        merged_rels = new_relations
        print(f"[ClientDialog][LINK] _save: Using relations from tree (not merged): {len(merged_rels)}")
        
        # Normalize to ensure id format
        normalized_rels = []
        for r in merged_rels:
            r_link = ensure_relation_link(r)
            print(f"[ClientDialog][LINK] _save: Processing relation: {r_link}")
            if r_link.get("id"):
                # Entity link - use link format
                print(f"[ClientDialog][LINK] _save: Found entity link with id: {r_link.get('id')}")
                normalized_rels.append(r_link)
            else:
                # Person relation - use dict format
                print(f"[ClientDialog][LINK] _save: Found person relation (no id)")
                normalized_rels.append(ensure_relation_dict(r))
        
        print(f"[ClientDialog][LINK] _save: Final normalized relations: {len(normalized_rels)}")
        print(f"[ClientDialog][LINK] _save: Normalized relations data: {normalized_rels}")
        self.result["relations"] = normalized_rels

        # --- Process bidirectional links ---
        try:
            if self._is_new_client():
                # For new clients, store post_save_links to process after client is saved
                self.result["post_save_links"] = [
                    {
                        "this_id": this_id,
                        "other_id": str(r.get("id") or "").strip(),
                        "role": (str(r.get("role") or "") or "linked_client").strip().lower(),
                    }
                    for r in normalized_rels
                    if this_id and str(r.get("id") or "").strip()
                ]
                if getattr(self, "debug_links", False):
                    self._dlog("New client: stored post_save_links (relations) for caller.")
            else:
                # For existing clients, handle both additions and removals
                # Get previous relations to find what was removed
                prev_relations = self._initial.get("relations", []) or []
                prev_other_ids = set()
                for pr in prev_relations:
                    pr_link = ensure_relation_link(pr)
                    link_id = pr_link.get("id")
                    if link_id:
                        prev_other_ids.add(link_id)
                
                # Find relations to unlink (in prev but not in new)
                current_other_ids = set()
                for r in normalized_rels:
                    r_link = ensure_relation_link(r)
                    link_id = r_link.get("id")
                    if link_id:
                        current_other_ids.add(link_id)
                
                print(f"[ClientDialog][LINK] _save: Previous relation IDs: {prev_other_ids}")
                print(f"[ClientDialog][LINK] _save: Current relation IDs: {current_other_ids}")
                removed_ids = prev_other_ids - current_other_ids
                print(f"[ClientDialog][LINK] _save: Relations to unlink: {removed_ids}")
                
                # Unlink removed relations - bidirectionally remove from both sides
                for removed_id in removed_ids:
                    if hasattr(self.master, "link_clients_relations"):
                        try:
                            print(f"[ClientDialog][LINK] _save: Unlinking this_id='{this_id}' from removed_id='{removed_id}'")
                            # This will remove B from A's relations and A from B's relations
                            self.master.link_clients_relations(this_id, removed_id, link=False)
                            print(f"[ClientDialog][LINK] _save: Successfully unlinked")
                        except Exception as e:
                            print(f"[ClientDialog][LINK] _save: Error unlinking: {e}")
                            import traceback
                            traceback.print_exc()
                    else:
                        print(f"[ClientDialog][LINK] _save: master has no link_clients_relations method")
                
                # Link new relations (only those that are actually new, not already linked)
                # Get currently linked IDs from app.items to avoid re-linking
                currently_linked_ids = set()
                if hasattr(self.master, "items") and this_id:
                    from vertex.utils.helpers import find_client_by_uid
                    try:
                        current_client = find_client_by_uid(self.master.items, this_id)
                        if isinstance(current_client, dict):
                            for rel in current_client.get("relations", []) or []:
                                rel_link = ensure_relation_link(rel)
                                rel_id = rel_link.get("id")
                                if rel_id:
                                    currently_linked_ids.add(rel_id)
                            print(f"[ClientDialog][LINK] _save: Currently linked IDs from app.items: {currently_linked_ids}")
                    except Exception as e:
                        print(f"[ClientDialog][LINK] _save: Error getting currently linked IDs: {e}")
                
                # Only link relations that aren't already linked
                new_relations_to_link = [r for r in normalized_rels if ensure_relation_link(r).get("id") not in currently_linked_ids]
                print(f"[ClientDialog][LINK] _save: About to call _apply_symmetric_links_now_if_possible with {len(new_relations_to_link)} new relations (out of {len(normalized_rels)} total)")
                if new_relations_to_link:
                    self._apply_symmetric_links_now_if_possible(this_id, new_relations_to_link)
                else:
                    print(f"[ClientDialog][LINK] _save: No new relations to link (all already linked)")
                print(f"[ClientDialog][LINK] _save: After _apply_symmetric_links_now_if_possible, result relations: {self.result.get('relations', [])}")
                
                # IMPORTANT: Use normalized_rels (from tree) as the source of truth for this client's relations.
                # The tree represents what the user wants, and link_clients_relations has already updated the other side.
                # Ensure self.result["relations"] is set to normalized_rels (it should already be, but make sure)
                self.result["relations"] = normalized_rels
                print(f"[ClientDialog][LINK] _save: Final result relations count: {len(normalized_rels)}")
                print(f"[ClientDialog][LINK] _save: Final result relations: {normalized_rels}")



        except Exception:
            if getattr(self, "debug_links", False):
                self._dlog("Save: exception in post-save linking stage:\n" + traceback.format_exc())
            # Don't block saving the dialog; caller can still proceed

        self.destroy()
