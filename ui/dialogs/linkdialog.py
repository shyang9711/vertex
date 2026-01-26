from __future__ import annotations

import sys, pathlib, types, re
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List, Callable

# Make imports work in BOTH layouts:
_BASE = pathlib.Path(__file__).resolve().parent
if "functions" not in sys.modules:
    _pkg = types.ModuleType("functions")
    _pkg.__path__ = [str(_BASE)]
    sys.modules["functions"] = _pkg
if str(_BASE) not in sys.path:
    sys.path.insert(0, str(_BASE))
_PARENT = _BASE.parent
if str(_PARENT) not in sys.path:
    sys.path.insert(0, str(_PARENT))

# AutocompletePopup import (try a few common locations)
try:
    from vertex.ui.components.autocomplete import AutocompletePopup
except Exception:
    from ui.components.autocomplete import AutocompletePopup
PHONE_DIGITS_RE = re.compile(r"\d")

def _digits9(x: str) -> str:
    ds = "".join(PHONE_DIGITS_RE.findall(str(x or "")))
    return ds[:9]  # first 9 digits (EIN/SSN)

def _norm_link_id(link_id: str) -> str:
    link_id = (link_id or "").strip()
    if not link_id:
        return ""
    if ":" not in link_id:
        return ""

    kind, val = link_id.split(":", 1)
    kind = (kind or "").strip().lower()
    val = (val or "").strip()

    # allow legacy formats to round-trip (so editing old saved links doesn't erase them)
    if kind in ("client", "idx"):
        return f"{kind}:{val}" if val else ""

    # canonical formats
    if kind in ("ein", "ssn"):
        val9 = _digits9(val)
        return f"{kind}:{val9}" if val9 else ""

    return ""


def _is_valid_link_id(link_id: str) -> bool:
    return bool(_norm_link_id(link_id))


def _matches_tokens(label: str, tokens: List[str]) -> bool:
    low = (label or "").lower()
    return all(t in low for t in tokens)

class LinkDialog(tk.Toplevel):
    """
    One dialog that supports BOTH:
    - Add (manual entry): allow user typing. Link is optional.
    - Link (search entity): selecting a candidate overwrites fields and locks editing.

    Returns officer/person dict with these keys:
      name, first_name, middle_name, last_name, nickname,
      email, phone, addr1, addr2, city, state, zip, dob,
      role, linked_client_id, linked_client_label
    """

    def __init__(
        self,
        master,
        title: str = "Link / Add Person",
        initial: Optional[Dict[str, str]] = None,
        link_candidates: Optional[list] = None,
        this_client_id: str = "",
        already_linked_ids: Optional[set[str]] = None,
        resolve_client: Optional[Callable[[str], Optional[Dict[str, Any]]]] = None,
    ):

        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master.winfo_toplevel() if hasattr(master, "winfo_toplevel") else master)
        self.grab_set()

        self.result: Optional[Dict[str, str]] = None

        self._initial = dict(initial or {})
        self._cands = list(link_candidates or [])
        self._this_client_id = _norm_link_id(this_client_id)
        self._already_linked_ids = { _norm_link_id(x) for x in (already_linked_ids or set()) if _norm_link_id(x) }
        self._resolve_client = resolve_client

        # candidate maps
        self._labels: List[str] = []
        self._label_to_id: Dict[str, str] = {}
        self._label_to_is_company: Dict[str, bool] = {}
        print(f"[LinkDialog] Initializing with {len(self._cands)} candidates")
        for i, c in enumerate(self._cands):
            if not isinstance(c, dict):
                print(f"[LinkDialog] Candidate {i} is not a dict: {type(c)}")
                continue
            lab = str(c.get("label") or "").strip()
            if not lab:
                print(f"[LinkDialog] Candidate {i} missing label")
                continue
            cid = _norm_link_id(str(c.get("id") or "").strip())
            # Allow candidates without IDs to be searchable (they just can't be linked)
            # Only require label for searchability
            self._labels.append(lab)
            if cid:
                # Only add to mapping if it has a valid ID (can be linked)
                self._label_to_id[lab] = cid
                self._label_to_is_company[lab] = bool(c.get("is_company", False))
            else:
                # Candidate without ID - can be searched but not linked
                print(f"[LinkDialog] Candidate {i} has no valid ID (label='{lab}') - searchable but not linkable")
        print(f"[LinkDialog] Built {len(self._labels)} searchable labels")
        print(f"[LinkDialog] Labels: {self._labels[:10]}...")  # Show first 10
        print(f"[LinkDialog] Linkable labels (with ID): {len(self._label_to_id)}")

        # ---------------- UI vars ----------------
        self.v_link = tk.StringVar(value=str(self._initial.get("linked_client_label") or "").strip())
        self.v_role = tk.StringVar(value=str(self._initial.get("role") or "").strip().lower())

        self.v_name  = tk.StringVar(value=str(self._initial.get("name") or ""))
        self.v_first = tk.StringVar(value=str(self._initial.get("first_name") or ""))
        self.v_mid   = tk.StringVar(value=str(self._initial.get("middle_name") or ""))
        self.v_last  = tk.StringVar(value=str(self._initial.get("last_name") or ""))
        self.v_nick  = tk.StringVar(value=str(self._initial.get("nickname") or ""))

        self.v_email = tk.StringVar(value=str(self._initial.get("email") or ""))
        self.v_phone = tk.StringVar(value=str(self._initial.get("phone") or ""))

        self.v_addr1 = tk.StringVar(value=str(self._initial.get("addr1") or ""))
        self.v_addr2 = tk.StringVar(value=str(self._initial.get("addr2") or ""))
        self.v_city  = tk.StringVar(value=str(self._initial.get("city") or ""))
        self.v_state = tk.StringVar(value=str(self._initial.get("state") or ""))
        self.v_zip   = tk.StringVar(value=str(self._initial.get("zip") or ""))
        self.v_dob   = tk.StringVar(value=str(self._initial.get("dob") or ""))

        self._linked_client_id = _norm_link_id(str(self._initial.get("linked_client_id") or ""))
        self._linked_client_label = str(self._initial.get("linked_client_label") or "").strip()
        self._linked_is_company = False
        
        # Determine if this client is individual or business
        self._this_is_individual = False
        if self._this_client_id:
            # Check from ID format first (ssn: = individual, ein: = business)
            if self._this_client_id.startswith("ssn:"):
                self._this_is_individual = True
            elif self._this_client_id.startswith("ein:"):
                self._this_is_individual = False
            else:
                # Try to resolve this client to determine type
                if callable(self._resolve_client):
                    try:
                        this_client = self._resolve_client(self._this_client_id)
                        if isinstance(this_client, dict):
                            self._this_is_individual = bool(this_client.get("is_individual")) or ((this_client.get("entity_type") or "").strip().casefold() == "individual")
                    except Exception:
                        pass

        # ---------------- layout ----------------
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        # Link row
        row = ttk.Frame(outer)
        row.pack(fill="x")
        ttk.Label(row, text="Link (search entity):").pack(side="left")
        self.ent_link = ttk.Entry(row, textvariable=self.v_link, width=52)
        self.ent_link.pack(side="left", padx=(8, 0), fill="x", expand=True)

        # Autocomplete popup
        self._popup = AutocompletePopup(self, self.ent_link, on_choose=self._on_choose_label)

        # IMPORTANT: do NOT show suggestions on FocusIn (fixes “always on / can’t type”)
        self.ent_link.bind("<KeyRelease>", self._on_link_key)
        self.ent_link.bind("<Key>", self._on_link_key_press)  # Trigger on key press to enable search
        self.ent_link.bind("<Down>", self._on_link_down)
        self.ent_link.bind("<Up>", self._on_link_up)
        self.ent_link.bind("<Return>", self._on_link_return)
        self.ent_link.bind("<Escape>", self._on_link_escape)
        self.ent_link.bind("<FocusOut>", self._on_link_focus_out)

        # Role row - will be updated when a client is linked
        row2 = ttk.Frame(outer)
        row2.pack(fill="x", pady=(10, 0))
        ttk.Label(row2, text="Role:").pack(side="left")
        self.cmb_role = ttk.Combobox(row2, textvariable=self.v_role, width=18, state="readonly")
        # Default values - will be updated based on client types
        self.cmb_role["values"] = ("spouse", "parent", "child", "relative", "employee", "officer", "business owner", "business")
        self.cmb_role.pack(side="left", padx=(8, 0))
        
        # Update role options initially
        self._update_role_options()

        # Fields grid
        grid = ttk.Frame(outer)
        grid.pack(fill="both", expand=True, pady=(10, 0))

        # columns: label + value, repeated
        for c in range(6):
            grid.grid_columnconfigure(c, weight=1)

        def add_pair(r, c_label, text, var, w=18, padx=(0, 8)):
            ttk.Label(grid, text=text).grid(row=r, column=c_label, sticky="e", padx=padx, pady=2)
            e = ttk.Entry(grid, textvariable=var, width=w)
            e.grid(row=r, column=c_label + 1, sticky="w", pady=2)
            return e

        # Row 0: Name (wide)
        ttk.Label(grid, text="Name:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=2)
        self.ent_name = ttk.Entry(grid, textvariable=self.v_name, width=50)
        self.ent_name.grid(row=0, column=1, columnspan=5, sticky="w", pady=2)

        # Row 1: First / Middle / Last
        self.ent_first = add_pair(1, 0, "First:",  self.v_first, w=18)
        self.ent_mid   = add_pair(1, 2, "Middle:", self.v_mid,   w=10)
        self.ent_last  = add_pair(1, 4, "Last:",   self.v_last,  w=18)

        # Row 2: Nickname
        self.ent_nick = add_pair(2, 0, "Nickname:", self.v_nick, w=18)

        # Row 3: Email / Phone
        self.ent_email = add_pair(3, 0, "Email:", self.v_email, w=30)
        self.ent_phone = add_pair(3, 4, "Phone:", self.v_phone, w=14)

        # Row 4: Addr1 (wide)
        ttk.Label(grid, text="Addr1:").grid(row=4, column=0, sticky="e", padx=(0, 8), pady=2)
        self.ent_addr1 = ttk.Entry(grid, textvariable=self.v_addr1, width=50)
        self.ent_addr1.grid(row=4, column=1, columnspan=5, sticky="w", pady=2)

        # Row 5: Addr2 (wide)
        ttk.Label(grid, text="Addr2:").grid(row=5, column=0, sticky="e", padx=(0, 8), pady=2)
        self.ent_addr2 = ttk.Entry(grid, textvariable=self.v_addr2, width=50)
        self.ent_addr2.grid(row=5, column=1, columnspan=5, sticky="w", pady=2)

        # Row 6: City / State / Zip
        self.ent_city  = add_pair(6, 0, "City:",  self.v_city,  w=18)
        self.ent_state = add_pair(6, 2, "State:", self.v_state, w=6)
        self.ent_zip   = add_pair(6, 4, "Zip:",   self.v_zip,   w=10)

        # Row 7: DOB
        self.ent_dob = add_pair(7, 0, "DOB:", self.v_dob, w=14)


        # Buttons
        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="Save", command=self._save).pack(side="right")
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Unlink", command=self._unlink_current).pack(side="left")

        # apply initial linked lock if already linked
        if self._linked_client_id:
            # try infer is_company from label match
            if self._linked_client_label and self._linked_client_label in self._label_to_is_company:
                self._linked_is_company = bool(self._label_to_is_company[self._linked_client_label])
            self._apply_link_lock(is_linked=True)

        self.protocol("WM_DELETE_WINDOW", self._cancel)

    # ---------------- autocomplete behavior ----------------

    def _match_labels(self, prefix: str) -> list[str]:
        tokens = (prefix or "").strip().lower().split()
        if not tokens:
            return self._labels
        return [lab for lab in self._labels if _matches_tokens(lab, tokens)]

    def _refresh_popup(self):
        items = self._match_labels(self.v_link.get())
        self._popup.show(items)
        if self._popup.listbox.size() > 0:
            self._popup.listbox.selection_clear(0, "end")
            self._popup.listbox.selection_set(0)
            self._popup.listbox.activate(0)

    def _on_link_key_press(self, event=None):
        """Triggered on key press to enable search even when there's an initial value"""
        # Allow search to work when user starts typing
        if event and event.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        # Schedule the search to happen after the key is processed
        self.after_idle(self._on_link_key, event)

    def _on_link_key(self, event=None):
        # Ignore navigation keys
        if event:
            if event.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
                return

        typed = self.v_link.get().strip()
        print(f"[LinkDialog] _on_link_key: typed='{typed}', total labels={len(self._labels)}")
        if not typed:
            self._popup.hide()
            return

        typed_lower = typed.lower()
        tokens = typed_lower.split()
        print(f"[LinkDialog] _on_link_key: tokens={tokens}")
        matches = [
            lab for lab in self._labels
            if all(t in lab.lower() for t in tokens)
        ]
        print(f"[LinkDialog] _on_link_key: found {len(matches)} matches")
        if len(matches) > 0:
            print(f"[LinkDialog] _on_link_key: first few matches: {matches[:5]}")

        if matches:
            print(f"[LinkDialog] _on_link_key: showing popup with {len(matches)} matches")
            self._popup.show(matches)
        else:
            print(f"[LinkDialog] _on_link_key: no matches, hiding popup")
            self._popup.hide()


    def _on_link_down(self, e=None):
        # open popup and move into listbox without mouse click
        if not self._popup.winfo_viewable():
            self._refresh_popup()
        if self._popup.winfo_viewable():
            try:
                self._popup.listbox.focus_set()
                self._popup.listbox.event_generate("<Down>")
            except Exception:
                pass
        return "break"

    def _on_link_up(self, e=None):
        if self._popup.winfo_viewable():
            try:
                self._popup.listbox.focus_set()
                self._popup.listbox.event_generate("<Up>")
            except Exception:
                pass
            return "break"
        return None

    def _on_link_return(self, event=None):
        # If popup is visible, choose highlighted item
        if self._popup and self._popup.winfo_viewable():
            txt = self._popup.current_text()
            if txt:
                self._on_choose_label(txt)
            self._popup.hide()
            return "break"

        # Otherwise try exact match
        val = self.v_link.get().strip()
        if val in self._label_to_id:
            self._on_choose_label(val)
            self._popup.hide()
            return "break"

        return "break"
        

    def _on_link_escape(self, e=None):
        self._popup.hide()
        return "break"

    def _on_link_focus_out(self, e=None):
        # delay hide so clicks into popup don't immediately close it
        self.after(80, self._maybe_hide_popup)

    def _maybe_hide_popup(self):
        try:
            foc = self.focus_get()
        except Exception:
            foc = None
        if foc in (self.ent_link, getattr(self._popup, "listbox", None)):
            return
        self._popup.hide()

    def _on_choose_label(self, label: str):
        label = str(label or "").strip()
        if not label:
            return

        cid = _norm_link_id(self._label_to_id.get(label, ""))
        if not cid:
            # Check if this label exists but just doesn't have an ID (not linkable)
            if label in self._labels:
                messagebox.showerror("Link", "That entity does not have a valid EIN/SSN id and cannot be linked.")
            else:
                messagebox.showerror("Link", "That entity does not have a valid EIN/SSN id.")
            return

        if cid in self._already_linked_ids:
            messagebox.showerror("Link", "That entity is already linked here.")
            return
        
        # Prevent entity from linking to itself
        if cid == self._this_client_id:
            messagebox.showerror("Link", "An entity cannot link to itself.")
            return

        # Store the normalized link ID
        self._linked_client_id = cid
        self._linked_client_label = label
        self._linked_is_company = bool(self._label_to_is_company.get(label, False))
        
        # Update role options now that we know the linked client type
        self._update_role_options()

        # If caller provided a resolver, pull real client data and fill fields as placeholders
        if callable(getattr(self, "_resolve_client", None)):
            try:
                c = self._resolve_client(cid)
            except Exception:
                c = None

            if isinstance(c, dict) and c:
                # Basic mapping (works for both individual + business style records)
                resolved_id = str(c.get("id") or "").strip()

                # If someone accidentally stored "client:..." in client["id"], ignore that here.
                if resolved_id.startswith(("ein:", "ssn:")):
                    self._linked_client_id = _norm_link_id(resolved_id)
                else:
                    self._linked_client_id = cid

                is_ind = bool(c.get("is_individual")) or (str(c.get("entity_type") or "").strip().lower() == "individual")

                # For editing, use existing values if present, otherwise use client data as placeholders
                if is_ind:
                    if not self.v_first.get().strip():
                        self.v_first.set(str(c.get("first_name") or "").strip())
                    if not self.v_mid.get().strip():
                        self.v_mid.set(str(c.get("middle_name") or "").strip())
                    if not self.v_last.get().strip():
                        self.v_last.set(str(c.get("last_name") or "").strip())
                    if not self.v_nick.get().strip():
                        self.v_nick.set(str(c.get("nickname") or "").strip())
                    # display name if you store it
                    if not self.v_name.get().strip():
                        self.v_name.set(str(c.get("name") or "").strip())
                else:
                    # business
                    if not self.v_name.get().strip():
                        self.v_name.set(str(c.get("name") or "").strip() or self.v_name.get())

                # Fill contact and address fields as placeholders if empty
                if not self.v_email.get().strip():
                    self.v_email.set(str(c.get("email") or "").strip())
                if not self.v_phone.get().strip():
                    self.v_phone.set(str(c.get("phone") or "").strip())
                if not self.v_addr1.get().strip():
                    self.v_addr1.set(str(c.get("addr1") or "").strip())
                if not self.v_addr2.get().strip():
                    self.v_addr2.set(str(c.get("addr2") or "").strip())
                if not self.v_city.get().strip():
                    self.v_city.set(str(c.get("city") or "").strip())
                if not self.v_state.get().strip():
                    self.v_state.set(str(c.get("state") or "").strip())
                if not self.v_zip.get().strip():
                    self.v_zip.set(str(c.get("zip") or "").strip())
                if not self.v_dob.get().strip():
                    self.v_dob.set(str(c.get("dob") or "").strip())

        # Fallback: if no resolver data (or resolver didn't set name), use label base
        if not self.v_name.get().strip():
            self.v_name.set(label.split(" (", 1)[0].strip())

        # Update role options based on client types
        self._update_role_options()
        
        # role rules
        if self._linked_is_company:
            # If linking to a business
            if self._this_is_individual:
                # Individual → Business: only "business" role allowed
                self.v_role.set("business")
                self.cmb_role.configure(state="disabled")
            else:
                # Business → Business: use provided role or default
                if not self.v_role.get().strip():
                    self.v_role.set("business")
                self.cmb_role.configure(state="readonly")
        else:
            # Linking to an individual
            if self._this_is_individual:
                # Individual → Individual: show spouse, parent, child, relative
                if self.v_role.get().strip().lower() not in ("spouse", "parent", "child", "relative"):
                    self.v_role.set("")
            else:
                # Business → Individual: show business owner, employee, officer
                if self.v_role.get().strip().lower() not in ("business owner", "employee", "officer"):
                    self.v_role.set("")
            self.cmb_role.configure(state="readonly")

        self.v_link.set(label)
        self._popup.hide()
        self._apply_link_lock(is_linked=True)

    # ---------------- link lock / unlink ----------------

    def _apply_link_lock(self, is_linked: bool):
        # if linked => user cannot modify filled fields
        state = "disabled" if is_linked else "normal"
        for w in (
            self.ent_name, self.ent_first, self.ent_mid, self.ent_last, self.ent_nick,
            self.ent_email, self.ent_phone,
            self.ent_addr1, self.ent_addr2, self.ent_city, self.ent_state, self.ent_zip, self.ent_dob
        ):
            try:
                w.configure(state=state)
            except Exception:
                pass

    def _update_role_options(self):
        """Update role combobox values based on this client and linked client types."""
        if not self._linked_client_id:
            # No link yet - show all options
            self.cmb_role["values"] = ("spouse", "parent", "child", "relative", "employee", "officer", "business owner", "business")
            return
        
        if self._this_is_individual:
            if self._linked_is_company:
                # Individual → Business: only "business" role
                self.cmb_role["values"] = ("business",)
            else:
                # Individual → Individual: spouse, parent, child, relative
                self.cmb_role["values"] = ("spouse", "parent", "child", "relative")
        else:
            if self._linked_is_company:
                # Business → Business: business role
                self.cmb_role["values"] = ("business",)
            else:
                # Business → Individual: business owner, employee, officer
                self.cmb_role["values"] = ("business owner", "employee", "officer")

    def _unlink_current(self):
        # unlink restores editability; does not delete user-entered values
        self._linked_client_id = ""
        self._linked_client_label = ""
        self._linked_is_company = False
        self._update_role_options()
        self.cmb_role.configure(state="readonly")
        if self.v_role.get().strip().lower() == "business":
            self.v_role.set("")
        self._apply_link_lock(is_linked=False)

    # ---------------- save/cancel ----------------

    def _save(self):
        role = (self.v_role.get() or "").strip().lower()

        # if linked, enforce ID requirement on BOTH sides
        if self._linked_client_id:
            if not self._this_client_id:
                messagebox.showerror(
                    "Link",
                    "Cannot link: this entity is missing its required ID.\n\n"
                    "Business must have EIN, Personnel must have SSN."
                )
                return

            if not _is_valid_link_id(self._linked_client_id):
                messagebox.showerror(
                    "Link",
                    "Cannot link: the selected entity is missing a valid link id.\n\n"
                    "Expected one of: client:<id>, ein:<9>, ssn:<9>."
                )
                return


            if self._linked_is_company:
                # Linking to a business
                if self._this_is_individual:
                    # Individual → Business: only "business" role
                    role = "business"
                else:
                    # Business → Business: use provided role or default to "business"
                    role = role or "business"
            else:
                # Linking to an individual
                if self._this_is_individual:
                    # Individual → Individual: spouse, parent, child, relative
                    if role not in ("spouse", "parent", "child", "relative"):
                        messagebox.showerror("Role", "For linking individual to individual, you must choose role: spouse, parent, child, or relative.")
                        return
                else:
                    # Business → Individual: business owner, employee, officer
                    if role not in ("business owner", "employee", "officer"):
                        messagebox.showerror("Role", "For linking business to individual, you must choose role: business owner, employee, or officer.")
                        return
        else:
            # manual entry: role optional, but normalize if they typed
            valid_roles = ("spouse", "parent", "child", "relative", "employee", "officer", "business owner", "business")
            if role and role not in valid_roles:
                role = ""

        self.result = {
            "name": self.v_name.get().strip(),
            "first_name": self.v_first.get().strip(),
            "middle_name": self.v_mid.get().strip(),
            "last_name": self.v_last.get().strip(),
            "nickname": self.v_nick.get().strip(),
            "email": self.v_email.get().strip(),
            "phone": self.v_phone.get().strip(),
            "addr1": self.v_addr1.get().strip(),
            "addr2": self.v_addr2.get().strip(),
            "city": self.v_city.get().strip(),
            "state": self.v_state.get().strip(),
            "zip": self.v_zip.get().strip(),
            "dob": self.v_dob.get().strip(),
            "role": role,
            "id": self._linked_client_id,  # Use "id" instead of "linked_client_id"
        }

        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()
