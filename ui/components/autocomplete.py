import tkinter as tk
from tkinter import ttk
from typing import Optional, Sequence, Callable

try:
    from styles.new_ui import NewUI
except Exception:
    class NewUI:
        BORDER = "#2b2b2b"

# Active suggestion popups (hide_all on app focus loss)
_ACTIVE_POPUPS: list["AutocompletePopup"] = []


def hide_all_autocomplete_popups():
    for p in list(_ACTIVE_POPUPS):
        try:
            p.hide()
        except Exception:
            pass


def should_keep_autocomplete_open_for_widget(widget) -> bool:
    """True if click/focus is on an autocomplete popup or its anchor entry."""
    w = widget
    for p in list(_ACTIVE_POPUPS):
        try:
            if not p.winfo_viewable():
                continue
            if str(w).startswith(str(p)):
                return True
            a = getattr(p, "anchor", None)
            if a is not None and (w is a or str(w).startswith(str(a))):
                return True
        except Exception:
            continue
    return False


class AutocompletePopup(tk.Toplevel):
    def __init__(self, master, anchor_entry: tk.Entry, on_choose: Callable[[str], None]):
        super().__init__(master)
        self.withdraw()
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.anchor = anchor_entry
        self.on_choose = on_choose

        self.configure(bg=NewUI.BORDER)
        self.listbox = tk.Listbox(self, height=8, activestyle="none",
                                  bd=0, highlightthickness=0,
                                  relief="flat", font=("Segoe UI", 10))
        self.listbox.pack(fill="both", expand=True, padx=1, pady=1)
        self.listbox.bind("<ButtonRelease-1>", self._on_click_choose)
        self.listbox.bind("<Button-1>", self._on_mouse_down)
        self.listbox.bind("<Double-Button-1>", self._choose)

        self.listbox.unbind("<Up>")
        self.listbox.unbind("<Down>")
        self.listbox.unbind("<Left>")
        self.listbox.unbind("<Right>")

        self.listbox.bind("<Up>", self._lb_up)
        self.listbox.bind("<Down>", self._lb_down)
        self.listbox.bind("<Left>", self._lb_horizontal)
        self.listbox.bind("<Right>", self._lb_horizontal)

        self.listbox.bind("<Return>", self._choose)
        self.bind("<FocusOut>", self._maybe_hide)
        self.listbox.bind("<FocusOut>", self._maybe_hide)
        self.listbox.bind("<Escape>", lambda e: self.hide())

        self.listbox.bind("<MouseWheel>", self._on_listbox_wheel)
        self.listbox.bind("<Button-4>", self._on_listbox_wheel)
        self.listbox.bind("<Button-5>", self._on_listbox_wheel)
        self.bind("<MouseWheel>", self._on_listbox_wheel)

    def _on_listbox_wheel(self, event=None):
        if not self.winfo_viewable():
            return
        if event.delta:
            self.listbox.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            if event.num == 4:
                self.listbox.yview_scroll(-3, "units")
            elif event.num == 5:
                self.listbox.yview_scroll(3, "units")
        return "break"

    def show(self, items: list[str]):
        self.listbox.delete(0, tk.END)
        for s in items[:20]:
            self.listbox.insert(tk.END, s)
        if not items:
            self.hide(); return
        if self not in _ACTIVE_POPUPS:
            _ACTIVE_POPUPS.append(self)
        x = self.anchor.winfo_rootx()
        y = self.anchor.winfo_rooty() + self.anchor.winfo_height()
        w = self.anchor.winfo_width()
        h = min(256, 22 * len(items))
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.deiconify()
        if self.listbox.size() > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)

    def move_selection(self, delta: int):
        if not self.winfo_viewable(): return
        if self.listbox.size() == 0: return
        cur = self.listbox.curselection()
        i = cur[0] if cur else 0
        i = max(0, min(self.listbox.size()-1, i + delta))
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(i)
        self.listbox.activate(i)
        try:
            self.listbox.see(i)
        except Exception:
            pass

    def bind_entry_arrows(self, entry, *, on_open=None):
        """Navigate this popup from an Entry with Up/Down/Left/Right.

        Up/Down move the highlight. Left/Right keep caret movement in the entry.
        Never uses event_generate — synthesizing those keys re-delivers them to
        the focused entry and overflows Tkinter's EventType conversion.
        """
        def _vertical(event):
            delta = 1 if event.keysym == "Down" else -1
            if not self.winfo_viewable():
                if on_open is not None:
                    on_open()
                return "break"
            self.move_selection(delta)
            return "break"

        def _horizontal(_event):
            return None

        entry.bind("<Up>", _vertical)
        entry.bind("<Down>", _vertical)
        entry.bind("<Left>", _horizontal)
        entry.bind("<Right>", _horizontal)

    def current_text(self) -> Optional[str]:
        cur = self.listbox.curselection()
        if not cur: return None
        return self.listbox.get(cur[0])

    def _lb_up(self, event=None):
        self.move_selection(-1)
        return "break"

    def _lb_down(self, event=None):
        self.move_selection(+1)
        return "break"

    def _lb_horizontal(self, event=None):
        return "break"

    def _choose(self, *_):
        txt = self.current_text()
        if txt is None: return
        self.on_choose(txt)
        self.hide()

    def hide(self):
        try:
            if self in _ACTIVE_POPUPS:
                _ACTIVE_POPUPS.remove(self)
        except ValueError:
            pass
        try:
            self.withdraw()
            self.update_idletasks()
        except Exception:
            pass


    def focus_listbox(self):
        self.listbox.focus_set()

    def _on_mouse_down(self, e):
        i = self.listbox.nearest(e.y)
        if 0 <= i < self.listbox.size():
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(i)
            self.listbox.activate(i)

    def _on_click_choose(self, e):
        # let the Listbox update selection, then choose
        self.after(1, self._choose)

    def _maybe_hide(self, _e=None):
        w = self.focus_get()

        # If focus is on the anchor Entry, keep popup open
        try:
            if w == self.anchor or str(w).startswith(str(self.anchor)):
                return
        except Exception:
            pass

        # If focus is inside this popup, keep it open
        if w and str(w).startswith(str(self)):
            return

        self.hide()
