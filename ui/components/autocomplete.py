import tkinter as tk
from tkinter import ttk
from typing import Optional, Sequence, Callable

try:
    from styles.new_ui import NewUI
except Exception:
    class NewUI:
        BORDER = "#2b2b2b"

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

        self.listbox.bind("<Up>", self._lb_up)
        self.listbox.bind("<Down>", self._lb_down)

        self.listbox.bind("<Return>", self._choose)
        self.bind("<FocusOut>", self._maybe_hide)
        self.listbox.bind("<FocusOut>", self._maybe_hide)
        self.listbox.bind("<Escape>", lambda e: self.hide())

    def show(self, items: list[str]):
        self.listbox.delete(0, tk.END)
        for s in items[:20]:
            self.listbox.insert(tk.END, s)
        if not items:
            self.hide(); return
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

    def _choose(self, *_):
        txt = self.current_text()
        if txt is None: return
        self.on_choose(txt)
        self.hide()

    def hide(self):
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
