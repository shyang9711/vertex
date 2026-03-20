import tkinter as tk
from tkinter import ttk

# All live ScrollFrame instances (for routed mousewheel)
_SCROLLFRAMES: list["ScrollFrame"] = []


def _widget_under_master(widget, master) -> bool:
    w = widget
    while w is not None:
        if w is master:
            return True
        try:
            w = w.master
        except Exception:
            break
    return False


def _is_under_autocomplete_popup(widget) -> bool:
    w = widget
    while w is not None:
        try:
            if w.__class__.__name__ == "AutocompletePopup":
                return True
        except Exception:
            pass
        try:
            w = w.master
        except Exception:
            break
    return False


class ScrollFrame(ttk.Frame):
    """Vertical scrollable frame. Mousewheel scrolls only when pointer is over this frame's content."""

    def __init__(self, master, height=520):
        super().__init__(master)

        self.canvas = tk.Canvas(self, highlightthickness=0, height=height)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)

        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        def _on_inner_configure(_e=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_configure(e):
            self.canvas.itemconfigure(self._win, width=e.width)

        self.inner.bind("<Configure>", _on_inner_configure)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        _SCROLLFRAMES.append(self)
        self.bind("<Destroy>", self._on_destroy, add=True)

        ScrollFrame._ensure_global_router(self.winfo_toplevel())

    def _on_destroy(self, _e=None):
        try:
            _SCROLLFRAMES.remove(self)
        except ValueError:
            pass

    _router_installed = False

    @classmethod
    def _ensure_global_router(cls, toplevel):
        if cls._router_installed:
            return
        try:
            toplevel.bind_all("<MouseWheel>", cls._route_mousewheel, add=True)
            toplevel.bind_all("<Button-4>", cls._route_mousewheel, add=True)
            toplevel.bind_all("<Button-5>", cls._route_mousewheel, add=True)
            cls._router_installed = True
        except Exception:
            pass

    @staticmethod
    def _route_mousewheel(event):
        if _is_under_autocomplete_popup(event.widget):
            return
        for sf in list(_SCROLLFRAMES):
            try:
                if not sf.winfo_exists():
                    continue
            except Exception:
                continue
            if _widget_under_master(event.widget, sf.inner) or event.widget == sf.canvas:
                sf._do_scroll(event)
                return "break"
        return None

    def _do_scroll(self, event):
        if not self.canvas.winfo_exists():
            return
        if event.delta:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            if event.num == 4:
                self.canvas.yview_scroll(-3, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(3, "units")
