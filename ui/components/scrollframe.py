import tkinter as tk
from tkinter import ttk

class ScrollFrame(ttk.Frame):
    """Simple vertical scrollable frame for long dialogs."""
    def __init__(self, master, height=520):
        super().__init__(master)

        self.canvas = tk.Canvas(self, highlightthickness=0, height=height)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)

        # Keep the window id so we can resize it to match the canvas width
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        def _on_inner_configure(_e=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_configure(e):
            # Force inner frame to match visible canvas width
            self.canvas.itemconfigure(self._win, width=e.width)

        self.inner.bind("<Configure>", _on_inner_configure)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        def _on_mousewheel(event):
            if not self.canvas.winfo_exists():
                return
            if event.delta:  # Windows / macOS
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:  # X11: Button-4/5
                if event.num == 4:
                    self.canvas.yview_scroll(-3, "units")
                elif event.num == 5:
                    self.canvas.yview_scroll(+3, "units")

        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind_all("<Button-4>", _on_mousewheel)
        self.canvas.bind_all("<Button-5>", _on_mousewheel)
