import tkinter as tk
from tkinter import ttk, font as tkfont

class NewUI:
    # Palette
    ACCENT        = "#4F46E5"   # indigo-600
    ACCENT_HOVER  = "#4338CA"   # indigo-700
    TEXT          = "#0F172A"   # slate-900
    MUTED         = "#6B7280"   # gray-500
    BG            = "#F8FAFC"   # slate-50
    PANEL         = "#FFFFFF"   # white
    BORDER        = "#E5E7EB"   # gray-200
    BORDER_STRONG = "#CBD5E1"   # slate-300
    ROW_ALT       = "#F1F5F9"   # slate-100
    FOCUS_BG      = "#EEF2FF"   # indigo-50

    @staticmethod
    def install(root: tk.Tk):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # --- Fonts / base ---
        base_font = tkfont.nametofont("TkDefaultFont")
        base_font.configure(size=10)
        heading_font = tkfont.Font(family=base_font.actual("family"), size=11, weight="bold")
        root.option_add("*Font", base_font)
        root.configure(bg=NewUI.BG)

        # Make default labels match the page bg (kills the gray highlight look)
        style.configure("TLabel", background=NewUI.BG, foreground=NewUI.TEXT)

        # Headline/secondary labels that sit on white cards
        style.configure("Header.TLabel",
                        background=NewUI.PANEL,
                        font=("Segoe UI", 18, "bold"),
                        foreground=NewUI.TEXT)
        style.configure("Subtle.TLabel",
                        background=NewUI.PANEL,
                        font=("Segoe UI", 11, "italic"),
                        foreground=NewUI.MUTED)

        # --- Frames / Cards / Notebook ---
        style.configure("TFrame", background=NewUI.BG)

        # Use a soft card only where you actually want a card
        style.configure("Card.TFrame",
                        background=NewUI.PANEL,
                        borderwidth=1,
                        relief="solid",
                        bordercolor=NewUI.BORDER,
                        lightcolor=NewUI.BORDER,
                        darkcolor=NewUI.BORDER)

        style.configure("TNotebook", background=NewUI.BG, borderwidth=0)
        style.configure("TNotebook.Tab",
                        padding=(16, 10),
                        font=heading_font)
        style.map("TNotebook.Tab",
                  background=[("selected", NewUI.PANEL), ("!selected", NewUI.BG)],
                  foreground=[("selected", NewUI.TEXT), ("!selected", NewUI.MUTED)])

        # --- Buttons ---
        # Base button (secondary): keep white, never turn black
        style.configure("TButton",
                        padding=(14, 9),
                        relief="flat",
                        background=NewUI.PANEL,
                        foreground=NewUI.TEXT,
                        borderwidth=1,
                        bordercolor=NewUI.BORDER_STRONG,
                        lightcolor=NewUI.BORDER_STRONG,
                        darkcolor=NewUI.BORDER_STRONG)
        style.map("TButton",
                  background=[
                      ("active", "#F8FAFF"),   # <-- fixed (was '##F8FAFF')
                      ("pressed", "#F1F5F9")
                  ],
                  foreground=[
                      ("disabled", "#9CA3AF"),
                      ("active", NewUI.TEXT),
                      ("pressed", NewUI.TEXT)
                  ],
                  relief=[("pressed", "sunken")],
                  bordercolor=[
                      ("focus", NewUI.ACCENT),
                      ("active", NewUI.BORDER_STRONG)
                  ])

        # Primary button
        style.configure("Accent.TButton",
                        padding=(16, 10),
                        relief="flat",
                        background=NewUI.ACCENT,
                        foreground="white",
                        borderwidth=0)
        style.map("Accent.TButton",
                  background=[("active", NewUI.ACCENT_HOVER), ("pressed", NewUI.ACCENT_HOVER)],
                  foreground=[("disabled", "#E5E7EB"), ("active", "white"), ("pressed", "white")])

        # Outline button
        style.configure("Outline.TButton",
                        padding=(14, 9),
                        background=NewUI.PANEL,
                        foreground=NewUI.ACCENT,
                        borderwidth=1,
                        bordercolor=NewUI.ACCENT,
                        lightcolor=NewUI.ACCENT,
                        darkcolor=NewUI.ACCENT)
        style.map("Outline.TButton",
                  background=[("active", NewUI.FOCUS_BG)],
                  foreground=[("active", NewUI.ACCENT)],
                  bordercolor=[("focus", NewUI.ACCENT)])

        # --- Entries & Combobox ---
        for w in ("TEntry", "TCombobox"):
            style.configure(w,
                            padding=(10, 8),
                            fieldbackground="white",
                            borderwidth=1,
                            bordercolor=NewUI.BORDER_STRONG,
                            lightcolor=NewUI.BORDER_STRONG,
                            darkcolor=NewUI.BORDER_STRONG)
            style.map(w,
                      fieldbackground=[("focus", "white")],
                      bordercolor=[("focus", NewUI.ACCENT)])

        root.option_add("*Entry.insertWidth", 2)

        # --- Treeview ---
        style.configure("Modern.Treeview",
                        background="white",
                        fieldbackground="white",
                        bordercolor=NewUI.BORDER,
                        lightcolor=NewUI.BORDER,
                        darkcolor=NewUI.BORDER,
                        rowheight=int(base_font.metrics("linespace") * 1.5))
        style.configure("Modern.Treeview.Heading",
                        background=NewUI.BG,
                        foreground=NewUI.TEXT,
                        relief="flat",
                        font=("Segoe UI", 10, "bold"))
        style.map("Modern.Treeview.Heading",
                  background=[("active", NewUI.PANEL)])

    @staticmethod
    def stripe_tree(tree: ttk.Treeview):
        tree.tag_configure("oddrow",  background="white")
        tree.tag_configure("evenrow", background=NewUI.ROW_ALT)
        for i, iid in enumerate(tree.get_children("")):
            tree.item(iid, tags=("evenrow" if i % 2 == 0 else "oddrow",))
