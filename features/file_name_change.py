import os
import tkinter as tk
from tkinter import filedialog, messagebox


def split_name_ext(filename: str):
    base, ext = os.path.splitext(filename)
    return base, ext


def safe_delete_front(base: str, n: int) -> str:
    n = max(0, int(n))
    if len(base) <= 1:
        return base
    # leave at least 1 char
    n = min(n, len(base) - 1)
    return base[n:]


def safe_delete_end(base: str, n: int) -> str:
    n = max(0, int(n))
    if len(base) <= 1:
        return base
    # leave at least 1 char
    n = min(n, len(base) - 1)
    return base[:-n] if n > 0 else base

def make_unique_path(folder: str, filename: str, used_names: set[str]) -> str:
    """
    If filename already exists (on disk or in used_names),
    append (1), (2), ... before extension until unique.
    """
    base, ext = os.path.splitext(filename)
    candidate = filename
    counter = 1

    while (
        os.path.exists(os.path.join(folder, candidate))
        or candidate.lower() in used_names
    ):
        candidate = f"{base} ({counter}){ext}"
        counter += 1

    used_names.add(candidate.lower())
    return os.path.join(folder, candidate)


class BatchRenameGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Batch Rename Files")
        self.root.geometry("760x460")

        # full paths of currently tracked files
        self.file_paths: list[str] = []

        # --- Top controls ---
        top = tk.Frame(root)
        top.pack(fill="x", padx=10, pady=10)

        tk.Button(top, text="Select File(s)...", command=self.select_files, width=18).pack(side="left")

        self.status_var = tk.StringVar(value="No files selected.")
        tk.Label(top, textvariable=self.status_var, anchor="w").pack(side="left", padx=12, fill="x", expand=True)

        # --- File list (scrollable) ---
        list_frame = tk.LabelFrame(root, text="Selected Files (updates after rename)")
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.listbox = tk.Listbox(list_frame, height=12)
        yscroll = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=yscroll.set)

        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        yscroll.pack(side="left", fill="y", padx=(0, 8), pady=8)

        # --- Operations area ---
        ops = tk.Frame(root)
        ops.pack(fill="x", padx=10, pady=(0, 10))

        # Add text panel
        add_panel = tk.LabelFrame(ops, text="Add Text")
        add_panel.pack(side="left", fill="both", expand=True, padx=(0, 8))

        tk.Label(add_panel, text="Text:").grid(row=0, column=0, sticky="e", padx=8, pady=8)
        self.add_text = tk.Entry(add_panel, width=30)
        self.add_text.grid(row=0, column=1, sticky="w", padx=8, pady=8)

        btn_row = tk.Frame(add_panel)
        btn_row.grid(row=1, column=0, columnspan=2, pady=(0, 10))

        tk.Button(btn_row, text="Add to Front", width=14, command=self.add_to_front).pack(side="left", padx=6)
        tk.Button(btn_row, text="Add to End", width=14, command=self.add_to_end).pack(side="left", padx=6)

        # Delete panel
        del_panel = tk.LabelFrame(ops, text="Delete Characters (base name only; keeps extension)")
        del_panel.pack(side="left", fill="both", expand=True, padx=(8, 0))

        tk.Label(del_panel, text="Count:").grid(row=0, column=0, sticky="e", padx=8, pady=8)
        self.del_count = tk.Entry(del_panel, width=10)
        self.del_count.insert(0, "1")
        self.del_count.grid(row=0, column=1, sticky="w", padx=8, pady=8)

        del_btn_row = tk.Frame(del_panel)
        del_btn_row.grid(row=1, column=0, columnspan=2, pady=(0, 10))

        tk.Button(del_btn_row, text="Remove Front", width=14, command=self.remove_front).pack(side="left", padx=6)
        tk.Button(del_btn_row, text="Remove End", width=14, command=self.remove_end).pack(side="left", padx=6)

        # Bottom controls
        bottom = tk.Frame(root)
        bottom.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(bottom, text="Finish", command=self.finish, width=12).pack(side="right")
        tk.Button(bottom, text="Clear List", command=self.clear_list, width=12).pack(side="right", padx=8)

    # ---------------- UI Helpers ----------------
    def select_files(self):
        paths = filedialog.askopenfilenames(title="Select file(s) to rename")
        if not paths:
            return
        self.file_paths = list(paths)
        self.refresh_listbox()
        self.status_var.set(f"{len(self.file_paths)} file(s) selected.")

    def clear_list(self):
        self.file_paths = []
        self.listbox.delete(0, tk.END)
        self.status_var.set("No files selected.")

    def finish(self):
        self.root.destroy()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for p in self.file_paths:
            self.listbox.insert(tk.END, os.path.basename(p))

    # ---------------- Rename Core ----------------
    def apply_rename(self, transform_fn):
        if not self.file_paths:
            messagebox.showwarning("No files", "Please select files first.")
            return

        errors = []
        updated_paths = []

        # Track names used in this batch to avoid internal collisions
        used_names = set()

        for old_path in self.file_paths:
            folder = os.path.dirname(old_path)
            old_file = os.path.basename(old_path)

            base, ext = split_name_ext(old_file)
            try:
                new_base, new_ext = transform_fn(base, ext)

                # Final safety: never allow empty base
                if not new_base or len(new_base) < 1:
                    new_base = base[:1] if base else "X"

                new_file = f"{new_base}{new_ext}"

                # 🔒 Ensure uniqueness
                new_path = make_unique_path(folder, new_file, used_names)

                # If name unchanged, keep it
                if os.path.abspath(new_path) == os.path.abspath(old_path):
                    updated_paths.append(old_path)
                    used_names.add(os.path.basename(old_path).lower())
                    continue

                os.rename(old_path, new_path)
                updated_paths.append(new_path)

            except Exception as e:
                updated_paths.append(old_path)
                errors.append(f"{old_file} -> ERROR: {e}")

        self.file_paths = updated_paths
        self.refresh_listbox()

        if errors:
            messagebox.showerror("Completed with errors", "\n".join(errors))
            self.status_var.set(f"Renamed with {len(errors)} error(s).")
        else:
            self.status_var.set("Renamed successfully.")


    # ---------------- Add Actions ----------------
    def add_to_front(self):
        text = self.add_text.get()
        if text is None or text == "":
            messagebox.showwarning("Missing text", "Enter text to add.")
            return

        def transform(base, ext):
            return f"{text}{base}", ext

        self.apply_rename(transform)

    def add_to_end(self):
        text = self.add_text.get()
        if text is None or text == "":
            messagebox.showwarning("Missing text", "Enter text to add.")
            return

        def transform(base, ext):
            return f"{base}{text}", ext

        self.apply_rename(transform)

    # ---------------- Delete Actions ----------------
    def _get_del_n(self) -> int:
        raw = self.del_count.get().strip()
        if raw == "":
            return 0
        try:
            n = int(raw)
            if n < 0:
                n = 0
            return n
        except ValueError:
            messagebox.showerror("Invalid count", "Delete count must be an integer (0 or greater).")
            return None

    def remove_front(self):
        n = self._get_del_n()
        if n is None:
            return

        def transform(base, ext):
            return safe_delete_front(base, n), ext

        self.apply_rename(transform)

    def remove_end(self):
        n = self._get_del_n()
        if n is None:
            return

        def transform(base, ext):
            return safe_delete_end(base, n), ext

        self.apply_rename(transform)


if __name__ == "__main__":
    root = tk.Tk()
    app = BatchRenameGUI(root)
    root.mainloop()
