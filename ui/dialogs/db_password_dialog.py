"""Password prompts for the encrypted Vertex database."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import NamedTuple

try:
    from vertex.utils.db_crypto import (
        MIN_PASSWORD_LEN,
        SqlcipherUnavailableError,
        current_key,
        encrypt_plaintext_file,
        is_plaintext_sqlite,
        keyring_get,
        remember_key,
        set_session_key,
        try_unlock_from_keyring,
        verify_open,
    )
except ModuleNotFoundError:
    from utils.db_crypto import (
        MIN_PASSWORD_LEN,
        SqlcipherUnavailableError,
        current_key,
        encrypt_plaintext_file,
        is_plaintext_sqlite,
        keyring_get,
        remember_key,
        set_session_key,
        try_unlock_from_keyring,
        verify_open,
    )


class PasswordResult(NamedTuple):
    password: str
    remember: bool


def _db_path():
    try:
        from vertex.utils.sql_store import db_path
    except ModuleNotFoundError:
        from utils.sql_store import db_path
    return db_path()


def _center(win: tk.Toplevel, parent) -> None:
    try:
        win.update_idletasks()
        pw = win.winfo_width()
        ph = win.winfo_height()
        if parent is not None:
            x = parent.winfo_rootx() + max(0, (parent.winfo_width() - pw) // 2)
            y = parent.winfo_rooty() + max(0, (parent.winfo_height() - ph) // 2)
        else:
            x = (win.winfo_screenwidth() - pw) // 2
            y = (win.winfo_screenheight() - ph) // 2
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass


class _PasswordDialog(tk.Toplevel):
    def __init__(
        self,
        master,
        *,
        title: str,
        message: str,
        confirm: bool,
        old_password: bool = False,
    ):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result: PasswordResult | None = None
        self._need_confirm = confirm
        self._need_old = old_password
        self.old_value = ""

        frm = ttk.Frame(self, padding=16)
        frm.grid(row=0, column=0, sticky="nsew")
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text=message, wraplength=420, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )

        row = 1
        self._old = None
        if old_password:
            ttk.Label(frm, text="Current password").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            self._old = ttk.Entry(frm, width=36, show="*")
            self._old.grid(row=row, column=1, sticky="ew", pady=4)
            row += 1

        ttk.Label(frm, text="Password" if not confirm else "New password").grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self._pw = ttk.Entry(frm, width=36, show="*")
        self._pw.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        self._pw2 = None
        if confirm:
            ttk.Label(frm, text="Confirm").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            self._pw2 = ttk.Entry(frm, width=36, show="*")
            self._pw2.grid(row=row, column=1, sticky="ew", pady=4)
            row += 1

        self._remember = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frm,
            text="Remember password on this PC",
            variable=self._remember,
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(8, 4))
        row += 1

        ttk.Label(
            frm,
            text=f"Use at least {MIN_PASSWORD_LEN} characters. "
            "Anyone with this password can open a copy of the database file.",
            wraplength=420,
            justify="left",
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 12))
        row += 1

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btns, text="OK", command=self._ok).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Return>", lambda _e: self._ok())
        self.bind("<Escape>", lambda _e: self._cancel())
        self.transient(master)
        _center(self, master)
        self.after(50, (self._old or self._pw).focus_set)
        self.grab_set()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()

    def _ok(self) -> None:
        if self._need_old:
            self.old_value = (self._old.get() if self._old else "") or ""
            if not self.old_value:
                messagebox.showerror("Database password", "Enter the current password.", parent=self)
                return
        pw = (self._pw.get() or "").strip()
        if len(pw) < MIN_PASSWORD_LEN:
            messagebox.showerror(
                "Database password",
                f"Password must be at least {MIN_PASSWORD_LEN} characters.",
                parent=self,
            )
            return
        if self._need_confirm:
            pw2 = (self._pw2.get() if self._pw2 else "") or ""
            if pw != pw2.strip():
                messagebox.showerror("Database password", "Passwords do not match.", parent=self)
                return
        self.result = PasswordResult(password=pw, remember=bool(self._remember.get()))
        self.destroy()


def prompt_unlock_password(parent, message: str) -> PasswordResult | None:
    dlg = _PasswordDialog(parent, title="Unlock database", message=message, confirm=False)
    parent.wait_window(dlg)
    return dlg.result


def prompt_set_password(parent, message: str, *, title: str = "Encrypt database") -> PasswordResult | None:
    dlg = _PasswordDialog(parent, title=title, message=message, confirm=True)
    parent.wait_window(dlg)
    return dlg.result


def prompt_change_password(parent) -> tuple[str, PasswordResult] | None:
    dlg = _PasswordDialog(
        parent,
        title="Change database password",
        message="Choose a new password for vertex.db. Existing backups encrypted "
        "with the old password will be replaced.",
        confirm=True,
        old_password=True,
    )
    parent.wait_window(dlg)
    if dlg.result is None:
        return None
    return dlg.old_value, dlg.result


def _purge_plaintext_leftovers() -> None:
    try:
        try:
            from vertex.utils.sql_store import purge_plaintext_backups
        except ModuleNotFoundError:
            from utils.sql_store import purge_plaintext_backups
        purge_plaintext_backups()
    except Exception:
        pass


def unlock_or_encrypt_database(parent) -> bool:
    """
    Make sure vertex.db can be opened this session.

    - No database yet: OK (JSON until File → Migrate to SQL).
    - Plaintext SQLite: prompt for a new password and encrypt in place.
    - Encrypted SQLite: use the credential locker, else prompt.
    """
    try:
        path = _db_path()
    except Exception:
        return True
    try:
        if not path.exists() or path.stat().st_size < 64:
            return True
    except Exception:
        return True

    try:
        from sqlcipher3 import dbapi2 as _unused  # noqa: F401
    except Exception as e:
        messagebox.showerror(
            "Encrypted database",
            "Vertex needs the sqlcipher3 package to open the database.\n\n"
            f"{e}",
            parent=parent,
        )
        return False

    if is_plaintext_sqlite(path):
        result = prompt_set_password(
            parent,
            "vertex.db is not encrypted. Set a password to encrypt it.\n\n"
            "JSON copies of client data will be removed after they are stored "
            "in the encrypted database.",
        )
        if result is None:
            return False
        try:
            encrypt_plaintext_file(path, result.password)
        except Exception as e:
            messagebox.showerror(
                "Encrypt database",
                f"Could not encrypt vertex.db:\n{e}",
                parent=parent,
            )
            return False
        set_session_key(result.password)
        remember_key(result.password, result.remember)
        _purge_plaintext_leftovers()
        return True

    if try_unlock_from_keyring(path):
        _purge_plaintext_leftovers()
        return True

    while True:
        result = prompt_unlock_password(
            parent,
            "Enter the password for the encrypted database.",
        )
        if result is None:
            return False
        if verify_open(path, result.password):
            set_session_key(result.password)
            remember_key(result.password, result.remember)
            _purge_plaintext_leftovers()
            return True
        messagebox.showerror("Unlock database", "That password did not open vertex.db.", parent=parent)


def ensure_password_for_new_database(parent) -> bool:
    """Prompt for a password if this session does not already have one."""
    if current_key():
        return True
    remembered = keyring_get()
    if remembered:
        set_session_key(remembered)
        return True
    result = prompt_set_password(
        parent,
        "Set a password for the new encrypted database (data/vertex.db).\n\n"
        "JSON files will be copied in once, then removed from the data folder.",
        title="Create encrypted database",
    )
    if result is None:
        return False
    set_session_key(result.password)
    remember_key(result.password, result.remember)
    return True


def change_database_password(parent) -> bool:
    try:
        from vertex.utils.sql_store import change_db_password, sql_db_ready
    except ModuleNotFoundError:
        from utils.sql_store import change_db_password, sql_db_ready

    if not sql_db_ready():
        messagebox.showinfo(
            "Change database password",
            "There is no encrypted database yet.\n\nUse File → Migrate to SQL first.",
            parent=parent,
        )
        return False

    pair = prompt_change_password(parent)
    if pair is None:
        return False
    old, result = pair
    try:
        change_db_password(old, result.password)
    except SqlcipherUnavailableError as e:
        messagebox.showerror("Change database password", str(e), parent=parent)
        return False
    except Exception as e:
        messagebox.showerror(
            "Change database password",
            f"Could not change the password:\n{e}",
            parent=parent,
        )
        return False
    remember_key(result.password, result.remember)
    messagebox.showinfo(
        "Change database password",
        "The database password was changed.",
        parent=parent,
    )
    return True
