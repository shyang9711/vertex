"""
SQLCipher helpers: passphrase handling, keyring, encrypt-in-place.

The passphrase is kept in process memory after unlock. Optionally it is also
stored in the OS credential locker (Windows Credential Manager) so tools and
later launches can open the database without prompting.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

try:
    from vertex.utils.app_logging import get_logger
except ModuleNotFoundError:
    from utils.app_logging import get_logger

LOG = get_logger("db_crypto")

KEYRING_SERVICE = "Vertex"
KEYRING_USER = "sqlcipher"
MIN_PASSWORD_LEN = 8
SQLITE_HEADER = b"SQLite format 3\x00"

_SESSION_KEY: str | None = None


class DatabaseLockedError(RuntimeError):
    """Encrypted database could not be opened (no key or wrong key)."""


class SqlcipherUnavailableError(RuntimeError):
    """sqlcipher3 is not installed or failed to import."""


def get_sqlite():
    try:
        from sqlcipher3 import dbapi2 as sqlite
    except Exception as e:
        raise SqlcipherUnavailableError(
            "SQLCipher is required to open Vertex's database. "
            "Install it with: pip install sqlcipher3"
        ) from e
    return sqlite


def current_key() -> str | None:
    return _SESSION_KEY


def set_session_key(key: str | None) -> None:
    global _SESSION_KEY
    _SESSION_KEY = key if key else None


def is_plaintext_sqlite(path: Path) -> bool:
    try:
        p = Path(path)
        if not p.is_file() or p.stat().st_size < 16:
            return False
        with p.open("rb") as f:
            return f.read(16) == SQLITE_HEADER
    except Exception:
        return False


def sidecar_paths(path: Path) -> list[Path]:
    p = Path(path)
    return [Path(str(p) + suffix) for suffix in ("-journal", "-wal", "-shm")]


def remove_sidecars(path: Path) -> None:
    for extra in sidecar_paths(path):
        try:
            if extra.exists():
                extra.unlink()
        except Exception:
            pass


def _keyring_backend():
    try:
        import keyring
        return keyring
    except Exception:
        return None


def keyring_get() -> str | None:
    kr = _keyring_backend()
    if kr is None:
        return None
    try:
        value = kr.get_password(KEYRING_SERVICE, KEYRING_USER)
        return value if value else None
    except Exception as e:
        LOG.warning("Credential locker read failed: %s", e)
        return None


def keyring_set(password: str) -> bool:
    kr = _keyring_backend()
    if kr is None:
        return False
    try:
        kr.set_password(KEYRING_SERVICE, KEYRING_USER, password)
        return True
    except Exception as e:
        LOG.warning("Credential locker save failed: %s", e)
        return False


def keyring_delete() -> None:
    kr = _keyring_backend()
    if kr is None:
        return
    try:
        kr.delete_password(KEYRING_SERVICE, KEYRING_USER)
    except Exception:
        pass


def remember_key(password: str, remember: bool) -> None:
    if remember:
        if not keyring_set(password):
            LOG.warning("Password was not stored in the credential locker")
    else:
        keyring_delete()


def apply_key(conn: Any, key: str) -> None:
    escaped = (key or "").replace("'", "''")
    conn.execute(f"PRAGMA key = '{escaped}'")


def apply_rekey(conn: Any, new_key: str) -> None:
    escaped = (new_key or "").replace("'", "''")
    conn.execute(f"PRAGMA rekey = '{escaped}'")


def verify_open(path: Path, key: str | None) -> bool:
    """True when the file opens as a SQLite database with this key (or none)."""
    sqlite = get_sqlite()
    conn = sqlite.connect(str(path), timeout=5.0)
    try:
        if key:
            apply_key(conn, key)
        conn.execute("SELECT count(*) FROM sqlite_master").fetchone()
        return True
    except Exception:
        return False
    finally:
        conn.close()


def try_unlock_from_keyring(path: Path | None = None) -> bool:
    """Load a remembered passphrase and keep it if it opens the database."""
    if current_key():
        return True
    remembered = keyring_get()
    if not remembered:
        return False
    if path is not None and path.exists() and not is_plaintext_sqlite(path):
        if not verify_open(path, remembered):
            return False
    set_session_key(remembered)
    return True


def encrypt_plaintext_file(path: Path, key: str) -> None:
    """Turn an unencrypted SQLite file into a SQLCipher database with `key`."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(str(path))
    if not is_plaintext_sqlite(path):
        if verify_open(path, key):
            set_session_key(key)
            return
        raise RuntimeError(f"Database is already encrypted: {path}")

    sqlite = get_sqlite()
    tmp = path.with_name(path.name + ".enc.tmp")
    if tmp.exists():
        tmp.unlink()

    conn = sqlite.connect(str(path), timeout=30.0)
    try:
        conn.execute("ATTACH DATABASE ? AS encrypted KEY ?", (str(tmp), key))
        conn.execute("SELECT sqlcipher_export('encrypted')")
        conn.execute("DETACH DATABASE encrypted")
        conn.commit()
    finally:
        conn.close()

    if not tmp.exists() or tmp.stat().st_size < 64:
        raise RuntimeError("SQLCipher export produced an empty file")
    if is_plaintext_sqlite(tmp):
        try:
            tmp.unlink()
        except Exception:
            pass
        raise RuntimeError("SQLCipher export did not encrypt the database")
    if not verify_open(tmp, key):
        try:
            tmp.unlink()
        except Exception:
            pass
        raise RuntimeError("Encrypted database could not be reopened")

    bak = path.with_name(path.name + ".plain.bak")
    try:
        if bak.exists():
            bak.unlink()
        path.replace(bak)
        tmp.replace(path)
        try:
            bak.unlink()
        except Exception:
            LOG.warning("Could not delete plaintext leftover %s", bak)
    except Exception:
        if tmp.exists() and not path.exists():
            tmp.replace(path)
        raise
    finally:
        remove_sidecars(path)
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass

    set_session_key(key)
    LOG.info("Encrypted database in place: %s", path)


def purge_plaintext_db_copies(folder: Path) -> int:
    """Delete leftover unencrypted vertex_*.db snapshots."""
    n = 0
    try:
        folder = Path(folder)
        if not folder.is_dir():
            return 0
        for p in folder.glob("vertex_*.db"):
            if is_plaintext_sqlite(p):
                try:
                    p.unlink()
                    n += 1
                except Exception:
                    pass
        leftover = folder.parent / "vertex.db.plain.bak"
        if leftover.exists() and is_plaintext_sqlite(leftover):
            try:
                leftover.unlink()
                n += 1
            except Exception:
                pass
    except Exception as e:
        LOG.warning("Plaintext backup purge failed: %s", e)
    return n
