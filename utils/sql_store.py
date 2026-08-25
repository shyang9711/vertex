"""
SQLCipher persistence for Vertex.

JSON files are a one-time import source. After File → Migrate to SQL
(or after leftover JSON is copied into an existing database), those JSON
files are removed. Load/save then use only the encrypted database.
"""

from __future__ import annotations

import json
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

SCHEMA_VERSION = 1

try:
    from vertex.utils.app_logging import get_logger
    from vertex.utils.db_crypto import (
        DatabaseLockedError,
        apply_key,
        apply_rekey,
        current_key,
        get_sqlite,
        is_plaintext_sqlite,
        purge_plaintext_db_copies,
        set_session_key,
        try_unlock_from_keyring,
        verify_open,
    )
except ModuleNotFoundError:
    from utils.app_logging import get_logger
    from utils.db_crypto import (
        DatabaseLockedError,
        apply_key,
        apply_rekey,
        current_key,
        get_sqlite,
        is_plaintext_sqlite,
        purge_plaintext_db_copies,
        set_session_key,
        try_unlock_from_keyring,
        verify_open,
    )

LOG = get_logger("sql_store")

_LIST_TABLES = {
    "clients": "client_id",
    "account_managers": "manager_id",
    "tasks": "task_id",
}

_MISSING_JSON_MIGRATED = False
_MISSING_JSON_MIGRATING = False
_READY_CACHE: bool | None = None
_SKIP_DB_BACKUP = False
_LAST_BACKUP_MONO = 0.0
_BACKUP_DEBOUNCE_SEC = 2.0
BACKUP_KEEP = 7


def _io():
    try:
        from vertex.utils import io as m
    except ModuleNotFoundError:
        from utils import io as m
    return m


def db_path() -> Path:
    return _io().DATA_ROOT / "vertex.db"


def sql_db_ready() -> bool:
    """True when vertex.db exists, can be unlocked, and has a readable schema."""
    global _READY_CACHE
    if _READY_CACHE is True:
        return True
    path = db_path()
    try:
        if not path.exists() or path.stat().st_size < 64:
            return False
        if not is_plaintext_sqlite(path) and not current_key():
            try_unlock_from_keyring(path)
            if not current_key():
                return False
        with _connect() as conn:
            row = conn.execute(
                "SELECT value FROM meta WHERE key = 'schema_version'"
            ).fetchone()
            if row is not None:
                _READY_CACHE = True
                return True
            return False
    except Exception:
        return False


def _require_key(path: Path, *, creating: bool) -> str | None:
    key = current_key()
    if key:
        return key
    if creating and not path.exists():
        raise DatabaseLockedError("Set a database password before creating vertex.db.")
    if path.exists() and is_plaintext_sqlite(path):
        return None
    try_unlock_from_keyring(path if path.exists() else None)
    key = current_key()
    if key:
        return key
    raise DatabaseLockedError("Encrypted database is locked. Open Vertex and enter the password.")


@contextmanager
def _connect(*, create: bool = False) -> Iterator[Any]:
    path = db_path()
    if create:
        path.parent.mkdir(parents=True, exist_ok=True)
    elif not path.exists():
        raise FileNotFoundError(str(path))
    key = _require_key(path, creating=create)
    sqlite = get_sqlite()
    conn = sqlite.connect(str(path), timeout=30.0)
    try:
        if key:
            apply_key(conn, key)
        conn.execute("SELECT count(*) FROM sqlite_master").fetchone()
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = DELETE")
        conn.execute("PRAGMA synchronous = FULL")
        yield conn
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()


def _init_schema(conn: Any) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS clients (
            sort_idx INTEGER PRIMARY KEY,
            client_id TEXT,
            payload TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS account_managers (
            sort_idx INTEGER PRIMARY KEY,
            manager_id TEXT,
            payload TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS tasks (
            sort_idx INTEGER PRIMARY KEY,
            task_id TEXT,
            payload TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS monthly_state (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            payload TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS json_files (
            kind TEXT NOT NULL,
            name TEXT NOT NULL,
            payload TEXT NOT NULL,
            PRIMARY KEY (kind, name)
        );
        """
    )


def _dumps(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, default=str)


def _loads(text: str, default: Any = None) -> Any:
    try:
        return json.loads(text)
    except Exception:
        return default


def _replace_list(conn: Any, table: str, id_field: str, rows: list) -> int:
    if table not in _LIST_TABLES or _LIST_TABLES[table] != id_field:
        raise ValueError(f"unsupported table {table}")
    conn.execute(f"DELETE FROM {table}")
    count = 0
    for i, obj in enumerate(rows or []):
        if not isinstance(obj, dict):
            continue
        conn.execute(
            f"INSERT INTO {table} (sort_idx, {id_field}, payload) VALUES (?, ?, ?)",
            (i, str(obj.get("id") or ""), _dumps(obj)),
        )
        count += 1
    return count


def _load_list(conn: Any, table: str) -> list:
    rows = conn.execute(f"SELECT payload FROM {table} ORDER BY sort_idx").fetchall()
    out = []
    for (payload,) in rows:
        obj = _loads(payload)
        if isinstance(obj, dict):
            out.append(obj)
    return out


def load_clients_payloads() -> list:
    with _connect() as conn:
        return _load_list(conn, "clients")


def save_clients_payloads(clients: list) -> None:
    backup_db()
    with _connect(create=True) as conn:
        _init_schema(conn)
        _replace_list(conn, "clients", "client_id", clients)


def load_account_managers_payloads() -> list:
    with _connect() as conn:
        return _load_list(conn, "account_managers")


def save_account_managers_payloads(managers: list) -> None:
    backup_db()
    with _connect(create=True) as conn:
        _init_schema(conn)
        _replace_list(conn, "account_managers", "manager_id", managers)


def load_tasks_payloads() -> list:
    with _connect() as conn:
        return _load_list(conn, "tasks")


def save_tasks_payloads(tasks: list) -> None:
    backup_db()
    with _connect(create=True) as conn:
        _init_schema(conn)
        _replace_list(conn, "tasks", "task_id", tasks)


def load_monthly_state_payload() -> dict:
    with _connect() as conn:
        row = conn.execute("SELECT payload FROM monthly_state WHERE id = 1").fetchone()
    if not row:
        return {}
    data = _loads(row[0], default={})
    return data if isinstance(data, dict) else {}


def save_monthly_state_payload(blob: dict) -> None:
    backup_db()
    with _connect(create=True) as conn:
        _init_schema(conn)
        conn.execute(
            "INSERT OR REPLACE INTO monthly_state (id, payload) VALUES (1, ?)",
            (_dumps(blob if isinstance(blob, dict) else {}),),
        )


def load_json_file(kind: str, name: str) -> Any | None:
    with _connect() as conn:
        row = conn.execute(
            "SELECT payload FROM json_files WHERE kind = ? AND name = ?",
            (kind, name),
        ).fetchone()
    if not row:
        return None
    return _loads(row[0], default=None)


def save_json_file(kind: str, name: str, obj: Any) -> None:
    backup_db()
    with _connect(create=True) as conn:
        _init_schema(conn)
        conn.execute(
            """
            INSERT INTO json_files (kind, name, payload)
            VALUES (?, ?, ?)
            ON CONFLICT(kind, name) DO UPDATE SET payload = excluded.payload
            """,
            (kind, name, _dumps(obj)),
        )


def load_json_files(kind: str) -> dict[str, Any]:
    with _connect() as conn:
        rows = conn.execute(
            "SELECT name, payload FROM json_files WHERE kind = ? ORDER BY name",
            (kind,),
        ).fetchall()
    out: dict[str, Any] = {}
    for name, payload in rows:
        out[str(name)] = _loads(payload, default={})
    return out


def save_json_files(kind: str, files: dict[str, Any]) -> int:
    backup_db()
    with _connect(create=True) as conn:
        _init_schema(conn)
        conn.execute("DELETE FROM json_files WHERE kind = ?", (kind,))
        n = 0
        for name, obj in (files or {}).items():
            conn.execute(
                "INSERT INTO json_files (kind, name, payload) VALUES (?, ?, ?)",
                (kind, str(name), _dumps(obj)),
            )
            n += 1
        return n


def json_file_exists(kind: str, name: str) -> bool:
    with _connect() as conn:
        row = conn.execute(
            "SELECT 1 FROM json_files WHERE kind = ? AND name = ?",
            (kind, name),
        ).fetchone()
    return row is not None


def backup_db(*, keep: int | None = None, force: bool = False) -> Path | None:
    """
    Snapshot vertex.db into data/backups/vertex_YYYYMMDD_HHMMSS.db and keep
    the newest ``keep`` copies (same 7-copy policy as JSON). The snapshot is
    encrypted with the same SQLCipher key as the live database.
    """
    global _LAST_BACKUP_MONO
    if _SKIP_DB_BACKUP and not force:
        return None
    now = time.monotonic()
    if not force and (now - _LAST_BACKUP_MONO) < _BACKUP_DEBOUNCE_SEC:
        return None
    path = db_path()
    try:
        if not path.exists() or path.stat().st_size == 0:
            return None
        key = current_key()
        if not key and not is_plaintext_sqlite(path):
            try_unlock_from_keyring(path)
            key = current_key()
        keep_n = int(keep if keep is not None else getattr(_io(), "BACKUP_KEEP", BACKUP_KEEP) or BACKUP_KEEP)
        backups_dir = path.parent / "backups"
        backups_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = backups_dir / f"vertex_{ts}.db"
        n = 1
        while dest.exists():
            dest = backups_dir / f"vertex_{ts}_{n}.db"
            n += 1
        sqlite = get_sqlite()
        src = sqlite.connect(str(path), timeout=30.0)
        try:
            if key:
                apply_key(src, key)
            dst = sqlite.connect(str(dest), timeout=30.0)
            try:
                if key:
                    apply_key(dst, key)
                src.backup(dst)
                dst.commit()
            finally:
                dst.close()
        finally:
            src.close()
        if key and is_plaintext_sqlite(dest):
            try:
                dest.unlink()
            except Exception:
                pass
            raise RuntimeError("SQL backup was not encrypted")
        _LAST_BACKUP_MONO = now
        existing = sorted(
            backups_dir.glob("vertex_*.db"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        for old in existing[keep_n:]:
            try:
                old.unlink()
            except Exception:
                pass
        LOG.info("SQL backup: %s", dest)
        return dest
    except Exception as e:
        LOG.warning("SQL backup failed: %s", e)
        return None


def purge_plaintext_backups() -> int:
    """Remove unencrypted vertex_*.db copies left over from before SQLCipher."""
    path = db_path()
    n = purge_plaintext_db_copies(path.parent / "backups")
    n += purge_plaintext_db_copies(path.parent)
    leftover = path.with_name(path.name + ".plain.bak")
    if leftover.exists() and is_plaintext_sqlite(leftover):
        try:
            leftover.unlink()
            n += 1
        except Exception:
            pass
    if n:
        LOG.info("Removed %s plaintext database leftover(s)", n)
    return n


def change_db_password(old: str, new: str) -> None:
    """Re-key the live database. `old` must open it; `new` becomes the session key."""
    global _READY_CACHE
    path = db_path()
    if not path.exists():
        raise FileNotFoundError(str(path))
    if not verify_open(path, old):
        raise RuntimeError("Current password is incorrect.")
    sqlite = get_sqlite()
    conn = sqlite.connect(str(path), timeout=30.0)
    try:
        apply_key(conn, old)
        apply_rekey(conn, new)
        conn.execute("SELECT count(*) FROM sqlite_master").fetchone()
        conn.commit()
    finally:
        conn.close()
    set_session_key(new)
    _READY_CACHE = None
    if not verify_open(path, new):
        raise RuntimeError("Password was changed but the database could not be reopened.")
    backups_dir = path.parent / "backups"
    if backups_dir.is_dir():
        for old_bak in backups_dir.glob("vertex_*.db"):
            try:
                old_bak.unlink()
            except Exception:
                pass
    backup_db(force=True)
    LOG.info("Database password changed")


@contextmanager
def _no_db_backup() -> Iterator[None]:
    global _SKIP_DB_BACKUP
    prev = _SKIP_DB_BACKUP
    _SKIP_DB_BACKUP = True
    try:
        yield
    finally:
        _SKIP_DB_BACKUP = prev


def ensure_schema() -> None:
    global _READY_CACHE
    now = datetime.now().isoformat(timespec="seconds")
    with _connect(create=True) as conn:
        _init_schema(conn)
        conn.execute(
            "INSERT OR IGNORE INTO meta (key, value) VALUES (?, ?)",
            ("schema_version", str(SCHEMA_VERSION)),
        )
        conn.execute(
            "INSERT OR IGNORE INTO meta (key, value) VALUES (?, ?)",
            ("migrated_at", now),
        )
    _READY_CACHE = True


def sql_has_list(table: str) -> bool:
    """True when that SQL table already has at least one row."""
    if table not in _LIST_TABLES or not sql_db_ready():
        return False
    try:
        with _connect() as conn:
            row = conn.execute(f"SELECT 1 FROM {table} LIMIT 1").fetchone()
        return row is not None
    except Exception:
        return False


def sql_has_monthly_state() -> bool:
    if not sql_db_ready():
        return False
    try:
        with _connect() as conn:
            row = conn.execute("SELECT 1 FROM monthly_state WHERE id = 1 LIMIT 1").fetchone()
        return row is not None
    except Exception:
        return False


def sql_has_any_match_rules() -> bool:
    if not sql_db_ready():
        return False
    try:
        with _connect() as conn:
            row = conn.execute(
                "SELECT 1 FROM json_files WHERE kind = ? LIMIT 1",
                ("match_rules",),
            ).fetchone()
        return row is not None
    except Exception:
        return False


def _json_file_present(path: Path) -> bool:
    try:
        p = Path(path)
        return p.exists() and p.stat().st_size > 0
    except Exception:
        return False


def _json_still_needed() -> bool:
    """True when a live JSON file still holds data that SQL does not."""
    if not sql_db_ready():
        return True
    io = _io()

    def _list_pending(path: Path, table: str) -> bool:
        if sql_has_list(table) or not _json_file_present(path):
            return False
        raw = io._read_json_file(path, [])
        return isinstance(raw, list) and len(raw) > 0

    if _list_pending(io.DATA_FILE, "clients"):
        return True
    if _list_pending(io.ACCOUNT_MANAGERS_FILE, "account_managers"):
        return True
    if _list_pending(io.TASKS_FILE, "tasks"):
        return True
    if _json_file_present(io.MONTHLY_STATE_FILE) and not sql_has_monthly_state():
        raw = io._read_json_file(io.MONTHLY_STATE_FILE, {})
        if isinstance(raw, dict) and raw:
            return True
    if io.MATCH_RULES_DIR.exists() and any(io.MATCH_RULES_DIR.glob("*.json")) and not sql_has_any_match_rules():
        return True
    return False


def _retire_json_if_safe() -> None:
    if _json_still_needed():
        return
    try:
        stats = _io().retire_json_artifacts()
        if stats.get("removed"):
            LOG.info("Retired JSON artifacts: %s", stats)
    except Exception:
        LOG.exception("JSON retire failed")


def migrate_missing_json_to_sql() -> dict:
    """
    On startup: for each store, copy JSON → SQL only when the JSON file exists
    and that store has no SQL data yet. Never overwrites a store that already
    has SQL rows. After a successful copy, JSON files are removed.
    """
    global _MISSING_JSON_MIGRATED, _MISSING_JSON_MIGRATING
    empty = {
        "clients": 0,
        "account_managers": 0,
        "tasks": 0,
        "monthly_state": 0,
        "match_rules": 0,
    }
    if _MISSING_JSON_MIGRATED or _MISSING_JSON_MIGRATING:
        return empty
    _MISSING_JSON_MIGRATING = True
    try:
        path = db_path()
        if not path.exists() or path.stat().st_size < 64:
            return empty
        if not current_key() and not is_plaintext_sqlite(path):
            try_unlock_from_keyring(path)
            if not current_key():
                return empty

        io = _io()
        need_clients = not sql_has_list("clients") and _json_file_present(io.DATA_FILE)
        need_ams = not sql_has_list("account_managers") and _json_file_present(io.ACCOUNT_MANAGERS_FILE)
        need_tasks = not sql_has_list("tasks") and _json_file_present(io.TASKS_FILE)
        need_monthly = not sql_has_monthly_state() and _json_file_present(io.MONTHLY_STATE_FILE)
        need_rules = not sql_has_any_match_rules() and io.MATCH_RULES_DIR.exists()

        if not (need_clients or need_ams or need_tasks or need_monthly or need_rules):
            _MISSING_JSON_MIGRATED = True
            _retire_json_if_safe()
            return empty

        clients_json = io._read_json_file(io.DATA_FILE, []) if need_clients else None
        ams_json = io._read_json_file(io.ACCOUNT_MANAGERS_FILE, []) if need_ams else None
        tasks_json = io._read_json_file(io.TASKS_FILE, []) if need_tasks else None
        monthly_json = io._read_json_file(io.MONTHLY_STATE_FILE, {}) if need_monthly else None
        rule_files: dict[str, Any] = {}
        if need_rules:
            for p in io.MATCH_RULES_DIR.glob("*.json"):
                rule_files[p.name] = io._read_json_file(p, {})

        need_clients = isinstance(clients_json, list) and len(clients_json) > 0
        need_ams = isinstance(ams_json, list) and len(ams_json) > 0
        need_tasks = isinstance(tasks_json, list) and len(tasks_json) > 0
        need_monthly = isinstance(monthly_json, dict) and len(monthly_json) > 0
        need_rules = bool(rule_files)

        if not (need_clients or need_ams or need_tasks or need_monthly or need_rules):
            _MISSING_JSON_MIGRATED = True
            _retire_json_if_safe()
            return empty

        ensure_schema()
        stats = dict(empty)
        with _no_db_backup():
            if need_clients:
                save_clients_payloads(clients_json)
                stats["clients"] = len(clients_json)
            if need_ams:
                save_account_managers_payloads(ams_json)
                stats["account_managers"] = len(ams_json)
            if need_tasks:
                save_tasks_payloads(tasks_json)
                stats["tasks"] = len(tasks_json)
            if need_monthly:
                save_monthly_state_payload(monthly_json)
                stats["monthly_state"] = 1
            if need_rules:
                for name, obj in rule_files.items():
                    save_json_file("match_rules", name, obj)
                stats["match_rules"] = len(rule_files)
        backup_db(force=True)
        LOG.info("Startup JSON→SQL migrate (missing stores only): %s", stats)
        _MISSING_JSON_MIGRATED = True
        _retire_json_if_safe()
        return stats
    finally:
        _MISSING_JSON_MIGRATING = False


def migrate_json_to_sql(
    *,
    clients: list | None = None,
    account_managers: list | None = None,
    tasks: list | None = None,
    overwrite: bool = False,
) -> dict:
    """
    Copy current JSON (and optional in-memory lists) into an encrypted vertex.db.
    JSON files are removed after a successful copy.
    """
    if not current_key():
        try_unlock_from_keyring(db_path() if db_path().exists() else None)
    if not current_key():
        raise DatabaseLockedError("Set a database password before migrating.")

    io = _io()
    existed = sql_db_ready()
    if existed and not overwrite:
        raise FileExistsError(str(db_path()))
    if existed:
        backup_db(force=True)

    if clients is None:
        raw = io._read_json_file(io.DATA_FILE, [])
        clients = raw if isinstance(raw, list) else []
    if account_managers is None:
        raw = io._read_json_file(io.ACCOUNT_MANAGERS_FILE, [])
        account_managers = raw if isinstance(raw, list) else []
    if tasks is None:
        raw = io._read_json_file(io.TASKS_FILE, [])
        tasks = raw if isinstance(raw, list) else []

    if existed:
        monthly = load_monthly_state_payload()
        match_rules = load_json_files("match_rules")
        if not monthly:
            raw = io._read_json_file(io.MONTHLY_STATE_FILE, {})
            monthly = raw if isinstance(raw, dict) else {}
        if not match_rules and io.MATCH_RULES_DIR.exists():
            match_rules = {
                p.name: io._read_json_file(p, {})
                for p in io.MATCH_RULES_DIR.glob("*.json")
            }
    else:
        raw = io._read_json_file(io.MONTHLY_STATE_FILE, {})
        monthly = raw if isinstance(raw, dict) else {}
        match_rules = {}
        if io.MATCH_RULES_DIR.exists():
            match_rules = {
                p.name: io._read_json_file(p, {})
                for p in io.MATCH_RULES_DIR.glob("*.json")
            }

    now = datetime.now().isoformat(timespec="seconds")
    with _connect(create=True) as conn:
        _init_schema(conn)
        n_clients = _replace_list(conn, "clients", "client_id", clients)
        n_ams = _replace_list(conn, "account_managers", "manager_id", account_managers)
        n_tasks = _replace_list(conn, "tasks", "task_id", tasks)
        conn.execute(
            "INSERT OR REPLACE INTO monthly_state (id, payload) VALUES (1, ?)",
            (_dumps(monthly if isinstance(monthly, dict) else {}),),
        )
        conn.execute("DELETE FROM json_files WHERE kind = ?", ("match_rules",))
        n_rules = 0
        for name, obj in (match_rules or {}).items():
            conn.execute(
                "INSERT INTO json_files (kind, name, payload) VALUES (?, ?, ?)",
                ("match_rules", str(name), _dumps(obj)),
            )
            n_rules += 1
        conn.execute(
            "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
            ("schema_version", str(SCHEMA_VERSION)),
        )
        conn.execute(
            "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
            ("migrated_at", now),
        )

    global _READY_CACHE
    _READY_CACHE = True
    stats = {
        "db_path": str(db_path()),
        "clients": n_clients,
        "account_managers": n_ams,
        "tasks": n_tasks,
        "monthly_state": 1 if monthly else 0,
        "match_rules": n_rules,
        "overwritten": existed,
        "migrated_at": now,
    }
    LOG.info("Migrated JSON to SQL: %s", stats)
    backup_db(force=True)
    try:
        retired = io.retire_json_artifacts()
        if retired.get("removed"):
            LOG.info("Retired JSON artifacts: %s", retired)
        stats["json_retired"] = len(retired.get("removed") or [])
    except Exception:
        LOG.exception("JSON retire failed")
        stats["json_retired"] = 0
    return stats
