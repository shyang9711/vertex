from __future__ import annotations
import logging, os
from logging.handlers import RotatingFileHandler
from pathlib import Path

_DEFAULT_LOG_NAME = "client_manager"

def _logs_dir() -> Path:
    """
    Prefer a logs/ folder next to client_manager.py.
    Fallback: ./logs
    """
    here = Path(__file__).resolve()
    # try <project_root>/logs
    root = here
    for _ in range(5):
        if (root / "client_manager.py").exists():
            break
        if root.parent == root:
            break
        root = root.parent
    logs = root / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    return logs

def get_logger(name: str = _DEFAULT_LOG_NAME, *, level: int = logging.INFO) -> logging.Logger:
    """
    Returns a module/page-specific logger with:
    - Rotating file logs: logs/app.log (1MB x 5 backups)
    - Console handler (only once per process)
    """
    logger = logging.getLogger(name)
    if getattr(logger, "_app_configured", False):
        return logger

    logger.setLevel(level)

    log_file = _logs_dir() / "app.log"
    fh = RotatingFileHandler(log_file, maxBytes=1_000_000, backupCount=5, encoding="utf-8")
    fh.setLevel(level)
    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Console once
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        ch = logging.StreamHandler()
        ch.setLevel(level)
        ch.setFormatter(fmt)
        logger.addHandler(ch)

    logger._app_configured = True  # type: ignore[attr-defined]
    return logger
