import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from config import DEFAULT_LOG_LEVEL

LOG_FILE = Path(__file__).with_name("automation.log")
_LOG_INITIALIZED = False

def configure_logging(level: str = DEFAULT_LOG_LEVEL) -> None:
    """Configure application-wide logging.

    Parameters
    ----------
    level: str
        Logging level name such as ``"INFO"`` or ``"DEBUG"``.
    """
    global _LOG_INITIALIZED
    root = logging.getLogger()
    # Remove previous handlers to allow reconfiguration
    for handler in root.handlers[:]:
        root.removeHandler(handler)
        handler.close()
    LOG_FILE.unlink(missing_ok=True)
    handler = RotatingFileHandler(
        LOG_FILE, maxBytes=1_000_000, backupCount=5, encoding="utf-8"
    )
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(name)s - %(message)s"
    )
    handler.setFormatter(formatter)
    root.addHandler(handler)
    root.setLevel(getattr(logging, level.upper(), logging.INFO))
    _LOG_INITIALIZED = True

def get_logger(name: str) -> logging.Logger:
    """Return a module-level logger."""
    if not _LOG_INITIALIZED:
        configure_logging()
    return logging.getLogger(name)
