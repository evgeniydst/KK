"""Журнал работы: файл с ротацией и перехват необработанных ошибок.

При запуске через pythonw.exe консоли нет, поэтому без журнала трассировки
ошибок терялись. Теперь они пишутся в %LOCALAPPDATA%\\KKPassport\\passport.log.
"""
from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .paths import user_data_dir

LOG_NAME = 'passport.log'
_log_path: Path | None = None


def log_path() -> Path | None:
    """Путь к журналу или None, если файл создать не удалось."""
    return _log_path


def setup_logging(directory: Path | None = None, level: int = logging.INFO) -> Path | None:
    global _log_path
    root = logging.getLogger()
    if any(getattr(handler, '_kk_passport', False) for handler in root.handlers):
        return _log_path
    root.setLevel(level)
    formatter = logging.Formatter('%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s')
    try:
        directory = Path(directory) if directory else user_data_dir()
        directory.mkdir(parents=True, exist_ok=True)
        handler: logging.Handler = RotatingFileHandler(
            directory / LOG_NAME, maxBytes=1_000_000, backupCount=3, encoding='utf-8')
        _log_path = directory / LOG_NAME
    except OSError:
        # В оконном .exe нет консоли (sys.stderr is None).
        handler = logging.StreamHandler() if sys.stderr is not None else logging.NullHandler()
        _log_path = None
    handler.setFormatter(formatter)
    handler._kk_passport = True
    root.addHandler(handler)

    def excepthook(exc_type, exc, tb):
        logging.getLogger('passport_macro').critical('Необработанная ошибка', exc_info=(exc_type, exc, tb))
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = excepthook
    return _log_path
