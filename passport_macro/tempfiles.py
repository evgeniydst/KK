"""Временные папки вырезок PDF/OCR: удаляются при выходе и при следующем запуске.

Раньше каждая загрузка PDF и каждое открытие проекта оставляли папку в %TEMP%.
Вырезки нужны, пока программа открыта; в сохранённом проекте они встроены
в файл .kkproject, поэтому после выхода их можно удалять.
"""
from __future__ import annotations

import atexit
import logging
import shutil
import tempfile
import threading
import time
from pathlib import Path

log = logging.getLogger(__name__)

PREFIXES = ('pdf_checklist_evidence_', 'kk_project_crops_')
_lock = threading.Lock()
_created: list[Path] = []


def make_temp_dir(prefix: str) -> Path:
    """Создать временную папку и запомнить её для удаления при выходе."""
    if prefix not in PREFIXES:
        raise ValueError('Неизвестный префикс временной папки: ' + prefix)
    path = Path(tempfile.mkdtemp(prefix=prefix))
    with _lock:
        _created.append(path)
    return path


def cleanup_created() -> None:
    with _lock:
        paths, _created[:] = list(_created), []
    for path in paths:
        shutil.rmtree(path, ignore_errors=True)


def cleanup_stale(max_age_hours: float = 48) -> int:
    """Удалить старые папки прошлых запусков (например, после аварийного закрытия).

    Порог по возрасту не трогает папки другой одновременно открытой копии программы.
    """
    limit = time.time() - max_age_hours * 3600
    removed = 0
    root = Path(tempfile.gettempdir())
    for prefix in PREFIXES:
        for path in root.glob(prefix + '*'):
            try:
                if path.is_dir() and path.stat().st_mtime < limit:
                    shutil.rmtree(path, ignore_errors=True)
                    removed += 1
            except OSError as exc:
                log.debug('Не удалось проверить %s: %s', path, exc)
    if removed:
        log.info('Удалено старых временных папок: %d', removed)
    return removed


atexit.register(cleanup_created)
