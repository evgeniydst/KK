"""Список недавних проектов для начального экрана."""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path

from .paths import user_data_dir

log = logging.getLogger(__name__)

RECENT_LIMIT = 8


def recent_file() -> Path:
    return user_data_dir() / 'recent.json'


def _key(path) -> str:
    return os.path.normcase(os.path.abspath(str(path)))


def _read(file: Path) -> list[str]:
    try:
        data = json.loads(file.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return []
    if not isinstance(data, list):
        return []
    result, seen = [], set()
    for value in data:
        if isinstance(value, str) and value and _key(value) not in seen:
            seen.add(_key(value))
            result.append(value)
    return result


def _write(file: Path, paths: list[str]) -> None:
    try:
        file.parent.mkdir(parents=True, exist_ok=True)
        temp = file.with_suffix('.tmp')
        temp.write_text(json.dumps(paths[:RECENT_LIMIT], ensure_ascii=False, indent=1),
                        encoding='utf-8')
        os.replace(temp, file)
    except OSError as exc:  # список недавних не должен мешать работе
        log.info('Не удалось сохранить список недавних проектов: %s', exc)


def load_recent(file: Path | None = None) -> list[str]:
    """Недавние проекты (новые первыми), только существующие файлы."""
    return [path for path in _read(file or recent_file()) if Path(path).is_file()]


def add_recent(project, file: Path | None = None) -> None:
    file = file or recent_file()
    project = str(Path(project).resolve())
    paths = [project] + [path for path in _read(file) if _key(path) != _key(project)]
    _write(file, paths)


def forget_recent(project, file: Path | None = None) -> None:
    file = file or recent_file()
    _write(file, [path for path in _read(file) if _key(path) != _key(project)])
