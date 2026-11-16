"""Автосохранение проекта и восстановление после сбоя или закрытия без сохранения.

Каждое окно программы пишет свою копию в %LOCALAPPDATA%\\KKPassport\\autosave:
  <id>.kkproject — проект (тот же формат, что и обычный файл проекта),
  <id>.json      — сведения для вопроса о восстановлении,
  <id>.lock      — заблокирован, пока окно открыто.
Если при запуске найдена копия с незаблокированным .lock, значит прошлое окно
закрылось аварийно — программа предложит восстановить проект. Копия другого
открытого окна (.lock заблокирован) не трогается.

Проект, закрытый кнопкой «Нет» (без сохранения), не удаляется сразу: последняя
такая копия хранится как last_closed и открывается из меню «Проект».
"""
from __future__ import annotations

import json
import logging
import os
import tempfile
import time
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .paths import user_data_dir
from .project import read_project_file, write_project_file

log = logging.getLogger(__name__)

AUTOSAVE_SECONDS = 60
LAST_CLOSED = 'last_closed'
PROJECT_SUFFIX = '.kkproject'


def autosave_dir() -> Path:
    return user_data_dir() / 'autosave'


def _try_lock(stream) -> bool:
    """Неблокирующая блокировка первого байта файла (Windows и Linux)."""
    try:
        stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except OSError:
        return False


def _unlock(stream) -> None:
    try:
        stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_UN)
    except OSError:
        log.debug('Не удалось снять блокировку автосохранения', exc_info=True)


def _session_alive(lock_file: Path) -> bool:
    """True, если окно-владелец копии ещё открыто (его .lock заблокирован)."""
    if not lock_file.exists():
        return False
    try:
        with open(lock_file, 'r+b') as stream:
            if _try_lock(stream):
                _unlock(stream)
                return False
            return True
    except OSError:
        return True  # нет доступа — на всякий случай считаем занятым


def _write_json(path: Path, value: dict) -> None:
    temporary = None
    try:
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=path.parent, prefix='.autosave_',
                                         suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(value, stream, ensure_ascii=False, indent=2)
        os.replace(temporary, path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def _remove(*paths: Path) -> None:
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            log.debug('Не удалось удалить %s', path, exc_info=True)


@dataclass
class RecoveryEntry:
    """Найденная копия проекта."""
    id: str
    project_file: Path
    meta_file: Path
    lock_file: Path
    project_path: str = ''
    name: str = ''
    saved_at: str = ''
    cards: int = 0

    def describe(self) -> str:
        """Строка для диалога: «Рама ТМ10…», КК: 3, 09.11.2026 12:58."""
        try:
            moment = datetime.fromisoformat(self.saved_at).strftime('%d.%m.%Y %H:%M')
        except ValueError:
            moment = 'время неизвестно'
        text = f'«{self.name or "Новый проект"}», КК: {self.cards}, сохранено {moment}'
        if self.project_path:
            text += f'\nФайл проекта: {self.project_path}'
        return text

    def load(self):
        """(cards, settings) из копии; ошибки — как у read_project_file."""
        return read_project_file(self.project_file)


def _entry(folder: Path, entry_id: str) -> RecoveryEntry:
    entry = RecoveryEntry(entry_id, folder / (entry_id + PROJECT_SUFFIX),
                          folder / (entry_id + '.json'), folder / (entry_id + '.lock'))
    try:
        meta = json.loads(entry.meta_file.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        meta = {}
    if not isinstance(meta, dict):
        meta = {}
    entry.project_path = str(meta.get('project_path') or '')
    entry.name = str(meta.get('name') or '')
    entry.saved_at = str(meta.get('saved_at') or '')
    if not entry.saved_at and entry.project_file.exists():
        entry.saved_at = datetime.fromtimestamp(entry.project_file.stat().st_mtime).isoformat(timespec='seconds')
    cards = meta.get('cards', 0)
    entry.cards = cards if isinstance(cards, int) else 0
    return entry


def find_recoverable(folder: str | Path | None = None) -> list[RecoveryEntry]:
    """Копии окон, закрытых аварийно (новые первыми). Хвосты без проекта убираются."""
    folder = Path(folder) if folder else autosave_dir()
    if not folder.is_dir():
        return []
    entries = []
    ids = {path.stem for pattern in ('*' + PROJECT_SUFFIX, '*.json', '*.lock')
           for path in folder.glob(pattern)}
    for entry_id in sorted(ids - {LAST_CLOSED}):
        entry = _entry(folder, entry_id)
        if _session_alive(entry.lock_file):
            continue
        if not entry.project_file.is_file():
            _remove(entry.meta_file, entry.lock_file)  # окно закрылось без несохранённых правок
            continue
        entries.append(entry)
    entries.sort(key=lambda item: item.saved_at, reverse=True)
    return entries


def last_closed_entry(folder: str | Path | None = None) -> RecoveryEntry | None:
    folder = Path(folder) if folder else autosave_dir()
    entry = _entry(folder, LAST_CLOSED)
    return entry if entry.project_file.is_file() else None


def set_aside(entry: RecoveryEntry) -> None:
    """Отказ от восстановления: копия становится «последним несохранённым» (меню «Проект»)."""
    folder = entry.project_file.parent
    os.replace(entry.project_file, folder / (LAST_CLOSED + PROJECT_SUFFIX))
    if entry.meta_file.exists():
        os.replace(entry.meta_file, folder / (LAST_CLOSED + '.json'))
    _remove(entry.lock_file)


def remove_entry(entry: RecoveryEntry) -> None:
    _remove(entry.project_file, entry.meta_file, entry.lock_file)


def mark_broken(entry: RecoveryEntry) -> None:
    """Повреждённая копия больше не предлагается, но остаётся на диске для ручного разбора."""
    try:
        os.replace(entry.project_file, entry.project_file.with_name(entry.project_file.name + '.broken'))
    except OSError:
        log.warning('Не удалось отложить повреждённую копию %s', entry.project_file, exc_info=True)
    _remove(entry.meta_file, entry.lock_file)


class AutosaveSession:
    """Автосохранение одного окна. Если папка недоступна, enabled = False и окно работает как раньше."""

    def __init__(self, folder: str | Path | None = None):
        self.folder = Path(folder) if folder else autosave_dir()
        self.id = time.strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:8]
        self.enabled = False
        self.last_payload = None
        self.saved_at: datetime | None = None
        self._lock_stream = None
        try:
            self.folder.mkdir(parents=True, exist_ok=True)
            stream = open(self._path('.lock'), 'w+b')
            stream.write(b'1')
            stream.flush()
            if not _try_lock(stream):
                stream.close()
                raise OSError('Файл блокировки автосохранения занят.')
            self._lock_stream = stream
            self.enabled = True
        except OSError:
            log.warning('Автосохранение недоступно: %s', self.folder, exc_info=True)

    def _path(self, suffix: str) -> Path:
        return self.folder / (self.id + suffix)

    @property
    def project_file(self) -> Path:
        return self._path(PROJECT_SUFFIX)

    def _write(self, stem: str, cards, settings, project_path: str) -> None:
        write_project_file(self.folder / (stem + PROJECT_SUFFIX), cards, settings)
        _write_json(self.folder / (stem + '.json'), {
            'version': 1, 'project_path': str(project_path or ''),
            'name': str(settings.get('name', '')), 'cards': len(cards),
            'saved_at': datetime.now().isoformat(timespec='seconds'),
        })

    def save(self, cards, settings, project_path: str = '', payload=None) -> bool:
        """Записать копию. payload — снимок project_payload: без изменений файл не переписывается."""
        if not self.enabled or (payload is not None and payload == self.last_payload):
            return False
        self._write(self.id, cards, settings, project_path)
        self.last_payload = payload
        self.saved_at = datetime.now()
        return True

    def discard(self) -> None:
        """Несохранённых правок нет — копия не нужна."""
        if self.enabled:
            _remove(self.project_file, self._path('.json'))
        self.last_payload = None
        self.saved_at = None

    def keep_last_closed(self, cards, settings, project_path: str = '') -> bool:
        """Пользователь отказался сохранять: оставить одну копию на случай ошибки."""
        if not self.enabled or not cards:
            return False
        self._write(LAST_CLOSED, cards, settings, project_path)
        return True

    def close(self) -> None:
        """Штатное закрытие окна: копия и блокировка удаляются."""
        self.discard()
        if self._lock_stream is not None:
            _unlock(self._lock_stream)
            self._lock_stream.close()
            self._lock_stream = None
            _remove(self._path('.lock'))
        self.enabled = False
