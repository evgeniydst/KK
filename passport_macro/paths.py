"""Пути программы: корень (папка запускающего файла или .exe), ресурсы, журнал."""
from __future__ import annotations

import os
import sys
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = PACKAGE_DIR / 'assets'

# Единая подсказка установки зависимостей (Windows, лаунчер py).
PIP_INSTALL = 'py -m pip install'


def app_dir() -> Path:
    """Папка запускающего файла; для сборки PyInstaller — папка .exe.

    Модули лежат в пакете, поэтому Path(__file__).parent здесь уже НЕ корень
    программы: tessdata и config.json ищутся относительно этой функции.
    """
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return PACKAGE_DIR.parent


def bundle_dir() -> Path:
    """Распакованные ресурсы PyInstaller (_MEIPASS), иначе корень программы."""
    return Path(getattr(sys, '_MEIPASS', app_dir()))


def user_data_dir() -> Path:
    """Записываемая папка пользователя для журнала (Program Files может быть только для чтения)."""
    base = os.environ.get('LOCALAPPDATA') or os.environ.get('XDG_STATE_HOME')
    root = Path(base) if base else Path.home() / '.local' / 'state'
    return root / 'KKPassport'
