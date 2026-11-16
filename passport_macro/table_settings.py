"""Размеры таблицы КК и личные значения по умолчанию."""
from __future__ import annotations

import copy
import json
import logging
import math
import os
import tempfile
from pathlib import Path

from .paths import user_data_dir

log = logging.getLogger(__name__)
TOTAL_WIDTH_MM = 134.5
BUILTIN_TABLE_LAYOUT = {
    'widths': [9.0, 65.0, 17.0, 14.0, 29.5],
    'header_min_mm': 5.0,
    'row_min_mm': 5.0,
}


def validate_table_layout(value: object) -> dict:
    if not isinstance(value, dict):
        raise ValueError('Некорректные настройки таблицы.')
    widths = value.get('widths')
    if not isinstance(widths, list) or len(widths) != 5:
        raise ValueError('Укажите ширину всех пяти столбцов.')
    def number(raw, label):
        if isinstance(raw, bool) or not isinstance(raw, (int, float)) or not math.isfinite(raw):
            raise ValueError(label + ': введите число в миллиметрах.')
        if abs(raw * 100 - round(raw * 100)) > 0.000001:
            raise ValueError(label + ': допускается не более двух знаков после запятой.')
        return round(float(raw), 2)
    widths = [number(raw, f'Столбец {i + 1}') for i, raw in enumerate(widths)]
    if any(w < 5 for w in widths):
        raise ValueError('Ширина каждого столбца должна быть не меньше 5 мм.')
    if abs(sum(widths) - TOTAL_WIDTH_MM) > 0.001:
        raise ValueError(f'Общая ширина должна быть {TOTAL_WIDTH_MM:g} мм; сейчас {sum(widths):g} мм.')
    header = number(value.get('header_min_mm'), 'Высота шапки')
    row = number(value.get('row_min_mm'), 'Высота строки данных')
    if not (3 <= header <= 40 and 3 <= row <= 40):
        raise ValueError('Минимальная высота строк должна быть от 3 до 40 мм.')
    return {'widths': widths, 'header_min_mm': header, 'row_min_mm': row}


def table_defaults_path() -> Path:
    return user_data_dir() / 'table_layout_defaults.json'


def load_table_defaults(path: str | Path | None = None) -> dict:
    try:
        return validate_table_layout(json.loads(Path(path or table_defaults_path()).read_text(encoding='utf-8-sig')))
    except FileNotFoundError:
        return copy.deepcopy(BUILTIN_TABLE_LAYOUT)
    except (OSError, ValueError, TypeError) as exc:
        log.warning('Не удалось прочитать настройки таблицы: %s; используются встроенные.', exc)
        return copy.deepcopy(BUILTIN_TABLE_LAYOUT)


def save_table_defaults(value: dict, path: str | Path | None = None) -> None:
    layout = validate_table_layout(value)
    target = Path(path or table_defaults_path())
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=target.parent,
                                         prefix='.table_layout_', suffix='.tmp', delete=False) as f:
            temp = Path(f.name)
            json.dump(layout, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp, target)
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)