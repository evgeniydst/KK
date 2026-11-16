"""Значения титульного листа по умолчанию: config.json рядом с программой.

ФИО и должности меняются в config.json без правки кода. Если файла нет или он
повреждён, используются встроенные должности с пустыми ФИО (в Word остаётся
линия для ручного заполнения), а причина пишется в журнал.
"""
from __future__ import annotations

import copy
import json
import logging
from functools import lru_cache
from pathlib import Path

from .paths import app_dir

log = logging.getLogger(__name__)

CONFIG_NAME = 'config.json'
TEXT_KEYS = ('approval_role', 'approval_name', 'developer_role', 'developer_name')
BUILTIN_COVER_DEFAULTS: dict = {
    'approval_role': 'главный технолог', 'approval_name': '',
    'developer_role': 'технолог сварки', 'developer_name': '',
    'reviewers': [{'role': 'конструктор', 'name': ''},
                  {'role': 'начальник ОТК', 'name': ''}],
}


def config_path() -> Path:
    return app_dir() / CONFIG_NAME


def parse_cover_defaults(data: object) -> dict:
    """Взять из JSON только корректные поля; остальное — встроенные значения."""
    result = copy.deepcopy(BUILTIN_COVER_DEFAULTS)
    section = data.get('cover_defaults', data) if isinstance(data, dict) else None
    if not isinstance(section, dict):
        log.warning('config.json: ожидается объект "cover_defaults"; используются встроенные значения.')
        return result
    for key in TEXT_KEYS:
        if key not in section:
            continue
        value = section[key]
        if isinstance(value, str) and len(value) <= 250:
            result[key] = value.strip()
        else:
            log.warning('config.json: поле %s пропущено (нужна строка до 250 символов).', key)
    reviewers = section.get('reviewers')
    if reviewers is not None:
        valid = (isinstance(reviewers, list) and len(reviewers) <= 20 and all(
            isinstance(person, dict)
            and all(isinstance(person.get(k, ''), str) and len(person.get(k, '')) <= 250
                    for k in ('role', 'name'))
            for person in reviewers))
        if valid:
            result['reviewers'] = [{k: person.get(k, '').strip() for k in ('role', 'name')}
                                   for person in reviewers]
        else:
            log.warning('config.json: список reviewers пропущен (до 20 записей {"role", "name"}).')
    return result


@lru_cache(maxsize=None)
def _cached_cover_defaults(path: str) -> dict:
    try:
        data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    except FileNotFoundError:
        log.info('config.json не найден (%s); используются встроенные значения.', path)
        return copy.deepcopy(BUILTIN_COVER_DEFAULTS)
    except (OSError, ValueError) as exc:
        log.warning('config.json не прочитан (%s): %s; используются встроенные значения.', path, exc)
        return copy.deepcopy(BUILTIN_COVER_DEFAULTS)
    return parse_cover_defaults(data)


def load_cover_defaults(path: str | Path | None = None) -> dict:
    """Файл читается один раз за запуск; вызывающий получает независимую копию."""
    return copy.deepcopy(_cached_cover_defaults(str(path or config_path())))
