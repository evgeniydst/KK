"""Проект: настройки титула, сохранение и открытие файла проекта."""
from __future__ import annotations

import base64
import copy
import datetime
import json
import os
import re
import tempfile
import textwrap
import uuid
from dataclasses import asdict
from pathlib import Path
from typing import TypedDict

from .config import load_cover_defaults
from .models import ControlCard, Item
from .tempfiles import make_temp_dir
from .table_settings import BUILTIN_TABLE_LAYOUT, load_table_defaults, validate_table_layout


class Person(TypedDict):
    role: str
    name: str


class ProjectSettings(TypedDict):
    name: str
    passport_name: str
    approval_role: str
    approval_name: str
    developer_role: str
    developer_name: str
    types_enabled: bool
    types_title: str
    types: list[str]
    reviewers: list[Person]
    word_layout_mode: str     # ключ WORD_LAYOUT_MODES
    word_layout_check: bool   # проверять переносы в Microsoft Word (Windows)
    pdf_export: bool          # сохранять PDF рядом с DOCX
    cover_image: str          # PNG в base64; хранится внутри проекта
    cover_dates_enabled: bool
    approval_date: str
    developer_date: str
    table_layout: dict


def passport_filename(name: str) -> str:
    """Only the suggested DOCX filename gains the prefix; project name is unchanged."""
    if not re.match(r'^паспорт(?:\s|$)', name.strip(), re.IGNORECASE):
        name = 'Паспорт ' + name
    return project_filename(name, '.docx')


WORD_LAYOUT_MODES = {
    "free": "Свободный перенос",
    "whole": "Таблицы целиком",
}


def default_project_settings() -> ProjectSettings:
    """Пустой проект. Должности и ФИО для титула берутся из config.json."""
    cover = load_cover_defaults()
    return {
        'name': '', 'passport_name': '',
        'approval_role': cover['approval_role'], 'approval_name': cover['approval_name'],
        'developer_role': cover['developer_role'], 'developer_name': cover['developer_name'],
        'types_enabled': False, 'types_title': '', 'types': [],
        'reviewers': cover['reviewers'],
        'word_layout_mode': 'free',
        'word_layout_check': False,
        'pdf_export': False,
        'cover_image': '',
        'cover_dates_enabled': False,
        'approval_date': '',
        'developer_date': '',
        'table_layout': load_table_defaults(),
    }


def validate_cover_date(value: object, enabled: bool, label: str) -> str:
    if not isinstance(value, str) or len(value) > 32:
        raise ValueError('Некорректное поле даты: ' + label)
    value = value.strip()
    if enabled and value:
        try:
            if not re.fullmatch(r'\d{2}\.\d{2}\.\d{4}', value):
                raise ValueError
            datetime.datetime.strptime(value, '%d.%m.%Y')
        except ValueError as exc:
            raise ValueError(f'{label}: введите существующую дату в формате ДД.ММ.ГГГГ.') from exc
    return value


def validate_project_settings(value: object) -> ProjectSettings:
    if not isinstance(value, dict):
        raise ValueError('Некорректные настройки проекта.')
    result = default_project_settings()
    # Old projects keep the layout they were created with, even if the user
    # has since saved different defaults for new projects.
    result['table_layout'] = validate_table_layout(value.get('table_layout', BUILTIN_TABLE_LAYOUT))
    mode = value.get('word_layout_mode', 'free')
    if not isinstance(mode, str) or mode not in WORD_LAYOUT_MODES:
        raise ValueError('Неизвестный режим переноса таблиц Word.')
    result['word_layout_mode'] = mode
    check = value.get('word_layout_check', False)
    if not isinstance(check, bool):
        raise ValueError('Настройка проверки Word должна быть включена или выключена.')
    result['word_layout_check'] = check
    pdf_export = value.get('pdf_export', False)
    if not isinstance(pdf_export, bool):
        raise ValueError('Настройка сохранения PDF должна быть включена или выключена.')
    result['pdf_export'] = pdf_export
    image = value.get('cover_image', '')
    # Полная проверка при открытии проекта / обработке изображения; здесь только тип и предел.
    if not isinstance(image, str) or len(image) > 16 * 1024 * 1024:
        raise ValueError('Некорректное изображение титульного листа.')
    result['cover_image'] = image
    dates_enabled = value.get('cover_dates_enabled', False)
    if not isinstance(dates_enabled, bool):
        raise ValueError('Настройка дат титульного листа должна быть включена или выключена.')
    result['cover_dates_enabled'] = dates_enabled
    for key, label in (('approval_date', 'Дата утверждающего'),
                       ('developer_date', 'Дата разработчика')):
        result[key] = validate_cover_date(value.get(key, ''), dates_enabled, label)
    for key in ('name', 'passport_name', 'approval_role', 'approval_name', 'developer_role', 'developer_name', 'types_title'):
        text = value.get(key, result[key])
        if not isinstance(text, str) or len(text) > 1000:
            raise ValueError('Некорректное поле титульника: ' + key)
        result[key] = text.strip()
    reviewers = value.get('reviewers', result['reviewers'])
    if not isinstance(reviewers, list) or len(reviewers) > 20:
        raise ValueError('Допускается не более 20 согласующих.')
    result['reviewers'] = []
    for person in reviewers:
        if not isinstance(person, dict) or any(not isinstance(person.get(k, ''), str)
                                               or len(person.get(k, '')) > 250 for k in ('role', 'name')):
            raise ValueError('Некорректные данные согласующего.')
        clean_person = {k: person.get(k, '').strip() for k in ('role', 'name')}
        person_date = validate_cover_date(person.get('date', ''), dates_enabled, 'Дата согласующего')
        if person_date:
            clean_person['date'] = person_date
        result['reviewers'].append(clean_person)
    enabled = value.get('types_enabled', False)
    variants = value.get('types', [])
    if not isinstance(enabled, bool):
        raise ValueError('Некорректное состояние блока типов изделия.')
    if not isinstance(variants, list) or len(variants) > 1000 or any(
            not isinstance(v, str) or len(v) > 250 for v in variants):
        raise ValueError('Некорректный список типов: не более 1000 записей, по 250 символов.')
    result['types_enabled'] = enabled
    result['types'] = [v.strip() for v in variants]
    if enabled and (not result['types_title'] or not variants or any(not v for v in result['types'])):
        raise ValueError('Для блока типов заполните заголовок и хотя бы одно обозначение; уберите пустые строки.')
    return result


def cover_types_layout(settings: dict):
    """Estimate wrapped type rows; spend additional height on the picture only.

    Word performs final line breaking. Reserve conservative character counts
    for the 35/65 mm columns at GOST 14 pt, without shrinking text or squares.
    """
    def lines(text, width):
        return sum(max(1, len(textwrap.wrap(line, width=width, break_on_hyphens=False)))
                   for line in text.split('\n'))
    enabled = settings.get('types_enabled', False)
    # Above four positions use compact spacing, not smaller text.
    row_unit = 6 if len(settings.get('types', [])) > 4 else 7.5
    heights = ([row_unit * max(lines(text, 20), lines(settings['types_title'], 11) if index == 0 else 1)
                for index, text in enumerate(settings['types'])] if enabled else [])
    types_height = sum(heights) + 2 if enabled else 0
    caption = 'ПАСПОРТ ' + (settings.get('passport_name') or settings.get('name', ''))
    # The 22 pt caption wraps earlier than the 14 pt list. Account for it
    # even when the executions block is disabled.
    heading_extra = 9.5 * max(0, lines(caption, 38) - 1)
    heading_extra += 6.5 * max(0, lines('Утвердил: ' + settings.get('approval_role', ''), 26) - 2)
    heading_extra += 6.5 * max(0, lines(settings.get('approval_name', ''), 26) - 1)
    title_gap = max(2, 12 - max(0, types_height - 16))
    # Fixed lower edge: let the picture use the space previously left empty.
    # Keep a small safety gap after the serial number / executions block.
    picture_height = 131 - types_height - (title_gap - 2) - heading_extra
    if not enabled:
        picture_height += 2  # no trailing executions-table paragraph
    if picture_height < 10:
        raise ValueError(
            'Блок типов слишком большой для одного титульного листа: '
            f'расчётная высота списка {types_height:.1f} мм, '
            'даже при минимальной области рисунка 10 мм места недостаточно. '
            'Сократите заголовок/длинные обозначения или уменьшите список. '
            'Ни одна позиция не будет обрезана.')
    return {'types_height': types_height, 'picture_height': picture_height,
            'row_heights': heights}


def project_payload(cards: list[ControlCard], settings: dict, embed_images: bool = False) -> dict:
    """Plain versioned JSON; images are embedded only for disk persistence."""
    value = {'format': 'kk-passport-project', 'version': 1,
             'settings': validate_project_settings(settings), 'cards': []}
    crops = set()
    for card in cards:
        details = {'control_card': False, 'warnings': [], 'page_types': []}
        details.update({k: copy.deepcopy(v) for k, v in card.details.items() if k != 'graphic_rows'})
        details['graphic_rows'] = [[page, number, warning, crop]
                                  for (page, number), (warning, crop)
                                  in card.details.get('graphic_rows', {}).items()]
        record = {key: getattr(card, key) for key in ('id', 'source', 'document_code', 'document_name', 'raw')}
        record.update(items=[asdict(item) for item in card.items], details=details)
        value['cards'].append(record)
        crops.update(item.source_crop for item in card.items if item.source_crop)
        crops.update(path for item in card.items for path in item.ocr_crops if path)
        crops.update(row[3] for row in details['graphic_rows'] if row[3])
        crops.update(row['crop'] for row in details.get('ocr_rows', []) if row.get('crop'))
    if embed_images:
        value['images'] = {path: base64.b64encode(Path(path).read_bytes()).decode('ascii')
                           for path in sorted(crops) if Path(path).is_file()}
    return value


def write_project_file(path: str | Path, cards: list[ControlCard], settings: dict) -> None:
    """Atomic replacement: a failed write does not destroy the existing project."""
    payload = project_payload(cards, settings, embed_images=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if len(text.encode('utf-8')) > 64 * 1024 * 1024:
        raise ValueError('Проект больше 64 МБ. Разделите его на несколько проектов.')
    destination = Path(path)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=destination.parent,
                                         prefix='.kkproject_', suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def read_project_file(path: str | Path):
    """Validate before replacing GUI state. Stored source paths are never opened."""
    if Path(path).stat().st_size > 64 * 1024 * 1024:
        raise ValueError('Файл проекта больше 64 МБ.')
    try:
        value = json.loads(Path(path).read_text(encoding='utf-8-sig'))
        if (not isinstance(value, dict) or value.get('format') != 'kk-passport-project'
                or type(value.get('version')) is not int or value['version'] != 1):
            raise ValueError('Неизвестный формат или версия проекта.')
        settings = validate_project_settings(value.get('settings'))
        if settings['cover_image']:
            from .cover_image import decode_image
            decode_image(settings['cover_image'])
        records = value.get('cards')
        if not isinstance(records, list) or len(records) > 1000:
            raise ValueError('Некорректный список КК.')
        def text(record, key, default=''):
            result = record.get(key, default)
            if not isinstance(result, str) or len(result) > 20 * 1024 * 1024:
                raise ValueError('Некорректное текстовое поле: ' + key)
            return result
        cards = []
        for record in records:
            if not isinstance(record, dict) or not isinstance(record.get('items'), list):
                raise ValueError('Некорректная карта контроля.')
            items = []
            for entry in record['items']:
                if (not isinstance(entry, dict) or type(entry.get('checked', False)) is not bool
                        or type(entry.get('manual', False)) is not bool):
                    raise ValueError('Некорректный пункт.')
                pages = entry.get('pages', [])
                warnings = entry.get('warnings', [])
                ocr_crops = entry.get('ocr_crops', [])
                if (not isinstance(pages, list) or any(type(p) is not int or p < 1 for p in pages)
                        or not isinstance(warnings, list) or any(not isinstance(w, str) for w in warnings)
                        or not isinstance(ocr_crops, list) or len(ocr_crops) > 1000
                        or any(not isinstance(p, str) for p in ocr_crops)):
                    raise ValueError('Некорректные страницы или предупреждения пункта.')
                items.append(Item(number=text(entry, 'number'), text=text(entry, 'text'),
                                  id=text(entry, 'id', str(uuid.uuid4())), pages=pages,
                                  checked=entry.get('checked', False), original_text=text(entry, 'original_text'),
                                  warnings=warnings, source_crop=text(entry, 'source_crop'),
                                  ocr_text=text(entry, 'ocr_text'), ocr_status=text(entry, 'ocr_status'),
                                  ocr_crops=ocr_crops, manual=entry.get('manual', False)))
            details = record.get('details', {})
            if not isinstance(details, dict) or type(details.get('control_card', False)) is not bool:
                raise ValueError('Некорректные сведения о PDF.')
            details.setdefault('control_card', False)
            details.setdefault('warnings', [])
            details.setdefault('page_types', [])
            if (not isinstance(details['warnings'], list)
                    or any(not isinstance(w, str) for w in details['warnings'])
                    or not isinstance(details['page_types'], list)
                    or any(not isinstance(p, dict) or type(p.get('page')) is not int
                           or p['page'] < 1 or p.get('type') not in
                           ('cover', 'control', 'sketch', 'unrecognized')
                           for p in details['page_types'])):
                raise ValueError('Некорректные предупреждения или типы листов.')
            graphics = details.get('graphic_rows', [])
            if not isinstance(graphics, list) or any(not isinstance(row, list) or len(row) != 4
                    or type(row[0]) is not int or row[0] < 1
                    or any(not isinstance(x, str) for x in row[1:]) for row in graphics):
                raise ValueError('Некорректные графические обозначения.')
            details['graphic_rows'] = {(row[0], row[1]): (row[2], row[3]) for row in graphics}
            ocr_rows = details.get('ocr_rows', [])
            if (not isinstance(ocr_rows, list) or len(ocr_rows) > 10000
                    or any(not isinstance(row, dict) or type(row.get('page')) is not int
                           or row['page'] < 1 or any(not isinstance(row.get(k, ''), str)
                           for k in ('number', 'crop', 'text', 'status')) for row in ocr_rows)):
                raise ValueError('Некорректные результаты OCR.')
            cards.append(ControlCard(source=text(record, 'source'), document_code=text(record, 'document_code'),
                                     document_name=text(record, 'document_name'), items=items, details=details,
                                     raw=text(record, 'raw'), id=text(record, 'id', str(uuid.uuid4()))))
        if len({card.id for card in cards}) != len(cards) or any(
                len({item.id for item in card.items}) != len(card.items) for card in cards):
            raise ValueError('В проекте повторяются внутренние идентификаторы.')
        images = value.get('images', {})
        if not isinstance(images, dict):
            raise ValueError('Некорректные вырезки PDF.')
        decoded = {}
        for key, encoded in images.items():
            if not isinstance(encoded, str):
                raise ValueError('Некорректная вырезка PDF.')
            data = base64.b64decode(encoded, validate=True)
            if not data.startswith(b'\x89PNG\r\n\x1a\n'):
                raise ValueError('Вырезка PDF не является PNG.')
            decoded[key] = data
    except (TypeError, KeyError, AttributeError, json.JSONDecodeError, UnicodeError, RecursionError) as exc:
        raise ValueError('Не удалось прочитать проект: повреждённые данные.') from exc
    crop_map = {}
    if decoded:
        directory = make_temp_dir('kk_project_crops_')
        for n, (key, data) in enumerate(decoded.items()):
            target = directory / f'crop_{n}.png'
            target.write_bytes(data)
            crop_map[key] = str(target)
    for card in cards:
        for item in card.items:
            item.source_crop = crop_map.get(item.source_crop, '')
            item.ocr_crops = [crop_map[p] for p in item.ocr_crops if p in crop_map]
        for row in card.details.get('ocr_rows', []):
            row['crop'] = crop_map.get(row.get('crop', ''), '')
        card.details['graphic_rows'] = {key: (warning, crop_map.get(crop, ''))
                                      for key, (warning, crop) in card.details['graphic_rows'].items()}
    return cards, settings


def project_filename(name: str, extension: str) -> str:
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip().rstrip('. ')
    stem = stem[:150] or 'Проект'
    if stem.upper().split('.')[0] in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1, 10)), *(f'LPT{i}' for i in range(1, 10))}:
        stem = '_' + stem
    return stem + extension
