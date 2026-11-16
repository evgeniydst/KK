"""Данные чек-листа: пункты, карты контроля и номера пунктов."""
from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field

# Supports 1. Text, 1) Text and 1.2. Text; requires a space after punctuation.
ITEM_START = re.compile(r"^\s*(\d+(?:\.\d+)*)[.)]\s+(.+?)\s*$")
PAGE_MARKER = re.compile(r"^<<<PDF_PAGE:(\d+)>>>$")
WELD_WORD = re.compile(r"\b(?:шов|шва|шву|шве|швом|швы|швов|швам|швами|швах)\b", re.IGNORECASE)


@dataclass
class Item:
    number: str
    text: str
    pages: list[int] = field(default_factory=list)
    checked: bool = False
    id: str = field(default_factory=lambda: str(uuid.uuid4()))

    original_text: str = ""
    warnings: list[str] = field(default_factory=list)
    source_crop: str = ""
    ocr_text: str = ""
    ocr_status: str = ""
    ocr_crops: list[str] = field(default_factory=list)
    manual: bool = False

    @property
    def label(self) -> str:
        return f"{self.number}. {self.text}" if self.number else self.text


@dataclass
class ControlCard:
    source: str
    document_code: str = ""
    document_name: str = ""
    items: list[Item] = field(default_factory=list)
    details: dict = field(default_factory=dict)
    raw: str = ""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))


_NUMBER = re.compile(r"[0-9]+(?:\.[0-9]+)*")
# ASCII: «٣» и другие нелатинские цифры дают понятную ошибку, а не сбой.
_EXPRESSION_NUMBER = re.compile(r'\d+(?:\.\d+)*', re.ASCII)
_EXPRESSION_RANGE = re.compile(r'(\d+)\s*-\s*(\d+)', re.ASCII)
_DASHES = str.maketrans({'‐': '-', '‑': '-', '–': '-', '−': '-', '—': '-'})


def deselect_weld_items(items: list[Item]) -> int:
    """Uncheck weld-word matches in the current edited text; never delete rows."""
    changed = 0
    for item in items:
        if item.checked and WELD_WORD.search(item.text):
            item.checked = False
            changed += 1
    return changed


def number_identity(number: str) -> tuple[int, ...] | None:
    number = number.strip()
    if _NUMBER.fullmatch(number):
        return tuple(int(part) for part in number.split("."))
    return None


def item_sort_key(item: Item) -> tuple:
    key = number_identity(item.number)
    return (0, key) if key is not None else (1, (item.number,))


def ordered_items(items: list[Item]) -> list[Item]:
    """Preserve the user's explicit sequence once a manual row is present."""
    return list(items) if any(item.manual for item in items) else sorted(items, key=item_sort_key)


def duplicate_numbers(items: list[Item]) -> list[str]:
    seen, duplicates = set(), []
    for item in items:
        key = number_identity(item.number)
        if key is not None:
            if key in seen and item.number not in duplicates:
                duplicates.append(item.number)
            seen.add(key)
    return duplicates


def select_number_expression(expression: str, items: list[Item]) -> tuple[list[str], int, list[str]]:
    """Match ORIGINAL point numbers, not row positions; never enumerate huge ranges.

    Returns selected item IDs, count of missing numbers, and up to 20 missing
    labels. Integer ranges do not implicitly include hierarchical subpoints.
    """
    expression = expression.strip().translate(_DASHES)
    if not expression or len(expression) > 2000:
        raise ValueError('Введите номера, например: 1-10 или 1,2,3-10,15.')
    intervals, subpoints = [], set()
    for token in expression.split(','):
        token = token.strip()
        if _EXPRESSION_NUMBER.fullmatch(token):
            key = number_identity(token)
            if len(key) == 1:
                intervals.append((key[0], key[0]))
            else:
                subpoints.add(key)
        else:
            match = _EXPRESSION_RANGE.fullmatch(token)
            if not match:
                raise ValueError('Неверная запись «' + token + '». Используйте 1-10,15,20 или отдельный подпункт 1.2.')
            start, end = map(int, match.groups())
            if start > end:
                raise ValueError('Начало диапазона больше конца: ' + token)
            intervals.append((start, end))
    merged = []
    for start, end in sorted(intervals):
        if merged and start <= merged[-1][1] + 1:
            merged[-1] = (merged[-1][0], max(end, merged[-1][1]))
        else:
            merged.append((start, end))
    keys = {number_identity(item.number) for item in items}
    numeric = sorted(k[0] for k in keys if k is not None and len(k) == 1)
    selected = [item.id for item in items if (key := number_identity(item.number)) is not None
                and (key in subpoints or len(key) == 1 and any(a <= key[0] <= b for a, b in merged))]
    missing_count = 0
    preview = []
    for start, end in merged:
        existing = [n for n in numeric if start <= n <= end]
        missing_count += end - start + 1 - len(existing)
        cursor = start
        for n in existing + [end + 1]:
            if len(preview) < 20 and cursor < n:
                preview.extend(str(v) for v in range(cursor, min(n, cursor + 20 - len(preview))))
            cursor = n + 1
    for key in sorted(subpoints - keys):
        missing_count += 1
        if len(preview) < 20:
            preview.append('.'.join(map(str, key)))
    return selected, missing_count, preview


def next_free_number(items: list[Item]) -> str:
    used = {number_identity(item.number) for item in items}
    number = 1
    while (number,) in used:
        number += 1
    return str(number)


def commit_item(items: list[Item], item: Item, number: str, text: str, new: bool = False,
                insert_index: int | None = None) -> None:
    """Validate before any mutation. Uniqueness is local to one control card."""
    number, text = number.strip(), text.strip()
    key = number_identity(number)
    if key is None and not item.manual:
        raise ValueError("Введите номер вида 1 или 4.1, без точки в конце.")
    if item.manual and (len(number) > 80 or '\n' in number or '\r' in number):
        raise ValueError("Номер или обозначение должно быть в одну строку и не длиннее 80 символов.")
    if not text:
        raise ValueError("Введите текст пункта.")
    if key is not None and any(other.id != item.id and number_identity(other.number) == key for other in items):
        raise ValueError(f"Пункт № {number} уже есть в этой КК. Укажите свободный номер.")
    if new and insert_index is not None and not 0 <= insert_index <= len(items):
        raise ValueError("Не удалось определить место вставки пункта.")
    item.number, item.text = number, text
    if new:
        if item.manual and insert_index is not None:
            items.insert(insert_index, item)
        else:
            items.append(item)
    if not any(other.manual for other in items):
        items.sort(key=item_sort_key)
