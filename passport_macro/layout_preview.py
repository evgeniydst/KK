"""Схема раскладки таблиц КК по страницам для миниатюры в окне.

Повторяет решения _build_control_tables (те же оценки высот и тот же
планировщик), но записывает, где окажется каждая часть таблицы. Это оценка
программы, а не вёрстка Word: точное положение строк определяет Word.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .docx_export import (CONTROL_COLUMN_WIDTHS, SPLIT_LINE_HEIGHT, SPLIT_MIN_LINES,
                          estimate_control_parts)
from .models import ordered_items
from .readiness import Issue
from .table_settings import BUILTIN_TABLE_LAYOUT

PT_PER_MM = 72 / 25.4
PAGE_WIDTH_MM = 297
PAGE_HEIGHT_MM = 210
MARGIN_MM = 10
COLUMN_GAP_MM = 8
COLUMN_WIDTH_MM = (PAGE_WIDTH_MM - 2 * MARGIN_MM - COLUMN_GAP_MM) / 2
# Как в docx_export: высота страницы минус поля минус 6 pt запаса.
COLUMN_CAPACITY = (PAGE_HEIGHT_MM - 2 * MARGIN_MM) * PT_PER_MM - 6
CARD_GAP = 2 * PT_PER_MM + 1  # Mm(2).pt + 1


@dataclass
class Segment:
    column: int      # номер колонки раздела таблиц, с нуля (2 колонки на лист)
    top: float       # pt от верха колонки
    bottom: float
    kind: str        # header | row | signatures
    card: int        # номер КК в паспорте (с 1)
    label: str = ''
    overflow: bool = False


@dataclass
class LayoutPreview:
    mode: str
    segments: list = field(default_factory=list)
    columns: int = 0
    pages: int = 0           # листы с таблицами, без титульного
    cards: int = 0
    items: int = 0
    issues: list = field(default_factory=list)
    capacity: float = COLUMN_CAPACITY


def selected_groups(cards: list) -> list:
    """(номер КК, КК, отмеченные пункты) без проверок — для схемы."""
    return [(n, card, [i for i in ordered_items(card.items) if i.checked])
            for n, card in enumerate(cards, 1) if any(i.checked for i in card.items)]


def place_free(card_no, labels, column, column_used, capacity, header, rows, signatures,
               has_previous=False, gap=0.0, splittable=None):
    """Тот же алгоритм, что plan_free_control_layout, но с записью частей.

    Возвращает (части, последняя колонка, занято в ней, перенос до таблицы).
    """
    if capacity <= 0 or not rows:
        raise ValueError('Нет места для контрольной таблицы или отсутствуют пункты.')
    segments = []
    first_block = header + rows[0] + (signatures if len(rows) == 1 else 0)
    required = first_block
    start = column_used + gap if has_previous else 1.0
    break_before = bool(has_previous and start + required > capacity)
    if break_before:
        column += 1
        start = 1.0
    segments.append(Segment(column, start, start + header, 'header', card_no))
    used = start + header
    last = len(rows) - 1
    for index, height in enumerate(rows):
        label = labels[index] if index < len(labels) else ''
        block = height + (signatures if index == last else 0)
        if (used + block > capacity and splittable and splittable[index]
                and capacity - used >= 2 * SPLIT_LINE_HEIGHT):
            segments.append(Segment(column, used, capacity, 'row', card_no, label))
            column += 1
            segments.append(Segment(column, 0.0, header, 'header', card_no, 'продолжение'))
            new_used = min(capacity, header + used + block - capacity)
            segments.append(Segment(column, header, new_used, 'row', card_no, label))
            used = new_used
            continue
        if used + block > capacity:
            column += 1
            segments.append(Segment(column, 0.0, header, 'header', card_no, 'продолжение'))
            used = header
        if used + block > capacity:
            segments.append(Segment(column, used, capacity, 'row', card_no, label, overflow=True))
            used = capacity
        else:
            segments.append(Segment(column, used, used + height, 'row', card_no, label))
            if index == last:
                segments.append(Segment(column, used + height, used + block, 'signatures', card_no))
            used += block
    return segments, column, used, break_before


def _flow_whole(card_no, labels, column, start, capacity, header, rows, signatures):
    """Таблица «целиком»: строки не разрываются, шапка повторяется в новой колонке."""
    segments = [Segment(column, start, start + header, 'header', card_no)]
    used = start + header
    last = len(rows) - 1
    for index, height in enumerate(rows):
        label = labels[index] if index < len(labels) else ''
        block = height + (signatures if index == last else 0)
        if used + block > capacity and used > header + 0.5:
            column += 1
            segments.append(Segment(column, 0.0, header, 'header', card_no, 'продолжение'))
            used = header
        overflow = used + block > capacity
        bottom = min(capacity, used + height)
        segments.append(Segment(column, used, bottom, 'row', card_no, label, overflow=overflow))
        if index == last:
            segments.append(Segment(column, bottom, min(capacity, used + block), 'signatures', card_no))
        used = min(capacity, used + block)
    return segments, column


def preview_layout(cards: list, settings: dict | None = None,
                   capacity: float = COLUMN_CAPACITY, estimate=estimate_control_parts) -> LayoutPreview:
    mode = (settings or {}).get('word_layout_mode', 'free')
    mode = mode if mode in ('free', 'whole') else 'free'
    result = LayoutPreview(mode=mode, capacity=capacity)
    groups = selected_groups(cards)
    table_layout = (settings or {}).get('table_layout', BUILTIN_TABLE_LAYOUT)
    widths = table_layout['widths']
    column, column_used, previous_long = 0, 0.0, False
    for group_index, (card_no, card, items) in enumerate(groups):
        if estimate is estimate_control_parts:
            header, rows, signatures = estimate(card, items, widths, table_layout['header_min_mm'],
                                                 table_layout['row_min_mm'])
        else:
            header, rows, signatures = estimate(card, items, widths)
        labels = [item.number or '?' for item in items]
        result.items += len(items)
        if mode == 'whole':
            estimated = header + sum(rows) + signatures
            keep = estimated <= capacity
            start_next = bool(group_index and (previous_long or not keep or
                                               column_used + CARD_GAP + estimated > capacity))
            if start_next:
                column += 1
                column_used = 0.0
            elif group_index:
                column_used += CARD_GAP
            segments, last_column = _flow_whole(card_no, labels, column, column_used, capacity,
                                                header, rows, signatures)
            if keep:
                column_used += estimated
            else:
                column, column_used = last_column, capacity
            previous_long = not keep
        else:
            splittable = [0 < index < len(items) - 1 and
                          height >= SPLIT_MIN_LINES * SPLIT_LINE_HEIGHT
                          for index, height in enumerate(rows)]
            segments, column, column_used, _ = place_free(
                card_no, labels, column, column_used, capacity, header, rows, signatures,
                has_previous=bool(group_index), gap=CARD_GAP, splittable=splittable)
        result.segments.extend(segments)
        for segment in segments:
            if segment.overflow:
                item = next((i for i in items if (i.number or '?') == segment.label), None)
                result.issues.append(Issue(
                    'warning', 'cards',
                    f'КК {card_no}, пункт № {segment.label}: текст длиннее колонки — '
                    'проверьте этот пункт в готовом Word.',
                    card_no - 1, item.id if item else None))
    result.cards = len(groups)
    if result.segments:
        result.columns = max(segment.column for segment in result.segments) + 1
        result.pages = (result.columns + 1) // 2
    return result


def column_fill(preview: LayoutPreview) -> list[float]:
    """Доля заполнения каждой колонки (0..1) — для подписи «пустого места»."""
    fill = [0.0] * preview.columns
    for segment in preview.segments:
        fill[segment.column] = max(fill[segment.column], segment.bottom / preview.capacity)
    return fill
