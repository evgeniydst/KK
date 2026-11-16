"""Проверка готовности проекта к формированию паспорта (без интерфейса).

Уровни замечаний:
  error   — паспорт не сформируется или титульный лист будет неполным;
  warning — сформируется, но стоит проверить;
  note    — просто к сведению, шаг не помечается как проблемный.
"""
from __future__ import annotations

from dataclasses import dataclass

from .models import duplicate_numbers, ordered_items

TITLE_STEP = 'title'
CARDS_STEP = 'cards'
LEVELS = ('error', 'warning', 'note')


@dataclass
class Issue:
    level: str
    step: str
    text: str
    card_index: int | None = None
    item_id: str | None = None


def _text(value) -> str:
    return value.strip() if isinstance(value, str) else ''


def title_issues(settings: dict) -> list[Issue]:
    issues = []
    if not _text(settings.get('name')):
        issues.append(Issue('error', TITLE_STEP,
                            'Не заполнено наименование и обозначение изделия.'))
    for key, fio_key, label in (('approval_role', 'approval_name', '«Утвердил»'),
                                ('developer_role', 'developer_name', '«Разработал»')):
        if not _text(settings.get(key)):
            issues.append(Issue('error', TITLE_STEP, f'Не заполнена должность в блоке {label}.'))
        elif not _text(settings.get(fio_key)):
            issues.append(Issue('note', TITLE_STEP,
                                f'ФИО в блоке {label} не заполнено — останется линия для подписи.'))
    for number, person in enumerate(settings.get('reviewers') or [], 1):
        if not _text(person.get('role')):
            issues.append(Issue('error', TITLE_STEP,
                                f'Согласующий {number}: не заполнена должность '
                                '(заполните или удалите строку).'))
    if settings.get('types_enabled') and not any(_text(t) for t in settings.get('types') or []):
        issues.append(Issue('warning', TITLE_STEP,
                            'Блок «Типы / исполнения» включён, но список пуст.'))
    return issues


def card_issues(cards: list) -> list[Issue]:
    if not cards:
        return [Issue('error', CARDS_STEP, 'Нет ни одной карты контроля — добавьте PDF.')]
    issues = []
    if not any(item.checked for card in cards for item in card.items):
        issues.append(Issue('error', CARDS_STEP,
                            'Не отмечен ни один пункт — отметьте пункты для паспорта.'))
    for index, card in enumerate(cards):
        name = f'КК {index + 1}'
        checked = [item for item in ordered_items(card.items) if item.checked]
        if not checked:
            label = _text(card.document_code) or 'без обозначения'
            issues.append(Issue('note', CARDS_STEP,
                                f'{name} ({label}): нет отмеченных пунктов — в паспорт не попадёт.',
                                index))
            continue
        if not _text(card.document_code):
            issues.append(Issue('error', CARDS_STEP, f'{name}: не заполнено обозначение детали.', index))
        if not _text(card.document_name):
            issues.append(Issue('error', CARDS_STEP, f'{name}: не заполнено наименование изделия.', index))
        repeated = duplicate_numbers(checked)
        if repeated:
            issues.append(Issue('error', CARDS_STEP,
                                f'{name}: повторяются номера пунктов {", ".join(repeated)}.', index))
        for item in checked:
            if not _text(item.text):
                issues.append(Issue('warning', CARDS_STEP,
                                    f'{name}, пункт № {item.number}: пустой текст.', index, item.id))
            elif item.warnings:
                issues.append(Issue('warning', CARDS_STEP,
                                    f'{name}, пункт № {item.number}: часть символов со скана '
                                    'распознана неуверенно — сверьте с PDF.', index, item.id))
    return issues


def collect_issues(cards: list, settings: dict) -> list[Issue]:
    return title_issues(settings) + card_issues(cards)


def count_by_step(issues: list[Issue]) -> dict:
    """{шаг: (ошибок, предупреждений)}; заметки не считаются."""
    counts = {TITLE_STEP: [0, 0], CARDS_STEP: [0, 0]}
    for issue in issues:
        row = counts.setdefault(issue.step, [0, 0])
        if issue.level == 'error':
            row[0] += 1
        elif issue.level == 'warning':
            row[1] += 1
    return {step: tuple(value) for step, value in counts.items()}


def blocking(issues: list[Issue]) -> list[Issue]:
    return [issue for issue in issues if issue.level == 'error']
