"""Склонение названия изделия для титула («паспорт рамы сварной»)."""
from __future__ import annotations

import re
import sys
from functools import lru_cache

from .paths import PIP_INSTALL

_PROTECTED = re.compile(r'[A-Za-zА-Яа-яЁё]+\d+(?:\.\d+)+(?:-\d+|\(-\d+\))?|[«"].*?[»"]')
_WORD = re.compile(r'[А-Яа-яЁё]+(?:-[А-Яа-яЁё]+)*')


@lru_cache(maxsize=1)
def _analyzer():
    """Словарь pymorphy3 загружается один раз и только при первом склонении."""
    try:
        import pymorphy3
    except ImportError as exc:
        raise RuntimeError(f'Для склонения названий установите: {PIP_INSTALL} pymorphy3') from exc
    path = _frozen_dictionary_path()
    return pymorphy3.MorphAnalyzer(path=path, lang='ru') if path else pymorphy3.MorphAnalyzer()


def _frozen_dictionary_path():
    """В собранном .exe словарь ищется напрямую (метаданные пакетов могут отсутствовать)."""
    if not getattr(sys, 'frozen', False):
        return None
    try:
        import pymorphy3_dicts_ru
    except ImportError:
        return None
    get_path = getattr(pymorphy3_dicts_ru, 'get_path', None)
    return get_path() if callable(get_path) else None


def decline_product_title(name: str) -> str:
    """Inflect the product head and agreeing words; preserve codes and dependencies.

    Russian morphology is ambiguous: the cover dialog also offers an explicit override.
    """
    if not name.strip():
        return name
    morph = _analyzer()
    protected = [m.span() for m in _PROTECTED.finditer(name)]
    words = list(_WORD.finditer(name))
    entries = []
    for match in words:
        word = match.group()
        if any(a <= match.start() < b for a, b in protected):
            continue
        analyses = morph.parse(word)
        nouns = [v for v in analyses if v.tag.POS == 'NOUN' and 'Abbr' not in v.tag
                 and not any(g in v.tag for g in ('Name', 'Surn', 'Patr', 'Geox', 'Orgn'))]
        if word.isupper() and len(word) > 1 and not nouns:
            continue
        # Product names are normally supplied in the nominative. A dictionary
        # may rank a homonymous plural genitive above the singular nominative
        # (e.g. «остов»). Prefer that singular nominative for the FIRST product
        # noun, including when an adjective precedes it. Dependent nouns such
        # as «рамы» / «двигателя» retain their original analysis.
        first_noun = not any(v.tag.POS == 'NOUN' for _, v in entries)
        coordinated = bool(entries and entries[-1][0].group().lower() in ('и', 'или'))
        if first_noun and analyses[0].tag.POS in ('ADJF', 'PRTF'):
            parsed = analyses[0]
        elif nouns and (first_noun or coordinated):
            nominative_singular = [v for v in nouns if 'nomn' in v.tag and 'sing' in v.tag]
            parsed = nominative_singular[0] if nominative_singular else nouns[0]
            if not nominative_singular:
                # A neighbouring plural adjective can disambiguate «каретки»
                # (singular genitive OR plural nominative). Do not guess solely
                # from capitalisation, and do not change an already-genitive name.
                position = words.index(match)
                neighbours = words[max(0, position - 1):position] + words[position + 1:position + 2]
                plural_agreement = any(
                    any(v.tag.POS in ('ADJF', 'PRTF') and 'nomn' in v.tag and 'plur' in v.tag
                        for v in morph.parse(neighbour.group())) for neighbour in neighbours
                    if not any(a <= neighbour.start() < b for a, b in protected))
                plural_nouns = [v for v in nouns if 'nomn' in v.tag and 'plur' in v.tag]
                if plural_agreement and plural_nouns:
                    parsed = plural_nouns[0]
        else:
            parsed = analyses[0]
        entries.append((match, parsed))
    head = next((parsed for _, parsed in entries if parsed.tag.POS == 'NOUN' and 'nomn' in parsed.tag), None)
    if head is None:
        return name  # Already genitive, an abbreviation or an unknown construction.
    replacements = []
    for match, parsed in entries:
        if 'nomn' not in parsed.tag:
            continue
        if parsed.tag.POS == 'NOUN':
            target = parsed.inflect({'gent'})
        elif parsed.tag.POS in ('ADJF', 'PRTF') and parsed.tag.number == head.tag.number and (
                head.tag.number == 'plur' or parsed.tag.gender == head.tag.gender):
            target = parsed.inflect({'gent'})
        else:
            continue
        if target:
            text = target.word
            # A running title uses lower case after «ПАСПОРТ»; codes/quoted names are protected.
            replacements.append((match.start(), match.end(), text))
    for a, b, text in reversed(replacements):
        name = name[:a] + text + name[b:]
    return name
