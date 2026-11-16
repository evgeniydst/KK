"""Чтение карт контроля из PDF: геометрия таблицы, ГОСТ-шрифты, OCR-вырезки."""
from __future__ import annotations

import logging
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

from .models import ITEM_START, PAGE_MARKER, ControlCard, Item, item_sort_key, number_identity
from .paths import PIP_INSTALL, bundle_dir
from .tempfiles import make_temp_dir

log = logging.getLogger(__name__)

_JOINED_PLUS_MINUS = re.compile(r"(?<=\d)\s*±\s*(?=\d)")
_MOJIBAKE = re.compile(r"[À-ÖÙ-ÿ]")
_AMBIGUOUS_VIEW_LETTER = re.compile(r"[-–—]Ø(?=\s|$|[.,;])")


def _fitz():
    """PyMuPDF загружается при первом обращении: окно открывается и без него."""
    try:
        import pymupdf as fitz
    except ImportError as exc:
        raise RuntimeError(f"Установите библиотеку: {PIP_INSTALL} PyMuPDF") from exc
    return fitz


def card_from_extraction(source: str, details: dict) -> ControlCard:
    items = build_card_items(details)
    items.sort(key=item_sort_key)
    return ControlCard(source, details.get("document_code", ""), details.get("document_name", ""),
                       items, details, details["raw"])


# Only explicit drawing references are removed; unknown phrases remain unchanged.
REFERENCE = re.compile(
    r"(?<!\w)(?:Эскиз|Виды?|Разрез|Сечение)\s+"
    r"(?:[А-ЯЁA-Z]\d*(?:\s*[-–—]\s*[А-ЯЁA-Z]\d*)?(?:,\s*[А-ЯЁA-Z]\d*)*|"
    r"\d+(?:\.\d+)*(?:,\s*\d+(?:\.\d+)*(?!\d|\s*мест))*)(?![\w-])"
    r"(?:,?\s*\d+\s*мест(?:а|о)?(?!\w))?\.?", re.IGNORECASE)


def remove_drawing_references(text: str) -> str:
    # Preserve requirement counts; only counts attached to a drawing reference
    # belong to that reference. Keep original_text untouched for manual review.
    marker = "\ufff0"
    result = []
    inherited = re.compile(
        r"\s*[А-ЯЁA-Z]\d*\s*[-–—]\s*[А-ЯЁA-Z]\d*"
        r"(?:,\s*\d+\s*мест(?:а|о)?)?[.\s]*", re.IGNORECASE)
    for line in text.splitlines():
        has_section = bool(re.search(r"\bРазрез\s+", line, re.IGNORECASE))
        pieces = line.split(";")
        cleaned = []
        for piece in pieces:
            if has_section and inherited.fullmatch(piece):
                piece = marker
            else:
                piece = REFERENCE.sub(marker, piece)
            # A conjunction between two removed references is not body text.
            piece = re.sub(marker + r"\s+и\s+" + marker, marker, piece,
                           flags=re.IGNORECASE)
            if marker in piece:
                piece = piece.replace(marker, "").strip()
                if not re.search(r"\w", piece):
                    continue
            cleaned.append(piece.rstrip())
        if cleaned:
            value = ";".join(cleaned).rstrip()
            if REFERENCE.search(line):
                value = value.rstrip("; ")
            if value:
                result.append(value)
    return "\n".join(result).strip()


def repair_gost_text(text: str, font: str) -> str:
    """Repair the confirmed Latin-byte mapping, never use errors='ignore'.

    ASCII, already-Cyrillic text, ±, ° and standalone diameter Ø are preserved.
    Short ambiguous Latin fragments are not repaired automatically.
    """
    if "GOST" not in font.upper() or re.search(r"[А-Яа-яЁё]", text):
        return text
    candidates = re.findall(r"[À-ÿ]", text)
    # A short GOST prefix in a dotted code is also evidence of the byte mapping.
    short_code = re.fullmatch(r"[À-ÿ]{1,8}[0-9]+(?:\.[0-9]+)+(?:-\d+|\(-\d+\))?(?:\.)?|[À-ÿ]{1,8}[0-9]+\.", re.sub(r'\s+', '', text.translate(str.maketrans({'–':'-', '−':'-', '‑':'-', '—':'-'}))))
    if len(candidates) < 3 and not short_code:
        return text
    result = []
    for i, char in enumerate(text):
        # Ø followed by a dimension is not treated as a Cyrillic letter.
        reference_letter = char == "Ø" and re.search(r"(?:Вид|Разрез|Сечение)\s+$", "".join(result), re.IGNORECASE)
        diameter = char == "Ø" and not reference_letter and (i + 1 == len(text) or text[i + 1].isdigit()
                                    or text[i + 1].isspace())
        if re.fullmatch(r"[À-ÿ]", char) and not diameter:
            try:
                char = char.encode("cp1252").decode("cp1251")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass  # Leave unverifiable symbols intact.
        # Confirmed byte B9 is the numero sign in the weld-number field.
        if char == "¹" and re.search(r"швов\s*$", "".join(result), re.IGNORECASE) and text[i + 1:i + 2].isdigit():
            char = "№"
        result.append(char)
    return "".join(result)


def product_code_key(text: str) -> str:
    """Equivalent execution suffixes share a key, but the displayed spelling is kept."""
    text = text.translate(str.maketrans({'‐': '-', '‑': '-', '–': '-', '−': '-', '—': '-'}))
    text = text.strip()
    # Normalize only the execution suffix: never join unrelated page numbers
    # or arbitrary digits separated by spaces into the base designation.
    return re.sub(r'\s*(?:\(\s*-\s*([0-9]+)\s*\)|-\s*([0-9]+))\s*$',
                  lambda match: '-' + (match.group(1) or match.group(2)), text)


def is_product_code(text: str) -> bool:
    """Ignore long service designations, but keep product execution suffixes."""
    key = product_code_key(text)
    base = re.sub(r'-[0-9]+$', '', key)
    # D40.60102.11020 is not a product designation for this workflow.
    # Reject the whole long format, not just this particular service number.
    # D40.11.020 (including execution suffixes) remains a product code.
    if re.fullmatch(r"[A-Za-zА-Яа-яЁё]+[0-9]+\.[0-9]{5}\.[0-9]{5}", base):
        return False
    return (re.fullmatch(r"[A-Za-zА-Яа-яЁё]+[0-9]+(?:\.[0-9]+){2,}(?:-[0-9]+)?", key) is not None
            and '02190' not in base.split('.')[1:])


def identity_code_spans(spans, page_rect):
    """Join only adjacent numeric/suffix fragments in the header, not other fields."""
    fitz = _fitz()
    result = []
    header = sorted([v for v in spans if v[0].y1 < page_rect.height * .3], key=lambda v: (v[0].y0, v[0].x0))
    for rect, text, _ in header:
        text = text.strip()
        if not re.match(r'^[A-Za-zА-Яа-яЁё]+[0-9]+\.', text):
            continue
        combined, bound = text.strip(), fitz.Rect(rect)
        best = (fitz.Rect(bound), combined) if is_product_code(combined) else None
        for _ in range(6):
            # Italic glyph boxes can overlap although the printed code and
            # execution are adjacent (e.g. .100 + (-01), in different fonts).
            # Allow only a small, height-scaled overhang for an explicit
            # suffix. Ordinary numeric fragments keep the stricter limit;
            # this must not join page numbers or normalize base=execution.
            near = [(r, t) for r, t, _ in header if
                    -(min(4, min(r.height, bound.height) * .25)
                      if re.match(r'^[()‐‑–−—-]', t.strip()) else .2)
                    <= r.x0 - bound.x1 <= (12 if re.match(r'^[()‐‑–−—-]', t.strip()) else 4)
                    and abs((r.y0+r.y1)/2 - (bound.y0+bound.y1)/2) < 2
                    and r.x1 > bound.x1 and re.fullmatch(r'[0-9.()‐‑–−—-]+', re.sub(r'\s+', '', t))]
            if not near:
                break
            r, t = min(near, key=lambda v: v[0].x0)
            # A bare page number cannot extend an already complete code.
            if best and re.fullmatch(r'[0-9]+', t.strip()) and not re.sub(r'\s+', '', combined).endswith(('-', '(', '.', '−', '–', '‑', '—')):
                break
            combined += t.strip(); bound |= r
            if is_product_code(combined):
                best = (fitz.Rect(bound), combined)
        if best:
            result.append(best)
    return result


def read_identity(spans, drawings, page_rect):
    """Read code and the bordered name field immediately below it, not body text."""
    fitz = _fitz()
    # Prefix may be Cyrillic or Latin; numeric group lengths are not fixed.
    # The broad pattern generates candidates, geometry below selects the field.
    header_codes = identity_code_spans(spans, page_rect)
    horizontal = []
    for drawing in drawings:
        for segment in drawing["items"]:
            if segment[0] == "l":
                a, b = segment[1:3]
                if abs(a.y - b.y) < .2:
                    horizontal.append((a.y, min(a.x, b.x), max(a.x, b.x)))
            elif segment[0] == "re":
                # Some PDF writers encode rules as thin filled rectangles, not lines.
                # Ignore large filled areas (redactions/backgrounds), only horizontal rules.
                rect = fitz.Rect(segment[1])
                if 0 < rect.height <= 2 and rect.width > max(10, rect.height * 10):
                    horizontal.append(((rect.y0 + rect.y1) / 2, rect.x0, rect.x1))
    names, paired_codes = [], []
    for code_rect, code in header_codes:
        candidates = []
        for r, text, _ in spans:
            if not (code_rect.y1 < r.y0 and r.y1 < code_rect.y1 + page_rect.height * .065
                    and r.x0 > page_rect.width * .3 and r.x1 < page_rect.width * .9
                    and len(re.findall(r"[А-Яа-яЁё]+", text)) >= 1):
                continue
            above = [(y, x0, x1) for y, x0, x1 in horizontal
                     if code_rect.y1 - 2 <= y <= r.y0 and x0 <= r.x0 and x1 >= r.x1]
            below = [(y, x0, x1) for y, x0, x1 in horizontal
                     if r.y1 <= y <= code_rect.y1 + page_rect.height * .075
                     and x0 <= r.x0 and x1 >= r.x1]
            # The field borders must extend beneath the code cell too.
            if above and below:
                upper = max(above)
                lower = min(below)
                if (upper[1] <= code_rect.x0 and upper[2] >= code_rect.x1
                        and lower[1] <= code_rect.x0 and lower[2] >= code_rect.x1
                        # First-sheet lower border can extend left through the
                        # signatories block; its right edge must still agree.
                        and lower[1] <= upper[1] + 2 and abs(upper[2] - lower[2]) < 2):
                    candidates.append((r, text, upper, lower))
        if candidates:
            nearest = min(r.y0 for r, *_ in candidates)
            row = sorted([v for v in candidates if abs(v[0].y0 - nearest) < 2], key=lambda v: v[0].x0)
            names.append(" ".join(v[1] for v in row))
            paired_codes.append(code)
    # On a cover/first sheet, select the code with the bordered name field below.
    # Form 2a may have no name field: its designation occupies the lower header row.
    if paired_codes:
        codes = paired_codes
    else:
        codes = [t for r, t in header_codes
                 if r.y0 > page_rect.height * .18 and r.x0 > page_rect.width * .45
                 and r.x1 < page_rect.width * .8]
    return sorted(set(codes)), sorted(set(names))


def vector_plus_minus(drawing) -> bool:
    """Recognize only the verified ± geometry: two bars and a crossing vertical stroke."""
    rect = drawing["rect"]
    segments = drawing["items"]
    if not (2 <= rect.width <= 12 and 3 <= rect.height <= 15 and len(segments) == 3
            and all(v[0] == "l" for v in segments)):
        return False
    bars, stems = [], []
    for _, a, b in segments:
        if abs(a.y - b.y) < .2:
            bars.append((a.y, min(a.x, b.x), max(a.x, b.x)))
        elif abs(a.x - b.x) < .2:
            stems.append((a.x, min(a.y, b.y), max(a.y, b.y)))
    if len(bars) != 2 or len(stems) != 1:
        return False
    upper, lower = sorted(bars)
    x, y0, y1 = stems[0]
    return (abs(upper[1] - lower[1]) < .3 and abs(upper[2] - lower[2]) < .3
            and upper[2] - upper[1] > 2
            and abs(x - (upper[1] + upper[2]) / 2) < .4
            and y0 + .5 < upper[0] < y1 - .5 and lower[0] > y1 + .5)


def parameter_lines(spans):
    """Attach small upper/lower tolerances to the main text line, not the next item."""
    if not spans:
        return []
    heights = sorted(r.height for r, t, _ in spans if t.strip())
    if not heights:
        return []
    body_heights = [h for h in heights if h >= heights[-1] * .8]
    regular_height = body_heights[len(body_heights) // 2]
    # Some PDFs store the sign and tolerance value in separate font fragments.
    # Join only a unique neighbouring number on the SAME baseline/height.
    fitz = _fitz()
    spans = list(spans)
    for sign in list(spans):
        if sign not in spans or sign[1].strip() not in ('+', '-', '−'):
            continue
        r = sign[0]
        nearby = [v for v in spans if v is not sign
                  and re.fullmatch(r"\d+(?:[.,]\d+)?", v[1].strip())
                  and abs(v[0].y0 - r.y0) < .5
                  and abs(v[0].height - r.height) < .5
                  and -3 <= v[0].x0 - r.x1 <= 3]
        if len(nearby) == 1:
            number = nearby[0]
            spans.remove(sign); spans.remove(number)
            spans.append((fitz.Rect(min(r.x0, number[0].x0), min(r.y0, number[0].y0),
                                    max(r.x1, number[0].x1), max(r.y1, number[0].y1)),
                          sign[1].strip().replace('−', '-') + number[1].strip(),
                          sign[2] + number[2]))
    normal = [v for v in spans if v[1].strip() and v[0].height >= regular_height * .8]
    small = [v for v in spans if v[1].strip() and v[0].height < regular_height * .8]
    rows = []
    for r, text, raw in sorted(normal, key=lambda v: (v[0].y0, v[0].x0)):
        if not rows or abs(rows[-1][0] - r.y0) > 3:
            rows.append((r.y0, []))
        rows[-1][1].append((r, text, raw))
    for value in small:
        r = value[0]
        if not rows:
            rows.append((r.y0, [value])); continue
        target = min(rows, key=lambda row: abs((r.y0 + r.y1) / 2 -
                     sum((v[0].y0 + v[0].y1) / 2 for v in row[1]) / len(row[1])))
        center = sum((v[0].y0 + v[0].y1) / 2 for v in target[1]) / len(target[1])
        if abs((r.y0 + r.y1) / 2 - center) > regular_height:
            rows.append((r.y0, [value]))
        else:
            target[1].append(value)
    for y, parts in rows:
        # Only explicitly signed, vertically stacked values are combined as +upper/-lower.
        upper = [v for v in parts if v[0].height < regular_height * .8
                 and re.fullmatch(r"\+\d+(?:[.,]\d+)?", v[1].strip())]
        for up in upper:
            lower = [v for v in parts if v[0].height < regular_height * .8
                     and re.fullmatch(r"-\d+(?:[.,]\d+)?", v[1].strip())
                     and abs(v[0].x0 - up[0].x0) < 3 and v[0].y0 > up[0].y0]
            if len(lower) == 1:
                down = lower[0]
                parts.remove(up); parts.remove(down)
                rect = fitz.Rect(min(up[0].x0, down[0].x0), y,
                                 max(up[0].x1, down[0].x1), y + regular_height)
                parts.append((rect, up[1].strip() + "/" + down[1].strip(),
                              up[2] + "/" + down[2]))
    return sorted(rows, key=lambda row: row[0])


def split_parameter_spans(page, chosen, right, safe_right):
    """Fallback only for overflowing spans; split by real glyph coordinates.

    Whitespace is not content. Tokens are never truncated character by character.
    Classify a crossing token by its centre and flag the source for human review.
    Returns parameter spans and source rectangles requiring review.
    """
    fitz = _fitz()
    raw_spans = None
    selected, reviews = [], []
    for rect, text, original in chosen:
        if rect.x1 <= safe_right:
            selected.append((rect, text, original))
            continue
        if raw_spans is None:
            raw_spans = [span for block in page.get_text('rawdict')['blocks']
                         for line in block.get('lines', []) for span in line['spans']]
        matches = [span for span in raw_spans
                   if all(abs(a - b) < .02 for a, b in zip(span['bbox'], rect))
                   and ''.join(c['c'] for c in span['chars']) == original]
        if len(matches) != 1:
            raise ValueError('Не удалось сопоставить координаты символов спорной строки; '
                             'автоматическое отделение инструментов небезопасно.')
        span = matches[0]
        chars = span['chars']
        words, current = [], []
        for char in chars:
            if char['c'].isspace():
                if current:
                    words.append(current); current = []
            else:
                current.append(char)
        if current:
            words.append(current)
        kept = []
        crossed = False
        removed = False
        for word in words:
            x0 = min(c['bbox'][0] for c in word)
            x1 = max(c['bbox'][2] for c in word)
            # Italic glyph boxes may overhang a divider even when the word is
            # in the tools column. Classify a WHOLE token by its centre, not
            # the left edge of its first slanted glyph; flag every crossing.
            crosses = x0 < right < x1
            crossed = crossed or crosses
            if (x0 + x1) / 2 >= right:
                removed = True
                continue
            kept.append(word)
        if kept:
            retained_chars = [c for word in kept for c in word]
            r = fitz.Rect(min(c['bbox'][0] for c in retained_chars),
                          min(c['bbox'][1] for c in retained_chars),
                          max(c['bbox'][2] for c in retained_chars),
                          max(c['bbox'][3] for c in retained_chars))
            raw = ' '.join(''.join(c['c'] for c in word) for word in kept)
            # Decode in the context of the original full span, so a short
            # retained Cyrillic word does not lose the established encoding.
            decoded = repair_gost_text(original, span['font'])
            char_index = {id(char): pos for pos, char in enumerate(chars)}
            text = (' '.join(''.join(decoded[char_index[id(c)]] for c in word) for word in kept)
                    if len(decoded) == len(chars) else repair_gost_text(raw, span['font']))
            selected.append((r, text, raw))
        if crossed or removed:
            message = ('Строка пересекает границу инструментов: текст разделён по координатам '
                       'символов. Сверьте требование, размеры и допуски с вырезкой PDF.')
            if crossed:
                message += ' Слово на границе отнесено к колонке по центру; проверьте его принадлежность.'
            reviews.append((rect, message))
    return selected, reviews


def ocr_tessdata_path():
    """Local models only. No network, executable search, or global env changes."""
    override = os.environ.get('KK_OCR_TESSDATA')
    roots = ([Path(override)] if override else
             [bundle_dir() / 'tessdata', Path(sys.executable).resolve().parent / 'tessdata'])
    for root in roots:
        if all((root / (language + '.traineddata')).is_file() for language in ('rus', 'eng')):
            return root
    raise RuntimeError("Локальные модели OCR не найдены. Нужна папка tessdata с "
                       "rus.traineddata и eng.traineddata из полного архива программы.")


def recognize_ocr_image(path, tessdata=None):
    """Read a parameter-only crop using MuPDF's built-in Tesseract; return raw OCR."""
    fitz = _fitz()
    models = Path(tessdata) if tessdata is not None else ocr_tessdata_path()
    pix = fitz.Pixmap(str(path))
    if pix.width * pix.height > 20_000_000:
        raise ValueError("Область OCR больше 20 мегапикселей. Нужна меньшая вырезка.")
    if pix.alpha:
        pix = fitz.Pixmap(pix, 0)
    if pix.colorspace is None or pix.colorspace.n != 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
    payload = pix.pdfocr_tobytes(language='rus+eng', tessdata=str(models), compress=True)
    with fitz.open('pdf', payload) as doc:
        text = '\n'.join(page.get_text('text', sort=True).strip() for page in doc).strip()
    if not text:
        raise ValueError("OCR не обнаружило читаемого текста в этой области.")
    # Never repair technical signs or infer confidence from this text.
    return text


def ocr_text_for_editor(text, number):
    """Strip only an exactly matching, punctuated leading item number."""
    text = text.strip()
    first = text.splitlines()[0] if text else ''
    match = ITEM_START.match(first)
    if match and number_identity(match.group(1)) == number_identity(number):
        return '\n'.join([match.group(2)] + text.splitlines()[1:]).strip()
    return text


def recognize_problem_regions(rows, recognizer=None, max_regions=30, seconds=60):
    """Bound auto OCR; failures never cancel native extraction or change its text."""
    recognizer = recognizer or recognize_ocr_image
    started = time.monotonic()
    for n, row in enumerate(rows):
        if not row.get('crop'):
            row.setdefault('text', '')
            row['status'] = row.get('status') or 'Нет безопасной вырезки для OCR.'
            continue
        if n >= max_regions or time.monotonic() - started >= seconds:
            row.update(text='', status='Авто-OCR отложено: лимит обработки. Доступен повтор в редакторе пункта.')
            continue
        try:
            row.update(text=recognizer(row['crop']),
                       status='Непроверенный вариант OCR. Сверьте номера, размеры, знаки и допуски с PDF.')
        except (RuntimeError, ValueError) as exc:  # нет моделей, пустая или слишком большая область
            row.update(text='', status='OCR недоступно или не удалось: ' + str(exc))
        except Exception as exc:  # сбой движка OCR не отменяет обычное чтение PDF
            log.warning('Авто-OCR: непредвиденная ошибка для %s', row.get('crop'), exc_info=True)
            row.update(text='', status='OCR недоступно или не удалось: ' + str(exc))
    return rows


def item_ocr_candidate(item):
    """Do not alter the requirement, its original text, checkbox or warnings."""
    return ocr_text_for_editor(item.ocr_text, item.number)


def ocr_technical_warning(native, candidate, number):
    """Compare visible numeric/sign tokens, not confidence or semantic correctness."""
    def tokens(value):
        value = value.strip()
        # For comparison ONLY, tolerate OCR's lost punctuation after a number.
        match = re.match(r'^(\d+(?:\.\d+)*)[.)]?\s+', value)
        if match and number_identity(match.group(1)) == number_identity(number):
            value = value[match.end():]
        value = value.translate(str.maketrans({'−': '-', '–': '-', '—': '-'}))
        return re.findall(r'\d+(?:[.,]\d+)*|[±+\-/]', value)
    if candidate.strip() and tokens(native) != tokens(candidate):
        return ("ВНИМАНИЕ: числа или знаки OCR отличаются от прочитанного текста PDF. "
                "Вариант нельзя считать исправлением без сверки каждого размера и допуска.")
    return "Совпадение чисел/знаков не гарантирует правильность OCR. Сверьте весь пункт с изображением."


def resolve_cover_identity(identity_pages):
    """Only the first recognized cover supplies identity; conflicts are warnings."""
    warnings = []
    covers = [page for page in identity_pages if page['is_cover']]
    if not covers:
        warnings.append("Титульный лист не найден. Обозначение и наименование не подставлены "
                        "с других листов; заполните их вручную. Документ открыт.")
        return '', '', warnings
    cover = covers[0]
    if len(covers) > 1:
        warnings.append(f"Найдено несколько титульных листов: "
                        + ', '.join(str(page['page']) for page in covers)
                        + f". Источник сведений — первый титульный лист (стр. {cover['page']}).")

    def unique_codes(values):
        result = {}
        for value in values:
            if is_product_code(value):
                result.setdefault(product_code_key(value), value)
        return list(result.values())

    codes = unique_codes(cover['codes'])
    names = list(dict.fromkeys(value.strip() for value in cover['names'] if value.strip()))
    code = codes[0] if len(codes) == 1 else ''
    name = names[0] if len(names) == 1 else ''
    for label, values in [('Обозначение', codes), ('Наименование', names)]:
        if not values:
            warnings.append(f"{label} на титульном листе (стр. {cover['page']}) не найдено; "
                            "заполните поле вручную.")
        elif len(values) > 1:
            warnings.append(f"{label}: на титульном листе (стр. {cover['page']}) "
                            "найдены разные значения: " + '; '.join(values)
                            + ". Поле оставлено пустым; заполните его вручную. Документ открыт.")
    for page in identity_pages:
        if page is cover:
            continue
        other_codes = unique_codes(page['codes'])
        other_names = list(dict.fromkeys(value.strip() for value in page['names'] if value.strip()))
        for label, expected, values, key in [
                ('обозначение', code, other_codes, product_code_key),
                ('наименование', name, other_names, lambda value: value)]:
            if expected:
                different = [value for value in values if key(value) != key(expected)]
                if different:
                    warnings.append(f"Страница {page['page']}: {label} "
                                    + '; '.join(different)
                                    + f" отличается от титульного листа (стр. {cover['page']}): {expected}. "
                                    "Сохранено значение с титульного листа; документ открыт.")
            elif values:
                warnings.append(f"Страница {page['page']}: {label} " + '; '.join(values)
                                + " не использовано: источник сведений — только титульный лист.")
    return code, name, warnings


@dataclass
class _Extraction:
    """Состояние разбора одного PDF, общее для всех страниц."""
    ocr: bool = False
    warnings: list = field(default_factory=list)
    graphic_rows: dict = field(default_factory=dict)
    ocr_rows: list = field(default_factory=list)
    identity_pages: list = field(default_factory=list)
    page_types: list = field(default_factory=list)
    raw_by_page: dict = field(default_factory=dict)
    profile_pages: int = 0
    crop_dir: Path | None = None

    def evidence_dir(self) -> Path:
        """Папка вырезок создаётся только когда она действительно нужна."""
        if self.crop_dir is None:
            self.crop_dir = make_temp_dir("pdf_checklist_evidence_")
        return self.crop_dir


def _page_spans(page, fitz) -> list:
    spans = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                text = repair_gost_text(span["text"], span["font"])
                spans.append((fitz.Rect(span["bbox"]), text, span["text"]))
    # Deduplicate repeated template text, not equal text at other positions.
    return list({(tuple(round(v, 2) for v in r), t): (r, t, raw)
                 for r, t, raw in spans}.values())


def _page_rules(drawings, page_width: float) -> tuple[list, list]:
    """Long horizontal table rules and all vertical rules of the page."""
    horizontals, verticals = [], []
    for drawing in drawings:
        for segment in drawing["items"]:
            if segment[0] != "l":
                continue
            a, b = segment[1:3]
            if abs(a.y - b.y) < .2 and abs(a.x - b.x) > page_width * .8:
                horizontals.append(a.y)
            if abs(a.x - b.x) < .2:
                verticals.append((a.x, min(a.y, b.y), max(a.y, b.y)))
    return horizontals, verticals


def _table_bounds(index, spans, horizontals, verticals, heading, tools):
    """Borders of the parameter column: (row boundaries, left, right, top, bottom)."""
    boundaries = sorted(set(round(y, 2) for y in horizontals if y > heading.y1))
    lefts = [x for x, y0, y1 in verticals
             if x < heading.x0 and y0 <= heading.y0 + 2 and y1 >= heading.y1]
    rights = [x for x, y0, y1 in verticals
              if heading.x1 < x < tools.x0 and y0 <= tools.y0 + 2 and y1 >= tools.y1]
    # The parameter text overflows into the empty code column in this PDF.
    # The last divider before the tools heading is the safe right boundary.
    if not boundaries or not lefts or not rights:
        raise ValueError(f"Страница {index}: заголовки найдены, но границы таблицы не подтверждены.")
    left, right = max(lefts), max(rights)
    top = boundaries[0]
    footer = [r.y0 for r, t, _ in spans if t.strip() in ("Технический контроль", "Операционная карта контроля") and r.y0 > top]
    if not footer:
        raise ValueError(f"Страница {index}: нижняя граница таблицы не подтверждена.")
    bottom = max(y for y in boundaries if y < min(footer))
    return boundaries, left, right, top, bottom


def _table_spans(page, index, spans, boundaries, left, right, top, bottom):
    """Spans of the parameter column and boundary-crossing reviews."""
    possible = [(r, t, raw) for r, t, raw in spans
                if left <= r.x0 < right and top <= (r.y0 + r.y1) / 2 < bottom]
    starts = [r.y0 for r, t, _ in possible if ITEM_START.match(t)]
    first_item_y = min(starts) if starts else bottom
    # Include upper tolerance values in the first item's physical table band.
    first_band = max((v for v in boundaries if v <= first_item_y), default=top)
    chosen = [(r, t, raw) for r, t, raw in possible
              if (r.y0 + r.y1) / 2 >= first_band and t.strip()]
    tool_starts = [r.x0 for r, t, _ in spans if t.strip() and r.x0 >= right
                   and top <= (r.y0 + r.y1) / 2 < bottom]
    # Allow small glyph-box overhang only into the empty gap before instruments.
    safe_right = min(right + 5, min(tool_starts) - 2) if tool_starts else right + 1
    try:
        return split_parameter_spans(page, chosen, right, safe_right)
    except ValueError as exc:
        raise ValueError(f"Страница {index}: {exc}") from exc


def _attach_vector_signs(fitz, drawings, area, rows) -> list:
    """Add vector-drawn «±» signs to their text rows; return recognised drawings."""
    recognized_signs = []
    for drawing in drawings:
        rect = drawing["rect"]
        if vector_plus_minus(drawing) and area.contains(rect) and rows:
            row = min(rows, key=lambda row: abs(row[0] + 7 - (rect.y0 + rect.y1) / 2))
            if abs(row[0] + 7 - (rect.y0 + rect.y1) / 2) <= 15:
                base = max(row[1], key=lambda v: v[0].height)[0]
                row[1].append((fitz.Rect(rect.x0, base.y0, rect.x1, base.y1), "±", ""))
                recognized_signs.append(drawing)
    return recognized_signs


def _merge_row_warning(state: _Extraction, index: int, number: str, messages: list[str]) -> str:
    """One warning per item: merge with the item's previous rows."""
    previous = state.graphic_rows.get((index, number), ('', ''))[0]
    warning = '\n'.join(dict.fromkeys((previous.splitlines() if previous else []) + messages))
    entry = f"Страница {index}, пункт {number or '?'}: {warning}"
    previous_entry = f"Страница {index}, пункт {number or '?'}: {previous}"
    if previous_entry in state.warnings:
        state.warnings.remove(previous_entry)
    if entry not in state.warnings:
        state.warnings.append(entry)
    return warning


def _prepare_ocr_crops(state: _Extraction, index, page, fitz, evidence_rects, parameter_rects) -> None:
    # OCR gets only the confirmed parameter area, NOT the
    # full-width verification crop that may contain tools.
    for key in evidence_rects:
        if key not in parameter_rects:
            continue
        clean = parameter_rects[key]
        crop = state.evidence_dir() / f"ocr_page_{index}_item_{key[1].replace('.', '_')}.png"
        scale = min(300 / 72, (18_000_000 / max(1, clean.get_area())) ** .5)
        entry = {'page': index, 'number': key[1], 'crop': '', 'text': '', 'status': ''}
        try:
            page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clean, alpha=False).save(str(crop))
            entry['crop'] = str(crop)
        except Exception as exc:  # MuPDF/диск: пункт остаётся без OCR, разбор продолжается
            log.warning("Страница %s, пункт %s: область OCR не подготовлена", index, key[1], exc_info=True)
            entry['status'] = 'Не удалось подготовить область OCR: ' + str(exc)
        state.ocr_rows.append(entry)


def _extract_control_page(state: _Extraction, index, page, fitz, spans, drawings,
                          horizontals, verticals, heading, tools) -> list[str]:
    """Read one control sheet; return its text lines."""
    boundaries, left, right, top, bottom = _table_bounds(index, spans, horizontals, verticals, heading, tools)
    area = fitz.Rect(left, top, right, bottom)
    chosen, boundary_reviews = _table_spans(page, index, spans, boundaries, left, right, top, bottom)
    rows = parameter_lines(chosen)
    recognized_signs = _attach_vector_signs(fitz, drawings, area, rows)
    evidence_right = max(right, max((r.x1 for r, _, _ in chosen), default=right))
    page_lines = []
    active_number = ""
    evidence_rects = {}
    parameter_rects = {}
    for y, parts in rows:
        parts.sort(key=lambda s: s[0].x0)
        text = " ".join(t for _, t, _ in parts)
        text = _JOINED_PLUS_MINUS.sub("±", text)
        page_lines.append(text)
        match = ITEM_START.match(text)
        if match:
            active_number = match.group(1)
        number = active_number
        broken_encoding = any("�" in t or _MOJIBAKE.search(t) for _, t, _ in parts)
        if broken_encoding:
            state.warnings.append(f"Страница {index}, пункт {number or '?'}: проверьте нераспознанные символы.")
        row_top = max(v for v in boundaries if v <= y)
        row_bottom = min(v for v in boundaries if v > y)
        row_rect = fitz.Rect(left, row_top, evidence_right, row_bottom)
        if number:
            key = (index, number)
            clean_rect = fitz.Rect(left, row_top, right, row_bottom)
            parameter_rects[key] = parameter_rects.get(key, clean_rect) | clean_rect
        marks = [d for d in drawings if d["rect"].width > 1 and d["rect"].height > 1
                 and row_rect.contains(d["rect"]) and d not in recognized_signs]
        row_reviews = [message for r, message in boundary_reviews
                       if row_top <= (r.y0 + r.y1) / 2 < row_bottom]
        messages = list(dict.fromkeys(row_reviews))
        if broken_encoding:
            messages.append("Есть неоднозначные символы кодировки. Сверьте текст с вырезкой PDF.")
        if _AMBIGUOUS_VIEW_LETTER.search(text):
            messages.append("Неоднозначная буква в ссылке на вид/разрез: сверьте исходный пункт PDF.")
        if any(t.strip() in ('+', '-', '−') for _, t, _ in parts):
            messages.append("Знак допуска записан отдельно: сверьте порядок и значение допуска с вырезкой PDF.")
        if marks:
            messages.append("В исходном пункте есть графическое обозначение: оно не представлено в тексте. Проверьте вырезку PDF.")
        if not messages:
            continue
        warning = _merge_row_warning(state, index, number, messages)
        crop = state.evidence_dir() / f"page_{index}_item_{number.replace('.', '_') or len(page_lines)}.png"
        evidence_top = min((r.y0 for r, t, raw in chosen
                            if ITEM_START.match(t) and ITEM_START.match(t).group(1) == number),
                           default=row_top)
        evidence_top = max(v for v in boundaries if v <= min(row_top, evidence_top))
        # Include full width for column-crossing review, otherwise
        # keep the previous compact graphical-sign evidence.
        crop_rect = fitz.Rect(left, evidence_top,
                              page.rect.width - 5 if row_reviews else evidence_right,
                              row_bottom)
        key = (index, number)
        if key in evidence_rects:
            crop_rect |= evidence_rects[key]
        evidence_rects[key] = fitz.Rect(crop_rect)
        page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=crop_rect).save(str(crop))
        state.graphic_rows[(index, number)] = (warning, str(crop))
    if state.ocr:
        _prepare_ocr_crops(state, index, page, fitz, evidence_rects, parameter_rects)
    return page_lines


def _extract_page(state: _Extraction, index: int, page, fitz) -> None:
    spans = _page_spans(page, fitz)
    parameter_headers = [r for r, t, _ in spans if t == "Контролируемые параметры"]
    tool_headers = [r for r, t, _ in spans if t == "Наименование средств ТО"]
    drawings = page.get_drawings()
    page_codes, page_names = read_identity(spans, drawings, page.rect)
    is_sketch = any(" ".join(t.split()) == "Карта эскизов контроля" for _, t, _ in spans)
    is_cover = any(" ".join(t.split()) == "Титульный лист" for _, t, _ in spans)
    is_first_control_sheet = any(" ".join(t.split()) == "Наименование операции" for _, t, _ in spans)
    # Non-cover identities are used only for non-fatal discrepancy messages.
    # Sketch headers may reference unrelated assemblies; do not compare them.
    if not is_sketch:
        state.identity_pages.append({"page": index, "is_cover": is_cover,
                                     "codes": page_codes,
                                     "names": page_names if is_cover or is_first_control_sheet else []})
    if parameter_headers and tool_headers:
        horizontals, verticals = _page_rules(drawings, page.rect.width)
        page_lines = _extract_control_page(state, index, page, fitz, spans, drawings, horizontals,
                                           verticals, parameter_headers[0], tool_headers[0])
        state.profile_pages += 1
        state.page_types.append({"page": index, "type": "control"})
        state.raw_by_page[index] = "\n".join(page_lines)
    elif is_cover:
        state.page_types.append({"page": index, "type": "cover"})
        state.raw_by_page[index] = ""
    elif is_sketch:
        state.page_types.append({"page": index, "type": "sketch"})
        state.raw_by_page[index] = ""
    else:
        state.page_types.append({"page": index, "type": "unrecognized"})
        state.raw_by_page[index] = page.get_text("text", sort=True)


def _page_chunks(state: _Extraction) -> list[str]:
    """Do not let cover/unknown-page text become a continuation of a requirement."""
    chunks = []
    for entry in state.page_types:
        index = entry["page"]
        payload = state.raw_by_page[index]
        if entry["type"] == "unrecognized":
            if state.profile_pages:
                payload = ""
                state.warnings.append(f"Страница {index}: не распознана как лист контроля; не включена в чек-лист. Проверьте, нет ли на ней нужных требований.")
            else:
                state.warnings.append(f"Страница {index}: формат карты не распознан; использован прежний алгоритм. Проверьте колонки и кодировку.")
        chunks.append(f"<<<PDF_PAGE:{index}>>>\n" + payload)
    return chunks


def extract_pdf_details(path: str, ocr: bool = False) -> dict:
    """Recognize this control-card layout using headings and vector borders.

    Unknown layouts retain the previous text extraction method with a warning.
    Evidence crops are created only for parameter rows containing vector marks.
    """
    fitz = _fitz()
    state = _Extraction(ocr=ocr)
    with fitz.open(path) as doc:
        if doc.needs_pass:
            raise ValueError("PDF защищён паролем. Откройте незашифрованную копию.")
        for index, page in enumerate(doc, 1):
            _extract_page(state, index, page, fitz)
        chunks = _page_chunks(state)
    document_code, document_name, identity_warnings = resolve_cover_identity(state.identity_pages)
    state.warnings.extend(identity_warnings)
    result = {"raw": "\n".join(chunks), "document_code": document_code,
              "warnings": state.warnings, "identity_warnings": identity_warnings,
              "graphic_rows": state.graphic_rows,
              "control_card": state.profile_pages > 0, "document_name": document_name,
              "page_types": state.page_types}
    if ocr:
        result['ocr_rows'] = recognize_problem_regions(state.ocr_rows)
        if state.ocr_rows:
            count = sum(bool(row['text']) for row in state.ocr_rows)
            result['warnings'].append(
                f"Резервное OCR: вариантов {count} из {len(state.ocr_rows)} областей. "
                "Текст пунктов НЕ заменён. Откройте проблемный пункт → «OCR…» "
                "и сверьте вариант с исходной вырезкой перед применением.")
    log.info("PDF %s: листов контроля %s, предупреждений %s",
             path, state.profile_pages, len(state.warnings))
    return result


def extract_pdf(path: str) -> str:
    """Compatibility wrapper; GUI uses details including verification warnings."""
    return extract_pdf_details(path)["raw"]


def build_card_items(details: dict, raw: str | None = None, every_line: bool = False) -> list[Item]:
    items = parse_items(details["raw"] if raw is None else raw, every_line)
    for item in items:
        item.original_text = item.text
        if details["control_card"]:
            item.text = remove_drawing_references(item.text)
        for page in item.pages:
            evidence = details["graphic_rows"].get((page, item.number))
            if evidence:
                item.warnings.append(evidence[0])
                item.source_crop = evidence[1]
        rows = [row for row in details.get('ocr_rows', [])
                if row['page'] in item.pages and row['number'] == item.number]
        item.ocr_crops = [row['crop'] for row in rows if row.get('crop')]
        item.ocr_text = '\n'.join(row['text'] for row in rows if row.get('text'))
        item.ocr_status = '\n'.join(dict.fromkeys(row.get('status', '') for row in rows))
    return items


def parse_items(raw: str, every_line: bool = False) -> list[Item]:
    """Numbered mode: non-numbered lines after an item are continuations.

    Text before the first number is ignored. Page markers provide provenance.
    This is a heuristic: footers/headings can require manual cleanup.
    """
    result: list[Item] = []
    page = 1
    for line in raw.splitlines():
        marker = PAGE_MARKER.fullmatch(line.strip())
        if marker:
            page = int(marker.group(1))
            continue
        if not line.strip():
            continue
        match = ITEM_START.match(line)
        if every_line:
            number, text = match.groups() if match else (str(len(result) + 1), line.strip())
            result.append(Item(number=number, text=text, pages=[page]))
        elif match:
            number, text = match.groups()
            result.append(Item(number=number, text=text, pages=[page]))
        elif result:
            result[-1].text += "\n" + line.strip()
            if page not in result[-1].pages:
                result[-1].pages.append(page)
    return result
