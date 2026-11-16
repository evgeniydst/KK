"""Формирование паспорта в Word: титульный лист и таблицы выбранных пунктов.

Разметка таблиц в два столбца рассчитывается заранее (plan_free_control_layout).
По желанию результат проверяется в настольном Microsoft Word
(finalize_passport_word_layout, только Windows + pywin32).
По галочке проекта рядом с DOCX сохраняется PDF (export_pdf: Word или LibreOffice).
"""
from __future__ import annotations

import logging
import base64
import io
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Mm, Pt
    from docx.enum.section import WD_SECTION_START, WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:  # понятное сообщение — в _require_docx(), окно программы работает и без Word
    Document = None

from .cover_image import decode_image, fit_size, PICTURE_WIDTH_MM, PICTURE_PADDING_X_MM, PICTURE_PADDING_Y_MM
from .models import ControlCard, duplicate_numbers, ordered_items
from .morphology import decline_product_title
from .paths import ASSETS_DIR, PIP_INSTALL
from .project import WORD_LAYOUT_MODES, cover_types_layout, default_project_settings, validate_project_settings
from .table_settings import BUILTIN_TABLE_LAYOUT

log = logging.getLogger(__name__)

LOGO_PATH = ASSETS_DIR / 'dst_logo.png'
CONTROL_COLUMN_WIDTHS = list(BUILTIN_TABLE_LAYOUT['widths'])  # мм; две колонки по 134.5 мм
HEADER_ROWS = 3  # «Сварочные операции», заголовки столбцов, № и обозначение КК
LAYOUT_BACKENDS = ('auto', 'word', 'native')
PDF_BACKENDS = ('auto', 'word', 'libreoffice')
PDF_TIMEOUT_SECONDS = 180


def _require_docx() -> None:
    if Document is None:
        raise RuntimeError(f"Для экспорта Word установите: {PIP_INSTALL} python-docx")


def _logo_path() -> Path:
    if not LOGO_PATH.is_file():
        raise RuntimeError(f"Не найден логотип {LOGO_PATH}. Скопируйте папку passport_macro целиком.")
    return LOGO_PATH


# ---------------------------------------------------------------- проверка в Microsoft Word

def _passport_frame_is_orphan(header_frame, first_data_frame):
    """Only the header/first-data boundary matters; later rows may flow freely."""
    return header_frame != first_data_frame


def _passport_word_frame(document, source_range, at_end=False):
    """Return actual (physical page, newspaper column) after Word pagination."""
    point = source_range.Duplicate
    if at_end:
        text = str(point.Text or "")
        trim = len(text) - len(text.rstrip("\r\x07"))
        point.SetRange(max(point.Start, point.End - trim - 1),
                       max(point.Start, point.End - trim - 1))
    else:
        point.Collapse(1)  # wdCollapseStart
    window = document.ActiveWindow
    window.ScrollIntoView(point, True)
    page = int(point.Information(3))  # wdActiveEndPageNumber, not printed page number
    x = float(point.Information(5))   # wdHorizontalPositionRelativeToPage, points
    if x < 0:
        point.Select()
        window.ScrollIntoView(point, True)
        x = float(document.Application.Selection.Information(5))
    if x < 0 or page < 1:
        raise RuntimeError("Word не смог определить положение строки на странице.")
    setup = point.Sections.Item(1).PageSetup
    midpoint = (float(setup.PageWidth) + float(setup.LeftMargin) -
                float(setup.RightMargin)) / 2
    return page, (1 if x < midpoint else 2)


def _passport_header_range(table):
    return table.Rows.Item(1).Cells.Item(1).Range.Paragraphs.Item(1).Range


def _passport_word_tables(document):
    """Only control tables, not the cover tables."""
    result = []
    for index in range(1, document.Tables.Count + 1):
        table = document.Tables.Item(index)
        if int(table.NestingLevel) != 1 or table.Rows.Count < HEADER_ROWS + 2:
            continue
        first_cell = table.Rows.Item(1).Cells.Item(1)
        if str(first_cell.Range.Text).lstrip("\r\x07 \t\n").startswith("Сварочные операции"):
            result.append(table)
    return result


def _passport_verify_word_tables(document):
    """Allow middle rows to span columns; reject orphan heads/tails or tall rows."""
    for index, table in enumerate(_passport_word_tables(document), 1):
        header = _passport_header_range(table)
        first = table.Rows.Item(HEADER_ROWS + 1).Range
        last = table.Rows.Item(table.Rows.Count - 1).Range
        signatures = table.Rows.Item(table.Rows.Count).Range
        head_frame = _passport_word_frame(document, header)
        first_frame = _passport_word_frame(document, first)
        if _passport_frame_is_orphan(head_frame, first_frame):
            raise RuntimeError(f"Карта {index}: шапка осталась отдельно от первого пункта.")
        if first_frame != _passport_word_frame(document, first, at_end=True):
            raise RuntimeError(f"Карта {index}: первая строка слишком высокая для одной колонки.")
        tail_frame = _passport_word_frame(document, last)
        if (tail_frame != _passport_word_frame(document, last, at_end=True) or
                tail_frame != _passport_word_frame(document, signatures) or
                tail_frame != _passport_word_frame(document, signatures, at_end=True)):
            raise RuntimeError(f"Карта {index}: последний пункт и подписи не уместились вместе. "
                               "Проверьте длину последнего пункта и размер блока подписей.")


def _passport_repair_word_headers(document):
    tables = _passport_word_tables(document)
    if not tables:
        raise RuntimeError("Word не обнаружил контрольные таблицы в паспорте.")
    moved = 0
    for index, table in enumerate(tables, 1):
        header = _passport_header_range(table)
        first = table.Rows.Item(HEADER_ROWS + 1).Range
        if not _passport_frame_is_orphan(
                _passport_word_frame(document, header),
                _passport_word_frame(document, first)):
            continue
        # A separator paragraph is deliberately created before each
        # control table by export_passport_docx, even the first one.
        anchor = document.Range(table.Range.Start - 1, table.Range.Start - 1)
        anchor.ParagraphFormat.KeepWithNext = 0
        anchor.ParagraphFormat.SpaceBefore = 0
        anchor.ParagraphFormat.SpaceAfter = 0
        anchor.ParagraphFormat.LineSpacingRule = 4  # wdLineSpaceExactly
        anchor.ParagraphFormat.LineSpacing = 1
        anchor.Font.Size = 1
        anchor.InsertBreak(8)  # wdColumnBreak, not a page break
        moved += 1
        document.Repaginate()
        # One move is sufficient unless header + first row is physically
        # taller than a whole column. Never add repeated blank columns.
        if _passport_frame_is_orphan(
                _passport_word_frame(document, _passport_header_range(table)),
                _passport_word_frame(document, table.Rows.Item(HEADER_ROWS + 1).Range)):
            raise RuntimeError(f"Карта {index}: шапка и первый пункт не уместились "
                               "в одной колонке даже после переноса.")
    return moved


def finalize_passport_word_layout(path: str | Path, backend: str = "auto") -> dict:
    """Repair actual orphan heads with Word, not estimated character heights.

    On Windows the default requires desktop Microsoft Word and pywin32. The
    operation uses a temporary copy, disables macros and never attaches to an
    already running Word instance. Only a verified copy replaces the DOCX.
    On other systems native DOCX rules are saved, with an explicit GUI warning.
    """
    if backend not in LAYOUT_BACKENDS:
        raise ValueError("Неизвестный движок проверки разметки.")
    if backend == "native" or (backend == "auto" and os.name != "nt"):
        return {"layout_checked": False, "moved_tables": 0,
                "note": "DOCX создан со свободным переносом строк. "
                        "Фактическая пагинация Word не проверена. Для автоматической "
                        "проверки нужны Windows, настольный Microsoft Word и pywin32."}
    if os.name != "nt":
        raise RuntimeError("Точная проверка разметки Microsoft Word требует Windows.")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("DOCX создан, но финальная проверка переноса не выполнена. "
                           "Установите pywin32: py -m pip install pywin32. "
                           "Затем повторите экспорт. Требуется настольный Microsoft Word.") from exc

    target = Path(path).resolve()
    app = document = None
    initialized = False
    with tempfile.TemporaryDirectory(prefix="passport_layout_", dir=str(target.parent)) as tmp:
        working = Path(tmp) / "passport.docx"
        shutil.copy2(target, working)
        try:
            pythoncom.CoInitialize()
            initialized = True
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            document = app.Documents.Open(str(working), False, False, False)
            document.ActiveWindow.View.Type = 3  # wdPrintView
            document.Repaginate()
            original_text = str(document.Content.Text).replace("\x0e", "")
            moved = _passport_repair_word_headers(document)
            document.Repaginate()
            _passport_verify_word_tables(document)
            if str(document.Content.Text).replace("\x0e", "") != original_text:
                raise RuntimeError("Проверка остановлена: изменился текст документа.")
            document.Save()
            document.Close(0)
            document = None
            # Word may repaginate during save; re-open once and verify saved file.
            document = app.Documents.Open(str(working), False, True, False)
            document.ActiveWindow.View.Type = 3
            document.Repaginate()
            _passport_verify_word_tables(document)
            if str(document.Content.Text).replace("\x0e", "") != original_text:
                raise RuntimeError("После сохранения Word изменился текст документа.")
            document.Close(0)
            document = None
            os.replace(working, target)
            log.info('Проверено в Word: %s, перенесено шапок: %d', target, moved)
            return {"layout_checked": True, "moved_tables": moved,
                    "note": "Проверено в Microsoft Word: шапки с первыми пунктами, "
                            "последние пункты с подписями. Остальные строки переносятся свободно."}
        except Exception as exc:
            raise RuntimeError("Черновой DOCX сохранён, но финальная проверка Microsoft Word "
                               f"не завершена: {exc}") from exc
        finally:
            if document is not None:
                try:
                    document.Close(0)
                except Exception:  # Word уже закрыл документ или завис
                    log.debug('Word: не удалось закрыть документ', exc_info=True)
            if app is not None:
                try:
                    app.Quit(0)
                except Exception:
                    log.debug('Word: не удалось завершить процесс', exc_info=True)
            if initialized:
                pythoncom.CoUninitialize()


def find_libreoffice() -> Path | None:
    """soffice из PATH или из стандартной папки установки Windows."""
    for name in ('soffice', 'soffice.exe', 'libreoffice'):
        found = shutil.which(name)
        if found:
            return Path(found)
    if os.name == 'nt':
        for base in (os.environ.get('ProgramFiles'), os.environ.get('ProgramFiles(x86)')):
            if base:
                candidate = Path(base) / 'LibreOffice' / 'program' / 'soffice.exe'
                if candidate.is_file():
                    return candidate
    return None


def _replace_file(produced: Path, target: Path) -> None:
    try:
        os.replace(produced, target)
    except PermissionError as exc:
        raise RuntimeError(f'Не удалось заменить {target.name}: закройте файл в программе '
                           'просмотра и повторите.') from exc


def _pdf_with_word(source: Path, target: Path) -> None:
    """Word экспортирует локальную рабочую копию; исходный DOCX не изменяется."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(f'для PDF через Word установите pywin32: {PIP_INSTALL} pywin32') from exc
    app = document = None
    initialized = False
    # Avoid making Word export into a network share and avoid opening the
    # original in read-only mode. Some Word installations reject that export.
    with tempfile.TemporaryDirectory(prefix='passport_pdf_') as tmp:
        working = Path(tmp) / 'passport.docx'
        produced = Path(tmp) / 'passport.pdf'
        try:
            shutil.copyfile(source, working)
            pythoncom.CoInitialize()
            initialized = True
            app = win32com.client.DispatchEx('Word.Application')
            app.Visible = False
            app.DisplayAlerts = 0
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            document = app.Documents.Open(str(working), False, False, False)
            document.ExportAsFixedFormat(str(produced), 17, False, 0)  # wdExportFormatPDF, для печати
            document.Close(0)
            document = None
        except Exception as exc:  # ошибки COM не наследуют RuntimeError
            raise RuntimeError(f'Word не сохранил PDF: {exc}') from exc
        finally:
            if document is not None:
                try:
                    document.Close(0)
                except Exception:
                    log.debug('Word: не удалось закрыть документ', exc_info=True)
            if app is not None:
                try:
                    app.Quit(0)
                except Exception:
                    log.debug('Word: не удалось завершить процесс', exc_info=True)
            if initialized:
                pythoncom.CoUninitialize()
        if not produced.is_file():
            raise RuntimeError('Word не создал PDF.')
        # Stage on the destination filesystem before replacing an existing PDF:
        # os.replace across a local disk and a network share is not supported.
        staged = None
        try:
            with tempfile.NamedTemporaryFile(prefix='passport_pdf_', suffix='.tmp',
                                             dir=str(target.parent), delete=False) as f:
                staged = Path(f.name)
            shutil.copyfile(produced, staged)
            _replace_file(staged, target)
        except OSError as exc:
            raise RuntimeError(f'Не удалось записать PDF в папку назначения: {exc}') from exc
        finally:
            if staged is not None:
                staged.unlink(missing_ok=True)


def _pdf_with_libreoffice(source: Path, target: Path) -> None:
    """PDF через LibreOffice без окна; отдельный профиль не мешает открытому LibreOffice."""
    soffice = find_libreoffice()
    if soffice is None:
        raise RuntimeError('LibreOffice не найден')
    with tempfile.TemporaryDirectory(prefix='passport_pdf_') as tmp:
        out_dir = Path(tmp) / 'out'
        command = [str(soffice), '-env:UserInstallation=' + (Path(tmp) / 'profile').as_uri(),
                   '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', str(out_dir), str(source)]
        flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        try:
            completed = subprocess.run(command, stdin=subprocess.DEVNULL, capture_output=True,
                                       timeout=PDF_TIMEOUT_SECONDS, creationflags=flags)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise RuntimeError(f'LibreOffice не сохранил PDF: {exc}') from exc
        produced = out_dir / (source.stem + '.pdf')
        if not produced.is_file():
            details = (completed.stderr or completed.stdout or b'').decode('utf-8', 'replace').strip()
            raise RuntimeError('LibreOffice не создал PDF' + (f': {details[-300:]}' if details else '.'))
        staged = target.with_name('.' + target.name + '.tmp')
        shutil.copyfile(produced, staged)
        try:
            _replace_file(staged, target)
        finally:
            staged.unlink(missing_ok=True)


def export_pdf(docx_path: str | Path, pdf_path: str | Path | None = None, backend: str = 'auto') -> Path:
    """Сохранить PDF из готового DOCX (по умолчанию — рядом, с тем же именем).

    auto: в Windows сначала Microsoft Word (точнее всего совпадает с печатью), затем LibreOffice.
    Старый PDF заменяется только готовым новым файлом.
    """
    if backend not in PDF_BACKENDS:
        raise ValueError('Неизвестный способ сохранения PDF.')
    source = Path(docx_path).resolve()
    target = Path(pdf_path).resolve() if pdf_path else source.with_suffix('.pdf')
    errors = []
    if backend in ('auto', 'word'):
        if os.name == 'nt':
            try:
                _pdf_with_word(source, target)
                log.info('PDF сохранён через Word: %s', target)
                return target
            except RuntimeError as exc:
                if backend == 'word':
                    raise
                errors.append(str(exc))
        elif backend == 'word':
            raise RuntimeError('PDF через Microsoft Word сохраняется только в Windows.')
    try:
        _pdf_with_libreoffice(source, target)
        log.info('PDF сохранён через LibreOffice: %s', target)
        return target
    except RuntimeError as exc:
        if backend == 'libreoffice':
            raise
        errors.append(str(exc))
    raise RuntimeError('Не удалось сохранить PDF. Подробности: ' + '; '.join(errors))


SPLIT_MIN_LINES = 5      # rows of 5+ estimated lines may continue in the next column
SPLIT_LINE_HEIGHT = 12.5  # pt, same line height as estimate_control_parts


def passport_item_text(text: str) -> str:
    """Точечная замена формулировки только в сформированном паспорте."""
    if ' '.join(text.split()).casefold() == 'клеймить ударным способом.':
        return 'Клеймо ОТК.'
    return text


def plan_free_control_layout(column_used, capacity, header, rows, signatures,
                             has_previous=False, gap=0.0, splittable=None):
    """Conservative native-DOCX planning, without keeping all data rows together.

    A card may start in the remainder when its header + first item fits there.
    A single-item card must also keep its signatures there. Predict
    its final occupied column row-by-row, including repeated headers and the
    protected last-item/signatures block. This is a size estimate, not Word's
    actual pagination; optional Word verification remains available separately.
    """
    if capacity <= 0 or not rows:
        raise ValueError("Нет места для контрольной таблицы или отсутствуют пункты.")
    first_block = header + rows[0] + (signatures if len(rows) == 1 else 0)
    required = first_block
    start_used = column_used + gap if has_previous else 1.0
    break_before = bool(has_previous and start_used + required > capacity)
    if break_before:
        start_used = 1.0  # separator's one-point line after the column break
    used = start_used + header
    for index, height in enumerate(rows):
        block = height + (signatures if index == len(rows) - 1 else 0)
        if (used + block > capacity and splittable and splittable[index]
                and capacity - used >= 2 * SPLIT_LINE_HEIGHT):
            # A long middle row fills the remainder of this column and
            # continues after the repeated header in the next one.
            used = min(capacity, header + used + block - capacity)
            continue
        if used + block > capacity:
            used = header  # native continuation starts with a repeated header
        if used + block > capacity:
            # An exceptionally tall row cannot be predicted safely. Do not
            # claim free space for the next card, or introduce an infinite loop.
            used = capacity
        else:
            used += block
    return break_before, used


# ---------------------------------------------------------------- оформление

def configure_section(section, columns=1):
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Mm(297), Mm(210)
    section.top_margin = section.bottom_margin = Mm(10)
    section.left_margin = section.right_margin = Mm(10)
    section.header_distance = section.footer_distance = Mm(4)
    cols = section._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols"); section._sectPr.append(cols)
    cols.set(qn("w:num"), str(columns))
    cols.set(qn("w:space"), str(int(Mm(8).twips)))
    cols.set(qn("w:equalWidth"), "1")

def style_run(run, font="Times New Roman", size=10, bold=False, underline=False):
    run.font.name = font
    fonts = run._element.get_or_add_rPr().rFonts
    for script in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn("w:" + script), font)
    run.font.size = Pt(size)
    run.bold = bold
    run.underline = underline
    run.italic = font == "GOST type A"

def set_text(cell, text, bold=False, centered=False, size=10,
             font="Times New Roman", underline=False):
    cell.text = text
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(1)
        paragraph.paragraph_format.space_after = Pt(1)
        paragraph.paragraph_format.keep_with_next = False
        paragraph.paragraph_format.widow_control = True
        if centered:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            style_run(run, font, size, bold, underline)

def table_options(table, widths, compact=False):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    props = table._tbl.tblPr
    width = props.find(qn("w:tblW"))
    width.set(qn("w:type"), "dxa")
    width.set(qn("w:w"), str(int(Mm(sum(widths)).twips)))
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement("w:" + edge)
        element.set(qn("w:val"), "single"); element.set(qn("w:sz"), "6")
        element.set(qn("w:color"), "000000"); borders.append(element)
    props.append(borders)
    margins = OxmlElement("w:tblCellMar")
    for edge, mm in (("top", 0 if compact else .8), ("bottom", 0 if compact else .8), ("left", 1), ("right", 1)):
        element = OxmlElement("w:" + edge)
        element.set(qn("w:w"), str(int(Mm(mm).twips)))
        element.set(qn("w:type"), "dxa"); margins.append(element)
    props.append(margins)
    for column, width_mm in zip(table.columns, widths):
        column.width = Mm(width_mm)
    for row in table.rows:
        for cell, width_mm in zip(row.cells, widths):
            cell.width = Mm(width_mm)

def protect_row(row, repeated=False, compact=False, allow_split=False, height_mm=None):
    props = row._tr.get_or_add_trPr()
    if not allow_split:
        props.append(OxmlElement("w:cantSplit"))
    if repeated:
        props.append(OxmlElement("w:tblHeader"))
    row.height = Mm((5 if compact else 7) if height_mm is None else height_mm)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

def borderless(table, widths):
    table_options(table, widths)
    borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    for element in borders:
        element.set(qn("w:val"), "nil")

def tiny_paragraph(paragraph):
    paragraph.paragraph_format.line_spacing = Pt(1)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    mark = OxmlElement('w:rPr')
    size = OxmlElement('w:sz'); size.set(qn('w:val'), '2')
    mark.append(size); paragraph._p.get_or_add_pPr().append(mark)

def estimate_control_parts(card: ControlCard, items: list, widths: list,
                           header_min_mm=5, row_min_mm=5) -> tuple[float, list[float], float]:
    """Conservative header, individual row and signature height estimates.

    Use ordinary Times Italic metrics, not the narrower GOST face, so the
    decision remains cautious if the GOST font is substituted. This is a
    fit estimate, not a substitute for the Word pagination engine.
    """
    try:
        import pymupdf as fitz
        measure_font = fitz.Font("Times-Italic")
        def text_width(text):
            return measure_font.text_length(text, fontsize=10) * 1.03
    except (ImportError, RuntimeError, ValueError):
        def text_width(text):
            # Dependency-free fallback in points, deliberately generous.
            return sum(2.8 if ch.isspace() else
                       3.2 if ch in ".,:;!|'ijlI1" else
                       8.5 if ch in "ЖШЩЮМФ@%W" else 5.5 for ch in text)

    def wrapped_lines(text, width_mm):
        limit = max(10, (width_mm - 2) * 72 / 25.4 - 1)
        count = 0
        for logical_line in (text.split("\n") or [""]):
            current = ""
            count += 1
            for word in logical_line.split():
                candidate = (current + " " + word) if current else word
                if text_width(candidate) <= limit:
                    current = candidate
                    continue
                if current:
                    count += 1
                    current = ""
                for character in word:
                    if current and text_width(current + character) > limit:
                        count += 1
                        current = character
                    else:
                        current += character
        return count

    header_minimum = header_min_mm * 72 / 25.4
    row_minimum = row_min_mm * 72 / 25.4
    caption_lines = max(wrapped_lines('Контролируемые параметры', widths[0] + widths[1]),
                        wrapped_lines('Факт', widths[2]),
                        wrapped_lines('ОК/\nНОК', widths[3]),
                        wrapped_lines('Комментарии', widths[4]))
    # Include border allowance and a mandatory tail paragraph in the header.
    header = (max(header_minimum, 13.2) + max(header_minimum, 24.5, 12.5 * caption_lines) +
              max(header_minimum, 12.5 * wrapped_lines(
                  card.document_code.strip() + " " + card.document_name.strip(), widths[1])) +
              3 * .75 + .05)
    rows = [max(row_minimum, 12.5 * max(
        wrapped_lines(item.number, widths[0]),
        wrapped_lines(passport_item_text(item.text), widths[1]))) + .75 for item in items]
    signatures = 43  # three 10 pt lines, 1.15 spacing and top separation
    return header, rows, signatures

def estimate_control_height(card: ControlCard, items: list, widths: list,
                            header_min_mm=5, row_min_mm=5) -> float:
    header, rows, signatures = estimate_control_parts(card, items, widths, header_min_mm, row_min_mm)
    return header + sum(rows) + signatures


# ---------------------------------------------------------------- части документа

def _selected_groups(cards: list[ControlCard]) -> list:
    """(номер КК, КК, отмеченные пункты); проверка до создания файла."""
    groups = [(n, card, [i for i in ordered_items(card.items) if i.checked])
              for n, card in enumerate(cards, 1) if any(i.checked for i in card.items)]
    if not groups:
        raise ValueError("Сначала отметьте пункты для паспорта.")
    for n, card, items in groups:
        repeated = duplicate_numbers(items)
        if repeated:
            raise ValueError(f"КК {n}: повторные номера {', '.join(repeated)}. Исправьте их перед формированием паспорта.")
        if not card.document_code.strip() or not card.document_name.strip():
            raise ValueError(f"КК {n}: заполните обозначение и наименование изделия.")
    return groups


def _new_document(settings: dict):
    doc = Document()
    doc.core_properties.title = settings["name"][:255] or "Паспорт — выбранные пункты карт контроля"
    doc.core_properties.author = ""
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.05
    return doc


def _build_cover(doc, settings: dict) -> None:
    """Титул: слева подписи, справа логотип, утверждение, название и место для рисунка."""
    types_layout = cover_types_layout(settings)
    types_height = types_layout['types_height']

    configure_section(doc.sections[0])
    # Borderless, editable two-half cover: approvals left, passport and picture right.
    cover = doc.add_table(rows=1, cols=2)
    cover.autofit = False
    for column in cover.columns:
        column.width = Mm(138.5)
    for cell in cover.rows[0].cells:
        cell.width = Mm(138.5)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge); el.set(qn("w:val"), "nil"); borders.append(el)
    cover._tbl.tblPr.append(borders)
    left, right = cover.rows[0].cells
    # Two explicit cover zones keep the drawing's bottom at one page position.
    # Only the dividing line (the picture's top) changes with its height.
    cover_right = right
    zones = cover_right.add_table(rows=2, cols=1)
    borderless(zones, [134.5])
    zone_margins = zones._tbl.tblPr.find(qn('w:tblCellMar'))
    for margin in zone_margins:
        margin.set(qn('w:w'), '0')
    trailing_space_mm = 2 * 25.4 / 72  # two mandatory 1 pt paragraphs around the picture
    picture_zone_height = types_layout['picture_height'] + trailing_space_mm
    for row, height in zip(zones.rows, (185 - picture_zone_height, picture_zone_height)):
        row.height = Mm(height)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
    right = zones.cell(0, 0)
    picture_parent = zones.cell(1, 0)
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    picture_parent.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    for paragraph in cover_right.paragraphs:
        tiny_paragraph(paragraph)
    tiny_paragraph(picture_parent.paragraphs[0])
    cover.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))

    # The right-hand fixed zones determine the cover height. Bottom-align
    # signoffs in the same row, so extra people expand the block upwards.
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    p = left.paragraphs[0]
    dates_enabled = settings['cover_dates_enabled']
    people = [("Разработал: " + settings["developer_role"], settings["developer_name"],
               settings['developer_date'] if dates_enabled else '')] + [
        ("Согласовано: " + person["role"], person["name"],
         person.get('date', '') if dates_enabled else '') for person in settings["reviewers"]]
    estimated_height = sum(7 * max(1, (len(role) + 27) // 28, (len(name) + 22) // 23)
                           for role, name, _ in people)
    if estimated_height > 180:
        raise ValueError('Блок подписей слишком большой для одного титульного листа. '
                         'Сократите длинные должности/ФИО или список согласующих.')
    tiny_paragraph(p)
    # Fixed columns keep names, signature labels and dates aligned on all four lines.
    signoffs = left.add_table(rows=len(people), cols=6)
    borderless(signoffs, [49, 32, 9, 21, 8, 16] if dates_enabled else [49, 34, 9, 24, 8, 11])
    # Narrow inner padding leaves enough room for labels and longer signature lines.
    margins = signoffs._tbl.tblPr.find(qn("w:tblCellMar"))
    for edge in ("left", "right"):
        margins.find(qn("w:" + edge)).set(qn("w:w"), str(int(Mm(.2).twips)))
    for row, (role, name, date) in zip(signoffs.rows, people):
        signature_line = "__________" if dates_enabled else "_____________"
        name_line = "______________" if dates_enabled else "__________________"
        for index, (cell, text) in enumerate(zip(row.cells, (role, name or name_line, "подп.",
                                                            signature_line, "Дата", date or "______"))):
            set_text(cell, text, size=9 if dates_enabled and index == 5 else 10)
        row.height = Mm(7)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))

    tiny_paragraph(left.paragraphs[-1])
    # The picture has an extra mandatory 1 pt paragraph in its inner zone.
    # Match that clearance to align the two blocks' bottom edges.
    left.paragraphs[-1].paragraph_format.line_spacing = Pt(2)

    # Logo and approval retain their positions; use standard character width without compression.
    p = right.paragraphs[0]
    p.paragraph_format.line_spacing = Pt(1)
    p.paragraph_format.space_after = Pt(0)
    top = right.add_table(rows=1, cols=2)
    borderless(top, [45, 90])
    logo_p = top.cell(0, 0).paragraphs[0]
    logo_p.paragraph_format.line_spacing = 1
    logo_p.add_run().add_picture(str(_logo_path()), width=Mm(42))
    approval = top.cell(0, 1)
    approval_date = settings['approval_date'] if dates_enabled else ''
    approval_lines = ("Утвердил: " + settings["approval_role"], settings["approval_name"] or "____________________________",
                      "Подпись __________  Дата " + (approval_date or "__________"))
    set_text(approval, "\n".join("  " + line for line in approval_lines),
             font="GOST type A", size=16)
    approval.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for p in approval.paragraphs:
        # The block sits at the right edge; its lines share one LEFT starting point.
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Mm(4)
        p.paragraph_format.right_indent = Mm(0)
        p.paragraph_format.first_line_indent = Mm(0)
        p.paragraph_format.line_spacing = 1.15
    p = right.paragraphs[-1]
    p.paragraph_format.line_spacing = Pt(1)
    p.paragraph_format.space_after = Mm(max(2, 12 - max(0, types_height - 16)))
    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Mm(2)
    style_run(p.add_run("ПАСПОРТ "), "GOST type A", 22)
    style_run(p.add_run(settings["passport_name"] or decline_product_title(settings["name"]) or "[Наименование и обозначение изделия]"), "GOST type A", 22)
    p = right.add_paragraph("Заводской номер: __________________________")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Mm(2)
    for run in p.runs:
        style_run(run, "GOST type A", 14)
    if settings['types_enabled']:
        variants = right.add_table(rows=len(settings['types']), cols=3)
        borderless(variants, [35, 65, 8])
        variants.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Some renderers add cell padding on top of the requested minimum
        # row height. The row itself already reserves breathing room.
        type_margins = variants._tbl.tblPr.find(qn('w:tblCellMar'))
        for edge in ('top', 'bottom'):
            type_margins.find(qn('w:' + edge)).set(qn('w:w'), '0')
        for index, (row, text) in enumerate(zip(variants.rows, settings['types'])):
            set_text(row.cells[0], settings['types_title'] if index == 0 else '', font='GOST type A', size=14)
            set_text(row.cells[1], text, font='GOST type A', size=14)
            for cell in row.cells[:2]:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = Pt(16)
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
            # A real outlined square rather than a font-dependent checkbox glyph.
            square_cell = row.cells[2]
            square_cell.paragraphs[0].paragraph_format.line_spacing = Pt(1)
            square_cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            style_run(square_cell.paragraphs[0].add_run(''), size=1)
            square = square_cell.add_table(rows=1, cols=1)
            table_options(square, [4.5])
            square.rows[0].height = Mm(4.5)
            square.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            square.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
            margins = square._tbl.tblPr.find(qn('w:tblCellMar'))
            for margin in margins:
                margin.set(qn('w:w'), '0')
            for paragraph in (square.cell(0, 0).paragraphs[0], square_cell.paragraphs[-1]):
                paragraph.paragraph_format.line_spacing = Pt(1)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                style_run(paragraph.add_run(''), size=1)
            row.height = Mm(types_layout['row_heights'][index])
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        right.paragraphs[-1].paragraph_format.space_after = Mm(1)
        right.paragraphs[-1].paragraph_format.line_spacing = Pt(1)
        right.paragraphs[-1].add_run().font.size = Pt(1)
    picture = picture_parent.add_table(rows=1, cols=1)
    table_options(picture, [124])
    picture.alignment = WD_TABLE_ALIGNMENT.LEFT
    picture_indent = OxmlElement('w:tblInd')
    picture_indent.set(qn('w:type'), 'dxa')
    picture_indent.set(qn('w:w'), str(int(Mm(5).twips)))
    picture._tbl.tblPr.append(picture_indent)
    picture.rows[0].height = Mm(types_layout['picture_height'])
    picture.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    picture_margins = picture._tbl.tblPr.find(qn('w:tblCellMar'))
    for edge in ('top', 'bottom'):
        picture_margins.find(qn('w:' + edge)).set(qn('w:w'), '0')
    cell = picture.cell(0, 0)
    encoded_image = settings.get('cover_image', '')
    if encoded_image:
        image = decode_image(encoded_image)
        width_mm, height_mm = fit_size(image.width, image.height,
                                       PICTURE_WIDTH_MM - 2 * PICTURE_PADDING_X_MM,
                                       types_layout['picture_height'] - 2 * PICTURE_PADDING_Y_MM)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1
        p.add_run().add_picture(io.BytesIO(base64.b64decode(encoded_image)),
                                width=Mm(width_mm), height=Mm(height_mm))
    else:
        set_text(cell, "[Место для изображения изделия]", centered=True, size=10)
    tiny_paragraph(picture_parent.paragraphs[-1])


def _build_control_tables(doc, groups: list, layout_mode: str, table_layout=None) -> None:
    """Второй раздел: таблицы КК в две колонки.

    free — строки переносятся свободно (длинные средние пункты могут
    продолжаться в соседней колонке); whole — короткие КК целиком.
    """
    section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(section, 2)
    compat = doc.settings.element.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat'); doc.settings.element.append(compat)
    if compat.find(qn('w:noColumnBalance')) is None:
        compat.insert(0, OxmlElement('w:noColumnBalance'))
    table_layout = table_layout or BUILTIN_TABLE_LAYOUT
    widths = table_layout['widths']
    header_min_mm, row_min_mm = table_layout['header_min_mm'], table_layout['row_min_mm']
    column_capacity = section.page_height.pt - section.top_margin.pt - section.bottom_margin.pt - 6
    column_used = 0.0
    previous_table_was_long = False
    for group_index, (n, card, items) in enumerate(groups):
        keep_table_together = False
        start_next_column = False
        splittable = [False] * len(items)
        if layout_mode == 'whole':
            # Previous behavior: pack complete short cards; let oversized cards
            # flow across columns, starting them and the following card fresh.
            estimated_height = estimate_control_height(card, items, widths, header_min_mm, row_min_mm)
            keep_table_together = estimated_height <= column_capacity
            gap = Mm(2).pt + 1
            start_next_column = bool(group_index and (
                previous_table_was_long or not keep_table_together or
                column_used + gap + estimated_height > column_capacity))
            if start_next_column:
                column_used = 0.0
            elif group_index:
                column_used += gap
            column_used += estimated_height
            previous_table_was_long = not keep_table_together
        else:
            # Start a new card when its header and first item fit; later rows
            # can continue in the next column even if the whole card is short.
            header_height, row_heights, signature_height = estimate_control_parts(
                card, items, widths, header_min_mm, row_min_mm)
            # The first row stays with the header and the last with signatures;
            # long rows in between may continue in the next column instead of
            # leaving the bottom of the current column empty.
            splittable = [0 < index < len(items) - 1 and
                          height >= SPLIT_MIN_LINES * SPLIT_LINE_HEIGHT
                          for index, height in enumerate(row_heights)]
            start_next_column, column_used = plan_free_control_layout(
                column_used, column_capacity, header_height, row_heights,
                signature_height, has_previous=bool(group_index), gap=Mm(2).pt + 1,
                splittable=splittable)
        # The separator is also a safe anchor for optional Word orphan repair.
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Mm(2) if group_index and not start_next_column else Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(1)
        p.paragraph_format.keep_with_next = not start_next_column
        run = p.add_run()
        run.font.size = Pt(1)
        if start_next_column:
            run.add_break(WD_BREAK.COLUMN)
        # The three header lines are ordinary rows of the same table, so all
        # vertical lines come from one column grid in Word and LibreOffice.
        # (A nested header table was shifted by the outer cell's side border.)
        # Repeated, unsplittable rows with keep-with-next stay together and
        # attached to the first requirement.
        table = doc.add_table(rows=HEADER_ROWS, cols=5)
        table_options(table, widths, compact=True)
        # Explicit zero indent: LibreOffice then places the frame like Word
        # (compatibility mode 14 moves it left by the cell margin).
        indent = OxmlElement('w:tblInd')
        indent.set(qn('w:w'), '0')
        indent.set(qn('w:type'), 'dxa')
        anchor = table._tbl.tblPr.find(qn('w:jc'))
        if anchor is None:
            anchor = table._tbl.tblPr.find(qn('w:tblW'))
        anchor.addnext(indent)
        look = table._tbl.tblPr.find(qn('w:tblLook'))
        if look is not None:
            look.set(qn('w:val'), '0000')
            for flag in ('firstRow', 'lastRow', 'firstColumn', 'lastColumn'):
                look.set(qn('w:' + flag), '0')
        title_row, caption_row, card_row = table.rows[:HEADER_ROWS]
        set_text(title_row.cells[0].merge(title_row.cells[4]), "Сварочные операции", bold=True, size=11)
        set_text(caption_row.cells[0].merge(caption_row.cells[1]), "Контролируемые параметры", bold=True)
        for index, text in ((2, "Факт"), (3, "ОК/\nНОК"), (4, "Комментарии")):
            set_text(caption_row.cells[index], text, bold=True, centered=True)
        set_text(card_row.cells[0], f"№{n}", centered=True, font="GOST type A")
        set_text(card_row.cells[1], card.document_code.strip() + " " + card.document_name.strip(), font="GOST type A", underline=True)
        for row in table.rows[:HEADER_ROWS]:
            protect_row(row, repeated=True, compact=True, height_mm=header_min_mm)
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True
                    paragraph.paragraph_format.keep_together = True
        for item, may_split in zip(items, splittable):
            row = table.add_row()
            for cell, mm in zip(row.cells, widths):
                cell.width = Mm(mm)
            set_text(row.cells[0], item.number, centered=True, font="GOST type A")
            set_text(row.cells[1], passport_item_text(item.text), font="GOST type A")
            for index in (2, 3, 4):
                set_text(row.cells[index], "")
            protect_row(row, compact=True, allow_split=may_split, height_mm=row_min_mm)
        # Explicit identical cell borders override conditional styles in Word.
        # Adjacent identical edges collapse into one ordinary single line.
        for row in table.rows:
            for cell in row.cells:
                props = cell._tc.get_or_add_tcPr()
                old = props.find(qn('w:tcBorders'))
                if old is not None:
                    props.remove(old)
                borders = OxmlElement('w:tcBorders')
                for edge in ('top', 'left', 'bottom', 'right'):
                    border = OxmlElement('w:' + edge)
                    for attr, value in (('val', 'single'), ('sz', '6'), ('space', '0'), ('color', '000000')):
                        border.set(qn('w:' + attr), value)
                    borders.append(border)
                props.append(borders)
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1
        # Free mode links only the last requirement to the signatures.
        # Whole mode additionally chains cards estimated to fit one column.
        linked_rows = table.rows if keep_table_together else table.rows[-1:]
        for linked_row in linked_rows:
            for cell in linked_row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True
        # A borderless final row is part of this same table. This also prevents
        # renderers that ignore keepNext across a table boundary from orphaning it.
        signature_row = table.add_row()
        signature_cell = signature_row.cells[0].merge(signature_row.cells[4])
        signature_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        cell_props = signature_cell._tc.get_or_add_tcPr()
        cell_borders = OxmlElement("w:tcBorders")
        for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
            element = OxmlElement("w:" + edge)
            element.set(qn("w:val"), "nil")
            cell_borders.append(element)
        cell_props.append(cell_borders)
        cell_margins = OxmlElement("w:tcMar")
        for edge in ("left", "right"):
            element = OxmlElement("w:" + edge)
            element.set(qn("w:w"), "0")
            element.set(qn("w:type"), "dxa")
            cell_margins.append(element)
        cell_props.append(cell_margins)
        signature_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for index, text in enumerate(("Исполнитель (мастер) ______________",
                                      "Контролёр ОТК __________________________",
                                      "Дата проверки __________________________")):
            p = signature_cell.paragraphs[0] if index == 0 else signature_cell.add_paragraph()
            p.add_run(text)
            p.paragraph_format.space_before = Pt(2) if index == 0 else Pt(0)
            p.paragraph_format.keep_with_next = index < 2
            p.paragraph_format.line_spacing = 1.15
            for run in p.runs:
                run.font.size = Pt(10)


def _normalize_character_width(doc) -> None:
    # Standard character width (100%) and normal tracking (0 pt) everywhere.
    # Include unstyled footer runs and inherited style settings, not just titles.
    for element in list(doc._element.xpath('.//w:rPr')) + list(doc.styles.element.xpath('.//w:rPr')):
        for tag, value in (('w:w', '100'), ('w:spacing', '0')):
            child = element.find(qn(tag))
            if child is None:
                child = OxmlElement(tag)
                element.append(child)
            child.set(qn('w:val'), value)


def export_passport_docx(path: str | Path, cards: list[ControlCard], settings: dict | None = None, *,
                         layout_backend: str = "auto", pdf_backend: str = "auto") -> dict:
    """Сохранить паспорт; вернуть итог (layout_checked, moved_tables, note, pdf_path, pdf_error, ...).

    Ошибка PDF не отменяет DOCX: она возвращается в pdf_error для предупреждения.
    """
    _require_docx()
    groups = _selected_groups(cards)
    settings = validate_project_settings(settings or default_project_settings())
    if layout_backend not in LAYOUT_BACKENDS:
        raise ValueError("Неизвестный движок проверки разметки.")
    if pdf_backend not in PDF_BACKENDS:
        raise ValueError("Неизвестный способ сохранения PDF.")
    layout_mode = settings['word_layout_mode']
    doc = _new_document(settings)
    _build_cover(doc, settings)
    _build_control_tables(doc, groups, layout_mode, settings['table_layout'])
    _normalize_character_width(doc)
    doc.save(path)
    log.info("Паспорт сохранён: %s (КК: %s, пунктов: %s, режим: %s)", path, len(groups),
             sum(len(items) for _, _, items in groups), layout_mode)
    # No Word/COM imports or processes at all when the checkbox is off.
    # Explicit API backends override the project checkbox; the GUI uses auto.
    check_requested = (layout_backend == "word" or
                       (layout_backend == "auto" and settings['word_layout_check']))
    if check_requested:
        layout_result = finalize_passport_word_layout(path, backend=layout_backend)
    else:
        layout_result = {
            "layout_checked": False, "moved_tables": 0,
            "note": "Проверка переносов через Word отключена. "
                    "DOCX создан без запуска Microsoft Word.",
        }
    layout_result['layout_check_requested'] = check_requested
    layout_result['layout_mode'] = layout_mode
    layout_result['note'] = "Режим: " + WORD_LAYOUT_MODES[layout_mode] + "\n" + layout_result['note']
    layout_result['pdf_requested'] = settings['pdf_export']
    layout_result['pdf_path'] = None
    layout_result['pdf_error'] = None
    if settings['pdf_export']:
        try:
            layout_result['pdf_path'] = str(export_pdf(path, backend=pdf_backend))
        except (OSError, RuntimeError) as exc:
            log.warning('PDF не создан для %s', path, exc_info=True)
            layout_result['pdf_error'] = str(exc)
    return layout_result
