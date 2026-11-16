"""Небольшие виджеты окна: подсказки, открытие файлов, миниатюра страниц."""
from __future__ import annotations

import os
import ntpath
import subprocess
import sys
from pathlib import Path

try:
    import tkinter as tk
except ImportError:  # модуль импортируется и без tkinter (тесты, сборка)
    tk = None

from .layout_preview import (COLUMN_GAP_MM, COLUMN_WIDTH_MM, MARGIN_MM, PAGE_HEIGHT_MM,
                             PAGE_WIDTH_MM, PT_PER_MM)

PREVIEW_COLORS = {
    'header': ('#BFDBFE', '#2563EB'),
    'row': ('#F9FAFB', '#9CA3AF'),
    'signatures': ('#FDE68A', '#B45309'),
    'overflow': ('#FCA5A5', '#B91C1C'),
}
LEGEND = (('header', 'шапка КК'), ('row', 'пункты'), ('signatures', 'подписи'),
          ('overflow', 'не помещается'))


class Tooltip:
    """Всплывающая подсказка при наведении мыши."""

    def __init__(self, widget, text, delay=550):
        self.widget, self.text, self.delay = widget, text, delay
        self.window = None
        self.after_id = None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self._hide, add='+')
        widget.bind('<ButtonPress>', self._hide, add='+')

    def _schedule(self, _event=None):
        self._cancel()
        self.after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self.after_id is not None:
            try:
                self.widget.after_cancel(self.after_id)
            except tk.TclError:
                pass
            self.after_id = None

    def _show(self):
        self.after_id = None
        if self.window is not None or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self.window = tk.Toplevel(self.widget)
            self.window.wm_overrideredirect(True)
            self.window.wm_geometry(f'+{x}+{y}')
            tk.Label(self.window, text=self.text, justify='left', wraplength=380,
                     background='#FFFBEA', foreground='#111827', relief='solid',
                     borderwidth=1, font=('Arial', 10), padx=8, pady=5).pack()
        except tk.TclError:
            self.window = None

    def _hide(self, _event=None):
        self._cancel()
        if self.window is not None:
            try:
                self.window.destroy()
            except tk.TclError:
                pass
            self.window = None


def tip(widget, text):
    Tooltip(widget, text)
    return widget


def open_path(path) -> None:
    """Открыть файл или папку программой по умолчанию."""
    path = str(path)
    if os.name == 'nt':
        # Tk's save dialog may return //server/share/...; ShellExecute requires
        # Windows separators, especially for UNC paths.
        os.startfile(ntpath.normpath(path))  # noqa: S606 - открываем файл пользователя
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def show_in_folder(path) -> None:
    path = Path(path)
    if os.name == 'nt' and path.exists():
        subprocess.Popen(['explorer', '/select,', ntpath.normpath(str(path))])
    else:
        open_path(path.parent if path.is_file() or not path.exists() else path)


def _draw_cover(canvas, x0, y0, scale, signatures=3, types=0, settings=None):
    """Схема титульного листа: слева подписи, справа логотип, «Утвердил», название и рисунок."""
    width, height = PAGE_WIDTH_MM * scale, PAGE_HEIGHT_MM * scale
    line = '#9CA3AF'
    mid = x0 + width / 2
    canvas.create_line(mid, y0 + height * 0.04, mid, y0 + height * 0.96, fill='#E5E7EB')
    canvas.create_rectangle(mid + width * 0.03, y0 + height * 0.05,
                            mid + width * 0.15, y0 + height * 0.13, outline=line, fill='#F3F4F6')
    for row in range(3):  # «Утвердил»
        y = y0 + height * (0.07 + 0.045 * row)
        canvas.create_line(mid + width * 0.22, y, mid + width * 0.46, y, fill=line)
    if settings is not None:
        from .project import cover_types_layout
        try:
            top_limit = (10 + 185 - 2 * 25.4 / 72 - cover_types_layout(settings)['picture_height']) * scale
        except (ValueError, KeyError, TypeError):
            top_limit = height * 0.9
        title_y = y0 + max(height * 0.18, top_limit - 20 * scale)
    else:
        title_y = y0 + height * 0.30
    canvas.create_text(mid + width * 0.25, title_y, text='ПАСПОРТ',
                       font=('Arial', max(7, int(9 * scale / 0.78)), 'bold'), fill='#374151')
    canvas.create_line(mid + width * 0.08, title_y + 6 * scale,
                       mid + width * 0.42, title_y + 6 * scale, fill=line, width=2)
    canvas.create_line(mid + width * 0.12, title_y + 12 * scale,
                       mid + width * 0.38, title_y + 12 * scale, fill=line)
    top = 0.49
    for row in range(min(int(types or 0), 6)):  # типы / исполнения
        y = (title_y + (16 + row * 5) * scale if settings is not None
             else y0 + height * (top + 0.035 * row))
        canvas.create_rectangle(mid + width * 0.14, y - 2, mid + width * 0.16, y + 2, outline=line)
        canvas.create_line(mid + width * 0.18, y, mid + width * 0.34, y, fill=line)
    if settings is None:
        picture_top = y0 + height * (top + (0.035 * min(int(types or 0), 6) + 0.03 if types else 0.02))
        left, right, bottom = mid + width * 0.05, mid + width * 0.45, y0 + height * 0.94
    else:
        from .project import cover_types_layout
        from .cover_image import PICTURE_WIDTH_MM
        try:
            picture_height = cover_types_layout(settings)['picture_height']
        except (ValueError, KeyError, TypeError):
            picture_height = 10
        # Та же область, что в Word: правый полулист, отступ 5 мм, ширина 124 мм.
        left = x0 + (10 + 138.5 + 5) * scale
        right = left + PICTURE_WIDTH_MM * scale
        bottom = y0 + (10 + 185 - 2 * 25.4 / 72) * scale
        picture_top = bottom - picture_height * scale
    canvas.create_rectangle(left, picture_top, right, bottom, outline=line)
    image = settings.get('cover_image', '') if settings else ''
    if image:
        from PIL import Image, ImageTk
        from .cover_image import decode_image, fit_size, PICTURE_PADDING_X_MM, PICTURE_PADDING_Y_MM
        source = decode_image(image)
        w, h = fit_size(source.width, source.height,
                        right - left - 2 * PICTURE_PADDING_X_MM * scale,
                        bottom - picture_top - 2 * PICTURE_PADDING_Y_MM * scale)
        source = source.resize((max(1, round(w)), max(1, round(h))), Image.Resampling.LANCZOS)
        canvas._cover_photo = ImageTk.PhotoImage(source, master=canvas)
        canvas.create_image((left + right) / 2, (picture_top + bottom) / 2,
                            image=canvas._cover_photo)
    else:
        canvas._cover_photo = None
        canvas.create_text((left + right) / 2, (picture_top + bottom) / 2, text='рисунок',
                           font=('Arial', 7), fill='#9CA3AF')
    count = max(1, min(int(signatures or 1), 10))
    for row in range(count):  # «Разработал» и «Согласовано» — снизу слева
        y = y0 + height * (0.92 - 0.055 * (count - 1 - row))
        canvas.create_line(x0 + width * 0.05, y, x0 + width * 0.22, y, fill=line)
        canvas.create_line(x0 + width * 0.25, y, x0 + width * 0.45, y, fill=line)


def draw_cover_thumbnail(canvas, settings, scale=0.9):
    """Миниатюра титульного листа для шага «Изделие и титульный лист»."""
    canvas.delete('all')
    x0 = y0 = 4
    width, height = PAGE_WIDTH_MM * scale, PAGE_HEIGHT_MM * scale
    canvas.create_rectangle(x0, y0, x0 + width, y0 + height, fill='#FFFFFF', outline='#6B7280')
    reviewers = len(settings.get('reviewers') or [])
    types = len(settings.get('types') or []) if settings.get('types_enabled') else 0
    _draw_cover(canvas, x0, y0, scale, signatures=1 + reviewers, types=types, settings=settings)
    return int(width + 2 * x0), int(height + 2 * y0)


def draw_layout_preview(canvas, preview, available_width, scale=0.78):
    """Нарисовать титульный лист и листы с таблицами. Возвращает высоту рисунка."""
    canvas.delete('all')
    page_w, page_h = PAGE_WIDTH_MM * scale, PAGE_HEIGHT_MM * scale
    gap_x, gap_y, label_h, left = 18, 18, 18, 12
    try:
        available_width = int(available_width)
    except (TypeError, ValueError):
        available_width = 900
    per_row = max(1, int((max(available_width, 300) - left) // (page_w + gap_x)))
    total_pages = 1 + preview.pages
    origins = []
    for page in range(total_pages):
        row, col = divmod(page, per_row)
        x0 = left + col * (page_w + gap_x)
        y0 = 8 + label_h + row * (page_h + gap_y + label_h)
        origins.append((x0, y0))
        title = 'Титульный лист' if page == 0 else f'Лист {page + 1}'
        canvas.create_text(x0, y0 - 9, text=title, anchor='w', font=('Arial', 9, 'bold'),
                           fill='#374151')
        canvas.create_rectangle(x0 + 3, y0 + 3, x0 + page_w + 3, y0 + page_h + 3,
                                fill='#D1D5DB', outline='')
        canvas.create_rectangle(x0, y0, x0 + page_w, y0 + page_h, fill='#FFFFFF',
                                outline='#6B7280')
    _draw_cover(canvas, *origins[0], scale)
    col_w = COLUMN_WIDTH_MM * scale
    pt = scale / PT_PER_MM
    for segment in preview.segments:
        page, side = divmod(segment.column, 2)
        x0, y0 = origins[1 + page]
        x = x0 + MARGIN_MM * scale + side * (COLUMN_WIDTH_MM + COLUMN_GAP_MM) * scale
        top = y0 + MARGIN_MM * scale + segment.top * pt
        bottom = y0 + MARGIN_MM * scale + segment.bottom * pt
        kind = 'overflow' if segment.overflow else segment.kind
        fill, outline = PREVIEW_COLORS[kind]
        canvas.create_rectangle(x, top, x + col_w, max(bottom, top + 1), fill=fill, outline=outline)
        if segment.kind == 'header' and bottom - top >= 8:
            text = f'КК {segment.card}' + (' (продолж.)' if segment.label else '')
            canvas.create_text(x + col_w / 2, (top + bottom) / 2, text=text,
                               font=('Arial', 7, 'bold'), fill='#1E3A8A')
        elif segment.kind == 'row' and bottom - top >= 9 and segment.label:
            canvas.create_text(x + 3, (top + bottom) / 2, text=segment.label, anchor='w',
                               font=('Arial', 6), fill='#374151')
    rows = (total_pages + per_row - 1) // per_row
    return int(8 + rows * (page_h + gap_y + label_h) + 4)
