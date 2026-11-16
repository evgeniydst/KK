"""Окно программы (tkinter): загрузка PDF, выбор пунктов, экспорт в Word."""
from __future__ import annotations

import copy
import datetime
import functools
import logging
import os
import queue
import sys
import threading
import webbrowser
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
except ImportError:  # понятное сообщение — в run_gui()
    tk = None

from .cover_image import from_clipboard as cover_image_from_clipboard, from_file as cover_image_from_file

from .models import (
    Item,
    PAGE_MARKER,
    WELD_WORD,
    commit_item,
    deselect_weld_items,
    duplicate_numbers,
    item_sort_key,
    next_free_number,
    select_number_expression,
)
from .project import (
    WORD_LAYOUT_MODES,
    cover_types_layout,
    default_project_settings,
    passport_filename,
    project_filename,
    project_payload,
    read_project_file,
    validate_project_settings,
    write_project_file,
)
from .morphology import (
    decline_product_title,
)
from .pdf_extract import (
    build_card_items,
    card_from_extraction,
    extract_pdf_details,
    ocr_technical_warning,
    ocr_text_for_editor,
    recognize_ocr_image,
)
from .docx_export import (
    export_passport_docx,
)
from .logs import log_path, setup_logging
from .tempfiles import cleanup_stale
from .paths import ASSETS_DIR
from .table_settings import (BUILTIN_TABLE_LAYOUT, TOTAL_WIDTH_MM, load_table_defaults,
                             save_table_defaults, validate_table_layout)
from .autosave import (
    AUTOSAVE_SECONDS,
    AutosaveSession,
    find_recoverable,
    last_closed_entry,
    mark_broken,
    remove_entry,
    set_aside,
)

from .layout_preview import preview_layout
from .readiness import blocking, collect_issues, count_by_step
from .recent import add_recent, forget_recent, load_recent
from .ui_widgets import (
    LEGEND,
    PREVIEW_COLORS,
    draw_cover_thumbnail,
    draw_layout_preview,
    open_path,
    show_in_folder,
    tip,
)

log = logging.getLogger(__name__)


def clipboard_shortcut(event, platform=None):
    platform = sys.platform if platform is None else platform
    # Windows virtual keycodes identify physical keys independently of RU/EN layout.
    physical = {65: 'SelectAll', 67: 'Copy', 86: 'Paste', 88: 'Cut'}
    if platform == 'win32' and event.keycode in physical:
        return physical[event.keycode]
    return {'a': 'SelectAll', 'c': 'Copy', 'v': 'Paste', 'x': 'Cut',
            'cyrillic_ef': 'SelectAll', 'cyrillic_es': 'Copy',
            'cyrillic_em': 'Paste', 'cyrillic_che': 'Cut',
            'ф': 'SelectAll', 'с': 'Copy', 'м': 'Paste', 'ч': 'Cut'}.get(event.keysym.lower())


def apply_clipboard_action(widget, action):
    if action == 'SelectAll':
        if widget.winfo_class() == 'Text':
            widget.tag_add('sel', '1.0', 'end-1c')
            widget.mark_set('insert', 'end-1c')
        else:
            widget.selection_range(0, 'end')
            widget.icursor('end')
    else:
        widget.event_generate('<<' + action + '>>')


def install_clipboard_support(root):
    """Class bindings cover existing and dynamically created input dialogs."""

    def perform(widget, action):
        try:
            apply_clipboard_action(widget, action)
        except tk.TclError:
            pass  # Empty/non-text clipboard or a destroyed widget.

    def shortcut(event):
        action = clipboard_shortcut(event)
        if action:
            perform(event.widget, action)
            return 'break'  # Never execute the native key binding a second time.

    menu = tk.Menu(root, tearoff=False)

    def context_menu(event):
        event.widget.focus_set()
        menu.delete(0, 'end')
        for label, action in [('Вырезать', 'Cut'), ('Копировать', 'Copy'),
                              ('Вставить', 'Paste'), ('Выделить всё', 'SelectAll')]:
            menu.add_command(label=label,
                command=lambda a=action, w=event.widget: perform(w, a))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return 'break'

    for cls in ('Entry', 'TEntry', 'Text'):
        root.bind_class(cls, '<Control-KeyPress>', shortcut)
        for key in 'acvxACVX':
            root.bind_class(cls, '<Control-' + key + '>', shortcut)
        root.bind_class(cls, '<Button-3>', context_menu)


def _scrollable_tab(tabs, title):
    """Вкладка с вертикальной прокруткой; возвращает (вкладка, рамка содержимого)."""
    page = ttk.Frame(tabs)
    tabs.add(page, text=title)
    canvas = tk.Canvas(page, highlightthickness=0)
    scroll = ttk.Scrollbar(page, orient='vertical', command=canvas.yview)
    scroll.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    canvas.configure(yscrollcommand=scroll.set)
    content = ttk.Frame(canvas, padding=12)
    window_id = canvas.create_window((0, 0), window=content, anchor='nw')
    content.bind('<Configure>', lambda event:
                 canvas.configure(scrollregion=canvas.bbox('all')))
    canvas.bind('<Configure>', lambda event:
                canvas.itemconfigure(window_id, width=event.width))
    return page, content


def _person_fields(parent, heading, role, fio):
    """Блок «Должность / ФИО»; возвращает (рамка, поле должности, поле ФИО)."""
    frame = ttk.LabelFrame(parent, text=heading, padding=8)
    frame.pack(fill='x', pady=(0, 8))
    frame.columnconfigure(1, weight=1)
    frame.columnconfigure(3, weight=1)
    ttk.Label(frame, text='Должность:').grid(row=0, column=0, sticky='w')
    role_entry = ttk.Entry(frame)
    role_entry.insert(0, role)
    role_entry.grid(row=0, column=1, sticky='ew', padx=(6, 12))
    ttk.Label(frame, text='ФИО:').grid(row=0, column=2, sticky='w')
    fio_entry = ttk.Entry(frame)
    fio_entry.insert(0, fio)
    fio_entry.grid(row=0, column=3, sticky='ew', padx=(6, 0))
    return frame, role_entry, fio_entry


def open_ocr_dialog(win, item, number, text, ocr_state):
    """Диалог непроверенного варианта OCR для редактора пункта.

    win — окно редактора, number/text — его поля, ocr_state — общий словарь text/status.
    """
    dialog = tk.Toplevel(win)
    dialog.title("Непроверенный вариант OCR")
    dialog.geometry("760x540")
    dialog.minsize(540, 400)
    dialog.transient(win)
    dialog.grab_set()
    controls = ttk.Frame(dialog, padding=10)
    controls.pack(side="bottom", fill="x")
    body = ttk.Frame(dialog, padding=12)
    body.pack(fill="both", expand=True)
    ttk.Label(body, wraplength=710, text=(
        "OCR может перепутать ±, минусы, цифры и допуски. Сверьте с изображением. "
        "Номер пункта, галочка и исходный текст не меняются. "
        "Применение меняет только черновик; затем нажмите «Сохранить» в редакторе.")).pack(fill="x")
    status = tk.StringVar(value=ocr_state['status'] or 'Результат OCR отсутствует.')
    ttk.Label(body, textvariable=status, wraplength=710).pack(fill="x", pady=8)
    comparison = tk.StringVar(value=ocr_technical_warning(
        item.original_text or item.text, ocr_state['text'], item.number))
    ttk.Label(body, textvariable=comparison, wraplength=710,
              foreground="#9c2500").pack(fill="x", pady=(0, 6))
    ttk.Label(body, text="Текст PDF (не заменялся):").pack(anchor="w")
    native = tk.Text(body, wrap="word", height=3, font=("Arial", 10))
    native.insert("1.0", item.original_text or item.text)
    native.configure(state="disabled")
    native.pack(fill="x", pady=(2, 6))
    ttk.Label(body, text="Вариант OCR:").pack(anchor="w")
    candidate = tk.Text(body, wrap="word", height=12, font=("Arial", 11))
    candidate.insert("1.0", ocr_state['text'])
    candidate.configure(state="disabled")
    candidate.pack(fill="both", expand=True)
    def show_images():
        for path in item.ocr_crops:
            if Path(path).is_file():
                webbrowser.open(Path(path).resolve().as_uri())
    def use_candidate():
        value = ocr_text_for_editor(ocr_state['text'], number.get().strip())
        if not value:
            return
        if messagebox.askyesno("Применить вариант OCR?",
                "Сверили номер, размеры, знаки и допуски с вырезкой?\n"
                "Заменить редактируемый текст этим вариантом? "
                "Автоматической проверки технических значений нет.", parent=dialog):
            text.delete("1.0", "end")
            text.insert("1.0", value)
            dialog.destroy()
            win.grab_set()
    result_queue = queue.Queue()
    def retry():
        retry_button.state(["disabled"])
        apply_button.state(["disabled"])
        status.set("Локальное OCR… документы никуда не отправляются.")
        def worker():
            try:
                values = [recognize_ocr_image(path) for path in item.ocr_crops]
                result_queue.put((True, '\n'.join(values)))
            except Exception as exc:  # ошибка показывается в диалоге, прежний вариант сохраняется
                log.warning('Повтор OCR не удался', exc_info=True)
                result_queue.put((False, str(exc)))
        threading.Thread(target=worker, daemon=True).start()
        dialog.after(100, poll_ocr)
    def poll_ocr():
        if not dialog.winfo_exists():
            return
        try:
            ok, value = result_queue.get_nowait()
        except queue.Empty:
            dialog.after(100, poll_ocr)
            return
        if ok:
            ocr_state.update(text=value, status="Непроверенный вариант OCR. Сверьте с изображением.")
        else:
            # Preserve a previous usable candidate on a failed retry.
            ocr_state['status'] = "Повтор OCR не удался: " + value
        status.set(ocr_state['status'])
        comparison.set(ocr_technical_warning(item.original_text or item.text,
                                           ocr_state['text'], item.number))
        candidate.configure(state="normal")
        candidate.delete("1.0", "end")
        candidate.insert("1.0", ocr_state['text'])
        candidate.configure(state="disabled")
        retry_button.state(["!disabled"])
        apply_button.state(["!disabled"] if ocr_state['text'] else ["disabled"])
    ttk.Button(controls, text="Исходная область OCR", command=show_images).pack(side="left")
    retry_button = ttk.Button(controls, text="Перечитать OCR", command=retry)
    retry_button.pack(side="left", padx=6)
    retry_button.state(["!disabled"] if item.ocr_crops else ["disabled"])
    apply_button = ttk.Button(controls, text="Применить в черновик", command=use_candidate)
    apply_button.pack(side="right")
    apply_button.state(["!disabled"] if ocr_state['text'] else ["disabled"])
    def close():
        dialog.destroy()
        win.grab_set()
    dialog.protocol("WM_DELETE_WINDOW", close)


# --- каркас интерфейса: 4 шага для новичка --------------------------------------

STEPS = (
    ('title', 'Изделие и титульный лист'),
    ('cards', 'Карты контроля'),
    ('check', 'Проверка'),
    ('export', 'Формирование Word'),
)
HEADER_BG = '#1F2937'
SIDEBAR_BG = '#F3F4F6'
SIDEBAR_ACTIVE = '#DBEAFE'
GREEN, GREEN_ACTIVE = '#C8E6C9', '#AED6B0'
BLUE, BLUE_ACTIVE = '#DBEAFE', '#BFDBFE'
STEP_COLORS = {'ok': '#166534', 'warn': '#92400E', 'idle': '#1F2937'}
LEVEL_LABELS = {'error': '✖ Ошибка', 'warning': '⚠ Проверьте', 'note': 'ⓘ Заметка'}
WORD_MODE_HINTS = {
    'free': 'Таблицы плотно заполняют колонки без пустых мест. Длинный пункт может продолжиться '
            'в соседней колонке; первый пункт остаётся с шапкой, последний — с подписями.',
    'whole': 'Короткая КК не разрывается между колонками; внизу колонки может оставаться пустое место.',
}
# Windows: код физической клавиши не зависит от раскладки (Ctrl+S = Ctrl+Ы).
SHORTCUT_KEYCODES = {83: 's', 79: 'o', 78: 'n', 90: 'z', 89: 'y', 49: '1', 50: '2', 51: '3', 52: '4'}
SHORTCUT_KEYSYMS = {
    's': 's', 'o': 'o', 'n': 'n', 'z': 'z', 'y': 'y', '1': '1', '2': '2', '3': '3', '4': '4',
    'cyrillic_yeru': 's', 'ы': 's', 'cyrillic_shcha': 'o', 'щ': 'o', 'cyrillic_te': 'n', 'т': 'n',
    'cyrillic_ya': 'z', 'я': 'z', 'cyrillic_en': 'y', 'н': 'y',
}
TITLE_KEYS = ('name', 'passport_name', 'approval_role', 'approval_name', 'developer_role', 'developer_name',
              'reviewers', 'types_enabled', 'types_title', 'types',
              'cover_dates_enabled', 'approval_date', 'developer_date')

HELP_GENERAL = (
    'Работа идёт по четырём шагам слева:\n'
    '  1. Изделие и титульный лист — название изделия, подписи, исполнения.\n'
    '  2. Карты контроля — добавьте PDF и отметьте галочками пункты для паспорта.\n'
    '  3. Проверка — список замечаний и схема того, как лягут таблицы.\n'
    '  4. Формирование Word — готовый паспорт .docx (и PDF по галочке).\n'
    'Красными и жёлтыми пометками в списке шагов программа подсказывает, где ещё есть что сделать.\n\n'
    'Проект (файл .kkproject) хранит всё: карты, галочки, правки и титульный лист. Сохраняйте его '
    'кнопкой «Сохранить проект» сверху справа. Пока есть несохранённые изменения, программа сама '
    'делает автокопию и предложит восстановить её после сбоя или случайного закрытия.\n\n'
    'Клавиши (работают и в русской раскладке):\n'
    '  Ctrl+N — новый проект,  Ctrl+O — открыть,  Ctrl+S — сохранить,  Ctrl+Shift+S — сохранить как\n'
    '  Ctrl+Z — отменить,  Ctrl+Y — повторить\n'
    '  Ctrl+1 … Ctrl+4 — перейти к шагу,  F1 — эта справка\n'
    '  В списке пунктов: пробел или щелчок по квадратику — галочка; Shift+щелчок — сразу диапазон; '
    'двойной щелчок — изменить текст; правая кнопка мыши — меню действий.'
)

HELP = {
    'start': ('Начальный экран',
              '«Новый проект» — если делаете паспорт на новое изделие. «Открыть проект» — если файл '
              '.kkproject уже есть. Ниже — недавние проекты: двойной щелчок открывает проект. Если '
              'программа закрылась без сохранения, здесь же можно вернуть последнюю автокопию.'),
    'title': ('Шаг 1. Изделие и титульный лист',
              'Введите наименование и обозначение изделия, например «Рама ТМ10.1133.000» — это и есть '
              'название проекта. Строку «паспорт на …» программа склоняет сама (кнопка «Склонить»), '
              'проверьте её. Далее — должности и ФИО для «Утвердил», «Разработал» и «Согласовано». '
              'Если у изделия есть исполнения, включите таблицу исполнений. Справа — миниатюра '
              'титульного листа: она меняется сразу. Поля со звёздочкой обязательны.'),
    'cards': ('Шаг 2. Карты контроля',
              'Нажмите «Добавить PDF карт» — можно выбрать сразу несколько файлов. Слева — список '
              'карт (их порядок = порядок в паспорте, меняется стрелками). Справа — пункты выбранной '
              'карты: отметьте галочками те, что должны попасть в паспорт. Жёлтые строки — символы '
              'распознаны неуверенно, сверьте их с PDF (двойной щелчок по строке). «Выбрать по номерам» '
              'отмечает пункты вида 1-10, 15. Обозначение и наименование детали над списком идут '
              'в шапку таблицы КК.'),
    'check': ('Шаг 3. Проверка',
              'Здесь собраны все замечания по проекту. ✖ Ошибки нужно исправить — без этого '
              'паспорт не сформируется; ⚠ желательно проверить. Двойной щелчок по замечанию '
              'открывает нужную карту и строку. Схема ниже показывает листы паспорта: титульный '
              'лист и две колонки таблиц на каждом листе. Это оценка — точные переносы строк '
              'считает Word.'),
    'export': ('Шаг 4. Формирование Word',
               'Выберите режим переноса таблиц и нажмите зелёную кнопку «Сформировать паспорт '
               'в Word». Если нужен и PDF, поставьте галочку под кнопкой. Паспорт собирается '
               'в фоне — окно не зависает. После формирования здесь появятся кнопки «Открыть '
               'паспорт» и «Показать в папке». Файл проекта при этом не меняется — не забудьте '
               'сохранить и его.'),
}


def _plural(n, one, few, many):
    """1 пункт, 2 пункта, 5 пунктов."""
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def _scroll_area(parent):
    """Прокручиваемая область: вернуть (внешняя рамка, внутренняя рамка для содержимого)."""
    outer = ttk.Frame(parent)
    canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0)
    scroll = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    inner = ttk.Frame(canvas)
    window = canvas.create_window(0, 0, window=inner, anchor='nw')
    inner.bind('<Configure>', lambda event: canvas.configure(scrollregion=canvas.bbox('all')))
    canvas.bind('<Configure>', lambda event: canvas.itemconfigure(window, width=event.width))

    def wheel(event):
        try:
            if not canvas.winfo_ismapped():
                return None
            num = getattr(event, 'num', None)
            if num == 4:
                step = -1
            elif num == 5:
                step = 1
            else:
                delta = getattr(event, 'delta', 0) or 0
                if not delta:
                    return None
                step = -1 if delta > 0 else 1
            top, bottom = canvas.yview()
            if top > 0 or bottom < 1:
                canvas.yview_scroll(step * 3, 'units')
        except tk.TclError:
            pass
        return None

    def enter(event):
        for sequence in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
            canvas.bind_all(sequence, wheel)

    def leave(event):
        try:
            widget = outer.winfo_containing(event.x_root, event.y_root)
        except (tk.TclError, KeyError):
            widget = None
        if widget is None or not str(widget).startswith(str(outer)):
            for sequence in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
                canvas.unbind_all(sequence)

    outer.bind('<Enter>', enter)
    outer.bind('<Leave>', leave)
    return outer, inner


class App:
    """Главное окно: список КК, чек-лист пунктов, настройки титула."""

    def __init__(self, root):
        self.root = root
        self.cards = []
        self.current_card = None
        self.items = []
        self.source = ""
        self.details = {"control_card": False, "graphic_rows": {}, "document_code": "", "warnings": []}
        self.document_code = tk.StringVar()
        self.document_name = tk.StringVar()
        self.events = queue.Queue()
        self.export_events = queue.Queue()
        self.busy = False
        self.check_anchor = None
        self.check_anchor_value = True
        self.settings = default_project_settings()
        self.project_path = ''
        self.undo_stack = []
        self.redo_stack = []
        self.redo_base = None
        self.saved_state = project_payload([], self.settings)
        self.autosave = AutosaveSession()
        self.project_label = tk.StringVar(value='Новый проект')
        self.pending_pdfs = []
        self.pdf_total = 0
        self.pdf_done = 0
        self.last_export = None
        self.last_saved_at = None
        self.last_autosave_at = None
        self.current_step = None
        self.title_form = None
        self._title_form_source = None
        self.issues = []
        self.preview = None
        self.recent = []
        root.protocol("WM_DELETE_WINDOW", self.close_app)
        self._set_window_icon(root)
        root.title("Паспорт КК")
        root.geometry("1240x800")
        root.minsize(980, 620)
        self._configure_styles()
        self.word_layout_mode = tk.StringVar(value=self.settings['word_layout_mode'])
        self.word_layout_check = tk.BooleanVar(value=self.settings['word_layout_check'])
        self.pdf_export = tk.BooleanVar(value=self.settings.get('pdf_export', False))
        self.status = tk.StringVar(value='')
        self._build_menu(root)
        self._build_header(root)
        self._build_statusbar(root)
        body = ttk.Frame(root)
        body.pack(side='top', fill='both', expand=True)
        self._build_sidebar(body)
        self.content = ttk.Frame(body)
        self.content.pack(side='left', fill='both', expand=True)
        self.pages = {}
        self._build_start_page()
        self._build_title_page()
        self._build_cards_page()
        self._build_check_page()
        self._build_export_page()
        self._bind_shortcuts(root)
        self.render_title_form()
        self.refresh()
        self.show_step('start')
        root.after(100, self.poll)
        root.after(AUTOSAVE_SECONDS * 1000, self.autosave_tick)
        root.after(400, self.offer_recovery)

    # --- оформление и каркас окна ---------------------------------------

    def _configure_styles(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=32, font=("Arial", 11))
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
        style.configure('Issues.Treeview', rowheight=26, font=('Arial', 10))
        style.configure('Big.TLabel', font=('Arial', 22, 'bold'), foreground='#111827')
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground='#111827')
        style.configure('Step.TLabel', font=('Arial', 10), foreground='#6B7280')
        style.configure('Hint.TLabel', font=('Arial', 10), foreground='#4B5563')
        style.configure('Strong.TLabel', font=('Arial', 12, 'bold'), foreground='#111827')
        style.configure('Ok.TLabel', font=('Arial', 11, 'bold'), foreground='#15803D')
        style.configure('Warn.TLabel', font=('Arial', 11, 'bold'), foreground='#B45309')
        style.configure('Error.TLabel', font=('Arial', 10, 'bold'), foreground='#B91C1C')
        style.configure('Section.TLabelframe.Label', font=('Arial', 11, 'bold'), foreground='#1D4ED8')
        return style

    def _build_menu(self, root):
        menu = tk.Menu(root)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label='Новый проект', command=self.start_new, accelerator='Ctrl+N')
        file_menu.add_command(label='Открыть проект…', command=self.start_open, accelerator='Ctrl+O')
        file_menu.add_command(label='Сохранить проект', command=self.save_project, accelerator='Ctrl+S')
        file_menu.add_command(label='Сохранить проект как…', command=self.save_project_as,
                              accelerator='Ctrl+Shift+S')
        file_menu.add_command(label='Переименовать проект…', command=self.rename_project)
        file_menu.add_separator()
        file_menu.add_command(label='Последний несохранённый проект…', command=self.open_last_closed)
        file_menu.add_command(label='Начальный экран', command=lambda: self.show_step('start'))
        file_menu.add_separator()
        file_menu.add_command(label='Выход', command=self.close_app)
        menu.add_cascade(label='Проект', menu=file_menu)
        edit_menu = tk.Menu(menu, tearoff=False)
        edit_menu.add_command(label='Отменить', command=self.undo_action, accelerator='Ctrl+Z')
        edit_menu.add_command(label='Повторить', command=self.redo_action, accelerator='Ctrl+Y')
        menu.add_cascade(label='Правка', menu=edit_menu)
        steps_menu = tk.Menu(menu, tearoff=False)
        for number, (key, title) in enumerate(STEPS, 1):
            steps_menu.add_command(label=f'{number}. {title}', accelerator=f'Ctrl+{number}',
                                   command=functools.partial(self.show_step, key))
        menu.add_cascade(label='Шаги', menu=steps_menu)
        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label='Как пользоваться', command=self.show_help, accelerator='F1')
        help_menu.add_command(label='Показать журнал программы', command=self.open_log_folder)
        menu.add_cascade(label='Справка', menu=help_menu)
        root.configure(menu=menu)

    def _build_header(self, root):
        header = tk.Frame(root, background=HEADER_BG, padx=14, pady=8)
        header.pack(side='top', fill='x')
        tk.Label(header, text='Паспорт КК', background=HEADER_BG, foreground='#FFFFFF',
                 font=('Arial', 13, 'bold')).pack(side='left')
        tk.Label(header, text='│', background=HEADER_BG, foreground='#6B7280',
                 font=('Arial', 13)).pack(side='left', padx=10)
        project = tk.Label(header, textvariable=self.project_label, background=HEADER_BG,
                           foreground='#E5E7EB', font=('Arial', 12), cursor='hand2')
        project.pack(side='left')
        project.bind('<Button-1>', lambda event: self.show_step('title'))
        tip(project, 'Название проекта = наименование изделия. Щёлкните, чтобы изменить.')
        self.save_state = tk.StringVar(value='')
        self.save_state_label = tk.Label(header, textvariable=self.save_state, background=HEADER_BG,
                                         foreground='#A7F3D0', font=('Arial', 10))
        self.save_state_label.pack(side='left', padx=(16, 0))
        options = dict(background='#374151', foreground='#FFFFFF', activebackground='#4B5563',
                       activeforeground='#FFFFFF', disabledforeground='#6B7280', relief='flat',
                       borderwidth=0, highlightthickness=0, font=('Arial', 10, 'bold'),
                       padx=12, pady=5, cursor='hand2')
        self.save_button = tk.Button(header, text='Сохранить проект', command=self.save_project, **options)
        self.save_button.pack(side='right')
        self.redo_button = tk.Button(header, text='↷ Повторить', command=self.redo_action,
                                     state='disabled', **options)
        self.redo_button.pack(side='right', padx=(0, 6))
        self.undo_button = tk.Button(header, text='↶ Отменить', command=self.undo_action,
                                     state='disabled', **options)
        self.undo_button.pack(side='right', padx=(0, 6))
        tip(self.save_button, 'Сохранить проект в файл .kkproject (Ctrl+S): галочки, правки и титульный лист.')
        tip(self.undo_button, 'Отменить последнее действие (Ctrl+Z)')
        tip(self.redo_button, 'Вернуть отменённое действие (Ctrl+Y)')

    def _build_statusbar(self, root):
        bar = tk.Frame(root, background='#E5E7EB', padx=12, pady=4)
        bar.pack(side='bottom', fill='x')
        self.progress = ttk.Progressbar(bar, mode='indeterminate', length=180)
        self.status_label = tk.Label(bar, textvariable=self.status, background='#E5E7EB',
                                     foreground='#374151', font=('Arial', 10), anchor='w')
        self.status_label.pack(side='left', fill='x', expand=True)

    def _build_sidebar(self, body):
        side = tk.Frame(body, background=SIDEBAR_BG, width=260)
        side.pack(side='left', fill='y')
        side.pack_propagate(False)
        tk.Label(side, text='ШАГИ', background=SIDEBAR_BG, foreground='#6B7280',
                 font=('Arial', 9, 'bold'), anchor='w').pack(fill='x', padx=18, pady=(18, 6))
        self.step_buttons = {}
        for number, (key, title) in enumerate(STEPS, 1):
            button = tk.Button(side, text=f'{number}. {title}', command=functools.partial(self.show_step, key),
                               anchor='w', justify='left', font=('Arial', 11), relief='flat',
                               borderwidth=0, highlightthickness=0, padx=14, pady=8,
                               background=SIDEBAR_BG, activebackground=SIDEBAR_ACTIVE,
                               foreground=STEP_COLORS['idle'], cursor='hand2')
            button.pack(fill='x', padx=8, pady=2)
            self.step_buttons[key] = (button, number, title)
        tk.Frame(side, background=SIDEBAR_BG).pack(fill='both', expand=True)
        small = dict(anchor='w', font=('Arial', 10), relief='flat', borderwidth=0, highlightthickness=0,
                     padx=14, pady=5, background=SIDEBAR_BG, activebackground='#E5E7EB',
                     foreground='#374151', cursor='hand2')
        tk.Button(side, text='⌂  Начальный экран', command=lambda: self.show_step('start'),
                  **small).pack(fill='x', padx=8)
        tk.Button(side, text='?  Как пользоваться (F1)', command=self.show_help,
                  **small).pack(fill='x', padx=8, pady=(0, 14))

    def _page_frame(self, key, number, title, hint):
        page = ttk.Frame(self.content)
        self.pages[key] = page
        head = ttk.Frame(page, padding=(24, 16, 24, 8))
        head.pack(side='top', fill='x')
        ttk.Label(head, text=f'Шаг {number} из {len(STEPS)}', style='Step.TLabel').pack(anchor='w')
        row = ttk.Frame(head)
        row.pack(fill='x')
        ttk.Label(row, text=title, style='Title.TLabel').pack(side='left')
        ttk.Button(row, text='? Подсказка', command=functools.partial(self.show_help, key)).pack(side='right')
        ttk.Label(head, text=hint, style='Hint.TLabel', wraplength=880, justify='left').pack(anchor='w', pady=(4, 0))
        return page

    def _page_nav(self, page, back=None, forward=None):
        """Кнопки «Назад» / «Далее» внизу шага; создаются до тела страницы, чтобы не уехать за край."""
        nav = ttk.Frame(page, padding=(24, 8, 24, 12))
        nav.pack(side='bottom', fill='x')
        ttk.Separator(page, orient='horizontal').pack(side='bottom', fill='x')
        if back:
            ttk.Button(nav, text=back[1], command=functools.partial(self.show_step, back[0])).pack(side='left')
        if forward:
            tk.Button(nav, text=forward[1], command=functools.partial(self.show_step, forward[0]),
                      background=BLUE, activebackground=BLUE_ACTIVE, foreground='#111827',
                      font=('Arial', 11, 'bold'), relief='solid', borderwidth=1, padx=16, pady=6,
                      cursor='hand2').pack(side='right')
        return nav

    # --- начальный экран ------------------------------------------------

    def _build_start_page(self):
        page = ttk.Frame(self.content, padding=(48, 36, 48, 24))
        self.pages['start'] = page
        ttk.Label(page, text='Паспорт изделия по картам контроля', style='Big.TLabel').pack(anchor='w')
        ttk.Label(page, text='Программа соберёт паспорт в Word из PDF карт контроля (КК). '
                  'Работа идёт по четырём шагам — они слева:\n'
                  '1) изделие и титульный лист  →  2) карты контроля и выбор пунктов  →  '
                  '3) проверка  →  4) формирование Word.',
                  style='Hint.TLabel', justify='left').pack(anchor='w', pady=(8, 24))
        actions = ttk.Frame(page)
        actions.pack(anchor='w')
        big = dict(foreground='#111827', font=('Arial', 13, 'bold'), relief='solid', borderwidth=1,
                   padx=22, pady=14, cursor='hand2')
        new_button = tk.Button(actions, text='+  Новый паспорт', command=self.start_new,
                               background=GREEN, activebackground=GREEN_ACTIVE, **big)
        new_button.pack(side='left')
        open_button = tk.Button(actions, text='Открыть проект…', command=self.start_open,
                                background=BLUE, activebackground=BLUE_ACTIVE, **big)
        open_button.pack(side='left', padx=12)
        self.continue_button = tk.Button(actions, text='Продолжить последний', command=self.start_continue,
                                         background='#F9FAFB', activebackground='#E5E7EB', **big)
        self.continue_button.pack(side='left')
        tip(new_button, 'Пустой проект: начните с названия изделия (Ctrl+N).')
        tip(open_button, 'Открыть сохранённый файл проекта .kkproject (Ctrl+O).')
        tip(self.continue_button, 'Открыть проект, с которым вы работали последним.')
        self.recovery_button = ttk.Button(page, text='Открыть последний несохранённый проект…',
                                          command=self.open_last_closed)
        box = ttk.LabelFrame(page, text='Недавние проекты', padding=12, style='Section.TLabelframe')
        box.pack(fill='both', expand=True, pady=(28, 0))
        self.recent_box = box
        holder = ttk.Frame(box)
        holder.pack(fill='both', expand=True)
        self.recent_list = tk.Listbox(holder, font=('Arial', 11), height=8, activestyle='none',
                                      exportselection=False, borderwidth=1, relief='solid',
                                      highlightthickness=0, selectbackground='#2563EB',
                                      selectforeground='#FFFFFF')
        self.recent_list.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(holder, orient='vertical', command=self.recent_list.yview)
        scroll.pack(side='right', fill='y')
        self.recent_list.configure(yscrollcommand=scroll.set)
        self.recent_list.bind('<Double-1>', lambda event: self.open_recent_selected())
        self.recent_list.bind('<Return>', lambda event: self.open_recent_selected())
        buttons = ttk.Frame(box)
        buttons.pack(fill='x', pady=(8, 0))
        ttk.Button(buttons, text='Открыть выбранный', command=self.open_recent_selected).pack(side='left')
        ttk.Button(buttons, text='Убрать из списка', command=self.forget_recent_selected).pack(side='left', padx=8)
        self.recent_hint = tk.StringVar(value='')
        ttk.Label(buttons, textvariable=self.recent_hint, style='Hint.TLabel').pack(side='left', padx=8)

    def update_start_page(self):
        self.recent = load_recent()
        self.recent_list.delete(0, 'end')
        for path in self.recent:
            item = Path(path)
            self.recent_list.insert('end', f'{item.stem}     —     {item.parent}')
        self.recent_hint.set('' if self.recent else
                             'Здесь появятся проекты, которые вы сохраните или откроете.')
        if self.recent:
            name = Path(self.recent[0]).stem
            self.continue_button.configure(
                text='Продолжить: ' + (name if len(name) <= 32 else name[:31] + '…'), state='normal')
        else:
            self.continue_button.configure(text='Продолжить последний', state='disabled')
        entry = None
        if self.autosave is not None and self.autosave.enabled:
            try:
                entry = last_closed_entry(self.autosave.folder)
            except (OSError, ValueError, TypeError):
                entry = None
        if entry is not None:
            self.recovery_button.pack(anchor='w', pady=(16, 0), before=self.recent_box)
        else:
            self.recovery_button.pack_forget()

    def start_new(self):
        if self.new_project():
            self.show_step('title')
            self.focus_title_name()

    def start_open(self):
        if self.open_project():
            self.show_step('cards')

    def start_continue(self):
        if self.recent:
            self.open_recent(self.recent[0])

    def open_recent(self, path):
        if not Path(path).is_file():
            messagebox.showwarning('Проект не найден', f'Файл не найден:\n{path}\n\n'
                                   'Он убран из списка недавних проектов.')
            forget_recent(path)
            self.update_start_page()
            return False
        if self.open_project(path):
            self.show_step('cards')
            return True
        return False

    def _recent_selection(self):
        selection = self.recent_list.curselection()
        if not selection:
            return None
        index = int(selection[0])
        return self.recent[index] if 0 <= index < len(self.recent) else None

    def open_recent_selected(self):
        path = self._recent_selection()
        if path is None:
            messagebox.showinfo('Недавние проекты', 'Выберите проект в списке.')
            return
        self.open_recent(path)

    def forget_recent_selected(self):
        path = self._recent_selection()
        if path:
            forget_recent(path)
            self.update_start_page()

    # --- шаг 1: изделие и титульный лист ----------------------------------

    def _build_title_page(self):
        page = self._page_frame('title', 1, 'Изделие и титульный лист',
                                'Эти данные попадут на титульный лист паспорта. Изменения сразу попадают '
                                'в проект — отдельной кнопки «Сохранить» здесь нет. Поля со звёздочкой * обязательны.')
        self._page_nav(page, None, ('cards', 'Далее: карты контроля  →'))
        body = ttk.Frame(page, padding=(24, 0, 24, 0))
        body.pack(fill='both', expand=True)
        aside = ttk.Frame(body, padding=(18, 0, 0, 0))
        aside.pack(side='right', fill='y')
        ttk.Label(aside, text='Титульный лист — схема', style='Strong.TLabel').pack(anchor='w')
        self.cover_canvas = tk.Canvas(aside, width=276, height=198, background='#F3F4F6', highlightthickness=0)
        self.cover_canvas.pack(anchor='w', pady=(6, 8))
        self.cover_summary = tk.StringVar(value='')
        ttk.Label(aside, textvariable=self.cover_summary, wraplength=280, justify='left').pack(anchor='w')
        ttk.Label(aside, text='Схема приблизительная; размер и положение рисунка соответствуют рамке в Word. Заводской номер и подписи заполняются в готовом Word.',
                  style='Hint.TLabel', wraplength=280, justify='left').pack(anchor='w', pady=(10, 0))
        self.title_error = tk.StringVar(value='')
        outer, self.title_host = _scroll_area(body)
        outer.pack(side='left', fill='both', expand=True)

    def _form_entry(self, parent, value=''):
        entry = ttk.Entry(parent, font=('Arial', 11))
        entry.insert(0, value or '')
        return entry

    def _bind_form_entry(self, entry):
        entry.bind('<FocusOut>', lambda event: self.apply_title_form(), add='+')
        entry.bind('<Return>', lambda event: self.apply_title_form(), add='+')
        entry.bind('<KeyRelease>', lambda event: self.update_types_fit(), add='+')

    @staticmethod
    def _title_snapshot(settings):
        return copy.deepcopy({key: settings.get(key) for key in TITLE_KEYS})

    def render_title_form(self):
        """Перестроить поля шага 1 по текущим настройкам (новый проект, отмена, открытие)."""
        host = self.title_host
        for child in host.winfo_children():
            child.destroy()
        settings = self.settings
        form = {'reviewers': [], 'types': []}
        self.title_form = form
        self._title_form_source = self._title_snapshot(settings)
        self.title_error.set('')
        ttk.Label(host, textvariable=self.title_error, style='Error.TLabel', wraplength=720).pack(anchor='w')
        product = ttk.LabelFrame(host, text='Изделие', padding=12, style='Section.TLabelframe')
        product.pack(fill='x', pady=(0, 10))
        ttk.Label(product, text='Наименование и обозначение изделия *').pack(anchor='w')
        form['name'] = self._form_entry(product, settings['name'])
        form['name'].pack(fill='x', pady=(3, 2))
        ttk.Label(product, text='Например: Рама ТМ10.1133.000. Это же название проекта и файла паспорта.',
                  style='Hint.TLabel').pack(anchor='w', pady=(0, 10))
        ttk.Label(product, text='Как писать после слова «ПАСПОРТ» (необязательно)').pack(anchor='w')
        row = ttk.Frame(product)
        row.pack(fill='x', pady=(3, 2))
        form['passport_name'] = self._form_entry(row, settings['passport_name'])
        form['passport_name'].pack(side='left', fill='x', expand=True)
        decline = ttk.Button(row, text='Склонить автоматически', command=self.auto_decline_title)
        decline.pack(side='left', padx=(8, 0))
        tip(decline, 'Программа поставит название в родительный падеж: «Рама» → «рамы». Результат можно поправить.')
        ttk.Label(product, text='Пусто — программа склонит сама при формировании. Например: «рамы D9.11.000».',
                  style='Hint.TLabel').pack(anchor='w')
        picture = ttk.LabelFrame(host, text='Изображение изделия', padding=12, style='Section.TLabelframe')
        picture.pack(fill='x', pady=(0, 10))
        buttons = ttk.Frame(picture)
        buttons.pack(anchor='w')
        ttk.Button(buttons, text='Вставить из буфера', command=self.paste_cover_image).pack(side='left')
        ttk.Button(buttons, text='Загрузить с диска…', command=self.load_cover_image).pack(side='left', padx=6)
        ttk.Button(buttons, text='Убрать изображение', command=self.remove_cover_image).pack(side='left')
        self.cover_image_status = tk.StringVar()
        ttk.Label(picture, textvariable=self.cover_image_status, style='Hint.TLabel',
                  wraplength=680).pack(anchor='w', pady=(7, 0))
        signs = ttk.LabelFrame(host, text='Подписи', padding=12, style='Section.TLabelframe')
        signs.pack(fill='x', pady=(0, 10))
        form['cover_dates_enabled'] = tk.BooleanVar(value=settings['cover_dates_enabled'])
        ttk.Checkbutton(signs, text='Проставить даты на титульном листе',
                        variable=form['cover_dates_enabled'], command=self._cover_dates_changed).pack(anchor='w', pady=(0, 6))
        ttk.Label(signs, text='При включении в пустые поля ставится сегодняшняя дата. '
                  'Можно изменить каждую дату или оставить её пустой. Формат: ДД.ММ.ГГГГ.',
                  style='Hint.TLabel', wraplength=650).pack(anchor='w', pady=(0, 6))
        approval_frame, form['approval_role'], form['approval_name'] = _person_fields(
            signs, 'Утвердил (должность *)', settings['approval_role'], settings['approval_name'])
        form['approval_date'] = self._add_cover_date_entry(approval_frame, settings['approval_date'])
        developer_frame, form['developer_role'], form['developer_name'] = _person_fields(
            signs, 'Разработал (должность *)', settings['developer_role'], settings['developer_name'])
        form['developer_date'] = self._add_cover_date_entry(developer_frame, settings['developer_date'])
        ttk.Label(signs, text='Пустое ФИО оставит линию для заполнения от руки.',
                  style='Hint.TLabel').pack(anchor='w')
        reviewers = ttk.LabelFrame(host, text='Согласовано', padding=12, style='Section.TLabelframe')
        reviewers.pack(fill='x', pady=(0, 10))
        ttk.Label(reviewers, text='Список можно дополнять и сокращать (до 20 человек).',
                  style='Hint.TLabel').pack(anchor='w', pady=(0, 6))
        form['people_box'] = ttk.Frame(reviewers)
        form['people_box'].pack(fill='x')
        for person in settings['reviewers']:
            self._add_reviewer_row(person['role'], person['name'], person.get('date', ''))
        ttk.Button(reviewers, text='+ Добавить согласующего', command=self.add_reviewer).pack(anchor='w', pady=(4, 0))
        types = ttk.LabelFrame(host, text='Типы / исполнения (необязательно)', padding=12,
                               style='Section.TLabelframe')
        types.pack(fill='x', pady=(0, 10))
        form['types_enabled'] = tk.BooleanVar(value=settings['types_enabled'])
        ttk.Checkbutton(types, text='Добавить на титульный лист блок с типами / исполнениями',
                        variable=form['types_enabled'], command=self.apply_title_form).pack(anchor='w')
        ttk.Label(types, text='Заголовок, например «Типы кабин:» или «Тип рамы:»').pack(anchor='w', pady=(8, 2))
        form['types_title'] = self._form_entry(types, settings['types_title'])
        form['types_title'].pack(fill='x')
        form['type_rows'] = ttk.Frame(types)
        form['type_rows'].pack(fill='x', pady=6)
        for designation in settings['types']:
            self._add_type_row(designation)
        ttk.Button(types, text='+ Добавить исполнение / обозначение', command=self.add_type).pack(anchor='w')
        form['types_fit'] = tk.StringVar(value='')
        ttk.Label(types, textvariable=form['types_fit'], style='Hint.TLabel', wraplength=700,
                  justify='left').pack(anchor='w', pady=(6, 0))
        ttk.Label(types, text='В Word рядом с каждым обозначением будет пустой квадрат для ручной отметки.',
                  style='Hint.TLabel').pack(anchor='w')
        for key in ('name', 'passport_name', 'approval_role', 'approval_name',
                    'developer_role', 'developer_name', 'types_title'):
            self._bind_form_entry(form[key])
        self.update_title_preview()

    def _set_cover_image(self, image):
        if not self.apply_title_form(silent=False):
            return
        if image != self.settings.get('cover_image', ''):
            self.remember_action()
            self.settings['cover_image'] = image
            self.update_project_title()
            self.update_title_preview()
            self.status.set('Изображение титульного листа обновлено.')

    def paste_cover_image(self):
        try:
            self._set_cover_image(cover_image_from_clipboard())
        except (ValueError, OSError) as exc:
            messagebox.showwarning('Изображение из буфера', str(exc), parent=self.root)

    def load_cover_image(self):
        path = filedialog.askopenfilename(parent=self.root, title='Выберите изображение изделия',
            filetypes=[('Изображения', '*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp'),
                       ('Все файлы', '*.*')])
        if path:
            try:
                self._set_cover_image(cover_image_from_file(path))
            except (ValueError, OSError) as exc:
                messagebox.showwarning('Изображение изделия', str(exc), parent=self.root)

    def remove_cover_image(self):
        self._set_cover_image('')

    def _add_cover_date_entry(self, frame, value=''):
        ttk.Label(frame, text='Дата:').grid(row=1, column=0, sticky='w', pady=(6, 0))
        entry = ttk.Entry(frame, width=14)
        entry.insert(0, value)
        entry.grid(row=1, column=1, sticky='w', padx=(6, 0), pady=(6, 0))
        if not self.title_form['cover_dates_enabled'].get():
            entry.configure(state='disabled')
        self._bind_form_entry(entry)
        return entry

    def _cover_dates_changed(self):
        form = self.title_form
        enabled = bool(form['cover_dates_enabled'].get())
        today = datetime.date.today().strftime('%d.%m.%Y')
        entries = [form['approval_date'], form['developer_date']] + [row[3] for row in form['reviewers']]
        for entry in entries:
            entry.configure(state='normal')
            if enabled and not entry.get().strip():
                entry.insert(0, today)
            if not enabled:
                entry.configure(state='disabled')
        self.apply_title_form()

    def _add_reviewer_row(self, role='', fio='', date=''):
        form = self.title_form
        frame, role_entry, fio_entry = _person_fields(form['people_box'], 'Согласовано', role, fio)
        date_entry = self._add_cover_date_entry(frame, date)
        row = (frame, role_entry, fio_entry, date_entry)
        form['reviewers'].append(row)
        ttk.Button(frame, text='Убрать', command=lambda: self.remove_reviewer(row)).grid(row=0, column=4, padx=(8, 0))
        self._bind_form_entry(role_entry)
        self._bind_form_entry(fio_entry)
        return row

    def add_reviewer(self):
        if len(self.title_form['reviewers']) >= 20:
            messagebox.showwarning('Согласующие', 'Максимум 20 согласующих.', parent=self.root)
            return
        date = datetime.date.today().strftime('%d.%m.%Y') if self.title_form['cover_dates_enabled'].get() else ''
        row = self._add_reviewer_row(date=date)
        row[1].focus_set()
        self.update_title_preview()

    def remove_reviewer(self, row):
        if row in self.title_form['reviewers']:
            self.title_form['reviewers'].remove(row)
            row[0].destroy()
            self.apply_title_form()

    def _add_type_row(self, text=''):
        form = self.title_form
        frame = ttk.Frame(form['type_rows'])
        frame.pack(fill='x', pady=2)
        ttk.Label(frame, text='□').pack(side='left', padx=(0, 6))
        entry = self._form_entry(frame, text)
        entry.pack(side='left', fill='x', expand=True)
        pair = (frame, entry)
        form['types'].append(pair)
        ttk.Button(frame, text='Убрать', command=lambda: self.remove_type(pair)).pack(side='right', padx=(6, 0))
        self._bind_form_entry(entry)
        return pair

    def add_type(self):
        pair = self._add_type_row()
        pair[1].focus_set()
        self.update_types_fit()

    def remove_type(self, pair):
        if pair in self.title_form['types']:
            self.title_form['types'].remove(pair)
            pair[0].destroy()
            self.apply_title_form()

    def _form_candidate(self):
        form = self.title_form
        candidate = dict(self.settings)
        for key in ('name', 'passport_name', 'approval_role', 'approval_name',
                    'developer_role', 'developer_name', 'types_title'):
            candidate[key] = form[key].get()
        candidate['types_enabled'] = bool(form['types_enabled'].get())
        candidate['types'] = [entry.get() for _, entry in form['types']]
        candidate['cover_dates_enabled'] = bool(form['cover_dates_enabled'].get())
        candidate['approval_date'] = form['approval_date'].get()
        candidate['developer_date'] = form['developer_date'].get()
        candidate['reviewers'] = [{'role': role.get(), 'name': fio.get(), 'date': date.get()}
                                  for _, role, fio, date in form['reviewers']]
        if (candidate['name'].strip() != self.settings['name']
                and candidate['passport_name'].strip() == self.settings['passport_name']):
            # Новое изделие — старое склонение больше не подходит.
            candidate['passport_name'] = ''
        return candidate

    def apply_title_form(self, silent=True):
        """Перенести поля шага 1 в проект. False — если данные нельзя принять."""
        if self.title_form is None:
            return True
        if self._title_snapshot(self.settings) != self._title_form_source:
            # проект изменился не через форму (переименование, восстановление) — форма устарела
            self.render_title_form()
            return True
        try:
            candidate = self._form_candidate()
        except tk.TclError:  # форма как раз перестраивается
            return True
        try:
            candidate = validate_project_settings(candidate)
        except ValueError as exc:
            self.title_error.set('⚠ ' + str(exc))
            if not silent:
                messagebox.showwarning('Проверьте титульный лист', str(exc), parent=self.root)
            return False
        self.title_error.set('')
        if candidate != self.settings:
            self.remember_action()
            self.settings = candidate
            self._title_form_source = self._title_snapshot(candidate)
            entry = self.title_form['passport_name']
            try:
                if entry.get().strip() != candidate['passport_name']:
                    entry.delete(0, 'end')
                    entry.insert(0, candidate['passport_name'])
            except tk.TclError:
                pass
            self.update_project_title()
        self.update_title_preview()
        return True

    def update_types_fit(self):
        form = self.title_form
        if form is None:
            return
        try:
            candidate = self._form_candidate()
            layout = cover_types_layout(candidate)
            text = ((f"Позиций: {len(form['types'])}. " if candidate['types_enabled'] else 'Блок выключен. ') +
                    f"Высота рисунка на титуле: {layout['picture_height']:.1f} мм. "
                    'Больше строк — меньше рисунок, шрифт не меняется.')
        except tk.TclError:
            return
        except (ValueError, KeyError, TypeError) as exc:
            text = str(exc)
        form['types_fit'].set(text)

    def update_title_preview(self):
        """Миниатюра и краткое содержание титульного листа справа от формы."""
        self.update_types_fit()
        settings = self.settings
        name = settings['name'].strip()
        passport = settings['passport_name'].strip()
        if not passport and name:
            try:
                passport = decline_product_title(name)
            except Exception:  # склонение недоступно — покажем исходное название
                passport = name
        reviewers = settings['reviewers']
        lines = ['ПАСПОРТ ' + (passport or '… (укажите изделие)'), '',
                 'Утвердил: ' + (settings['approval_role'] or '— должность не указана —') +
                 (' ' + settings['approval_name'] if settings['approval_name'] else ''),
                 'Разработал: ' + (settings['developer_role'] or '— должность не указана —') +
                 (' ' + settings['developer_name'] if settings['developer_name'] else ''),
                 'Согласовано: ' + (f'{len(reviewers)} чел.' if reviewers else 'никого')]
        if settings['types_enabled']:
            lines.append((settings['types_title'] or 'Типы') + f" {len(settings['types'])} поз.")
        self.cover_summary.set('\n'.join(lines))
        self.cover_image_status.set('Изображение добавлено. Чтобы заменить его, вставьте новое или выберите файл.'
                                    if settings.get('cover_image') else
                                    'Изображение не выбрано: в Word останется пустая рамка с подсказкой.')
        try:
            draw_cover_thumbnail(self.cover_canvas, settings, scale=0.9)
        except tk.TclError:
            pass

    def auto_decline_title(self):
        form = self.title_form
        try:
            value = decline_product_title(form['name'].get())
        except RuntimeError as exc:
            messagebox.showwarning('Склонение названия', str(exc), parent=self.root)
            return
        form['passport_name'].delete(0, 'end')
        form['passport_name'].insert(0, value)
        self.apply_title_form()

    def focus_title_name(self):
        if self.title_form is not None:
            try:
                self.title_form['name'].focus_set()
            except tk.TclError:
                pass

    def edit_project_settings(self):
        """Раньше — отдельное окно; теперь титульный лист — это шаг 1."""
        self.show_step('title')

    # --- шаг 2: карты контроля ----------------------------------------------

    def _build_cards_page(self):
        page = self._page_frame('cards', 2, 'Карты контроля',
                                'Добавьте PDF карт контроля и отметьте галочками пункты, которые войдут в паспорт. '
                                'Порядок карт в списке слева — это порядок в паспорте.')
        self._page_nav(page, ('title', '←  Изделие'), ('check', 'Далее: проверка  →'))
        panes = ttk.Panedwindow(page, orient='horizontal')
        panes.pack(fill='both', expand=True, padx=24, pady=(4, 8))
        left = ttk.Frame(panes, padding=(0, 0, 10, 0))
        right = ttk.Frame(panes, padding=(10, 0, 0, 0))
        panes.add(left, weight=1)
        panes.add(right, weight=4)
        self.open_button = tk.Button(left, text='+  Добавить PDF карты…', command=self.open_pdf,
                                     background=GREEN, activebackground=GREEN_ACTIVE, foreground='#111827',
                                     font=('Arial', 11, 'bold'), relief='solid', borderwidth=1, pady=8,
                                     cursor='hand2')
        self.open_button.pack(fill='x')
        tip(self.open_button, 'Выберите один или несколько PDF. Программа найдёт обозначение, '
                              'наименование и нумерованные пункты.')
        ttk.Label(left, text='Карты в паспорте (по порядку):', style='Hint.TLabel').pack(anchor='w', pady=(12, 4))
        holder = ttk.Frame(left)
        holder.pack(fill='both', expand=True)
        self.card_list = tk.Listbox(holder, font=('Arial', 11), activestyle='none', exportselection=False,
                                    borderwidth=1, relief='solid', highlightthickness=0,
                                    selectbackground='#2563EB', selectforeground='#FFFFFF', width=34)
        self.card_list.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(holder, orient='vertical', command=self.card_list.yview)
        scroll.pack(side='right', fill='y')
        self.card_list.configure(yscrollcommand=scroll.set)
        self.card_list.bind('<<ListboxSelect>>', self.switch_card)
        self.card_list.bind('<Button-3>', self.card_context_menu)
        order = ttk.Frame(left)
        order.pack(fill='x', pady=(6, 0))
        self.move_up_button = ttk.Button(order, text='↑ Выше', width=8,
                                         command=lambda: self.move_current_card(offset=-1))
        self.move_up_button.pack(side='left')
        self.move_down_button = ttk.Button(order, text='↓ Ниже', width=8,
                                           command=lambda: self.move_current_card(offset=1))
        self.move_down_button.pack(side='left', padx=4)
        self.card_more_button = ttk.Menubutton(order, text='Ещё ▾')
        self.card_more_menu = tk.Menu(self.card_more_button, tearoff=False)
        self.card_more_menu.add_command(label='Поставить на место…', command=self.choose_card_position)
        self.card_more_menu.add_separator()
        self.card_more_menu.add_command(label='Удалить КК из паспорта…', command=self.remove_current_card)
        self.card_more_button.configure(menu=self.card_more_menu)
        self.card_more_button.pack(side='left')
        tip(self.move_up_button, 'Переместить выбранную КК выше в паспорте')
        tip(self.move_down_button, 'Переместить выбранную КК ниже в паспорте')
        self.card_empty = ttk.Frame(right, padding=(24, 32))
        ttk.Label(self.card_empty, text='Пока нет ни одной карты контроля', style='Title.TLabel').pack(anchor='w')
        ttk.Label(self.card_empty, text='Нажмите «+ Добавить PDF карты…» слева. Можно выбрать сразу несколько '
                  'файлов — они добавятся по очереди.\n\n'
                  'Программа сама найдёт обозначение детали, наименование изделия и нумерованные '
                  'пункты. Потом останется отметить галочками пункты для паспорта.',
                  style='Hint.TLabel', wraplength=640, justify='left').pack(anchor='w', pady=(10, 0))
        self.card_editor = ttk.Frame(right)
        identity = ttk.Frame(self.card_editor)
        identity.pack(fill='x')
        identity.columnconfigure(1, weight=1)
        ttk.Label(identity, text='Обозначение детали:').grid(row=0, column=0, sticky='w', pady=2)
        self.code_entry = ttk.Entry(identity, textvariable=self.document_code, font=('Arial', 11))
        self.code_entry.grid(row=0, column=1, sticky='ew', padx=(8, 0), pady=2)
        ttk.Label(identity, text='Наименование изделия:').grid(row=1, column=0, sticky='w', pady=2)
        self.name_entry = ttk.Entry(identity, textvariable=self.document_name, font=('Arial', 11))
        self.name_entry.grid(row=1, column=1, sticky='ew', padx=(8, 0), pady=2)
        for entry in (self.code_entry, self.name_entry):
            entry.bind('<FocusIn>', lambda event: self.remember_action())
            entry.bind('<FocusOut>', lambda event: self.refresh_card_list())
        self.card_summary = tk.StringVar(value='')
        ttk.Label(self.card_editor, textvariable=self.card_summary, style='Hint.TLabel').pack(anchor='w', pady=(6, 4))
        toolbar = ttk.Frame(self.card_editor)
        toolbar.pack(fill='x', pady=(0, 6))
        for text, command, hint in (
                ('☑ Отметить все', lambda: self.set_all(True), 'Отметить все пункты этой КК'),
                ('☐ Снять все', lambda: self.set_all(False), 'Снять все галочки в этой КК'),
                ('+ Пункт', self.add_item, 'Добавить свой пункт до или после выбранной строки'),
                ('Удалить пункт', self.delete_item, 'Удалить выбранный пункт из списка')):
            button = ttk.Button(toolbar, text=text, command=command)
            button.pack(side='left', padx=(0, 6))
            tip(button, hint)
        more = ttk.Menubutton(toolbar, text='Ещё ▾')
        more_menu = tk.Menu(more, tearoff=False)
        more_menu.add_command(label='Отметить по номерам…', command=self.choose_numbers)
        more_menu.add_command(label='Снять галочки со сварных швов', command=self.deselect_welds)
        more_menu.add_separator()
        more_menu.add_command(label='Собрать пункты из исходного текста заново',
                              command=lambda: self.rebuild(False))
        more_menu.add_command(label='Каждая строка текста — отдельный пункт', command=lambda: self.rebuild(True))
        more.configure(menu=more_menu)
        more.pack(side='left')
        self.item_menu = tk.Menu(self.root, tearoff=False)
        self._build_tabs(self.card_editor)

    def _build_tabs(self, parent):
        self.tabs = ttk.Notebook(parent)
        self.tabs.pack(fill="both", expand=True)
        table = ttk.Frame(self.tabs, padding=6)
        rawtab = ttk.Frame(self.tabs, padding=6)
        self.tabs.add(table, text="Чек-лист")
        self.tabs.add(rawtab, text="Исходный текст из PDF")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=("check", "number", "text", "page"),
                                 show="headings", selectmode="browse")
        for name, title, width in (("check", "В паспорт", 90), ("number", "№", 70),
                                  ("text", "Текст пункта", 740), ("page", "Стр.", 70)):
            self.tree.heading(name, text=title)
            self.tree.column(name, width=width, minwidth=50, stretch=name == "text")
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        sy.grid(row=0, column=1, sticky="ns")
        sx = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        sx.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.bind("<Button-1>", self.click)
        self.tree.bind("<space>", self.toggle_focus)
        self.tree.bind("<Shift-space>", self.toggle_focus)
        self.tree.bind("<Double-1>", self.double_click)
        self.tree.bind("<Button-3>", self.item_context_menu)
        ttk.Label(table, text="☑ — пункт войдёт в паспорт. Щелчок по галочке или пробел — отметить, "
                  "Shift + щелчок — диапазон. Двойной щелчок по тексту — открыть и исправить. "
                  "Жёлтые строки — сверьте с PDF.", style='Hint.TLabel',
                  wraplength=900).grid(row=2, column=0, sticky="w", pady=8)
        rawtab.rowconfigure(1, weight=1)
        rawtab.columnconfigure(0, weight=1)
        ttk.Label(rawtab, text="Текст, извлечённый из PDF. Можно удалить лишние заголовки и колонтитулы, "
                  "затем нажать «Собрать пункты заново».", style='Hint.TLabel',
                  wraplength=900).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.raw = tk.Text(rawtab, wrap="word", font=("Arial", 11), undo=True)
        self.raw.grid(row=1, column=0, sticky="nsew")
        self.raw.bind('<FocusIn>', lambda event: self.remember_action())
        self.raw.bind('<FocusOut>', lambda event: self.update_project_title())
        rs = ttk.Scrollbar(rawtab, command=self.raw.yview)
        rs.grid(row=1, column=1, sticky="ns")
        self.raw.configure(yscrollcommand=rs.set)
        tools = ttk.Frame(rawtab)
        tools.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(tools, text="Собрать пункты заново", command=lambda: self.rebuild(False)).pack(side="left")
        ttk.Button(tools, text="Каждая строка — отдельный пункт", command=lambda: self.rebuild(True)).pack(side="left", padx=8)
        ttk.Label(rawtab, text="Строки <<<PDF_PAGE:…>>> обозначают страницы. Пересборка сбросит галочки этой КК.",
                  style='Hint.TLabel', wraplength=900).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def item_context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return None
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        item = self.selected_item()
        if item is None:
            return None
        menu = self.item_menu
        menu.delete(0, 'end')
        menu.add_command(label='Открыть и исправить…', command=lambda: self.edit_item(item))
        menu.add_command(label='Добавить свой пункт…', command=self.add_item)
        menu.add_command(label='Снять галочку' if item.checked else 'Отметить для паспорта',
                         command=lambda: self.toggle(iid))
        menu.add_separator()
        menu.add_command(label='Удалить пункт', command=self.delete_item)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return 'break'

    def card_context_menu(self, event):
        index = self.card_list.nearest(event.y)
        if not 0 <= index < len(self.cards):
            return None
        if self.cards[index] is not self.current_card:
            self.activate_card(self.cards[index])
        try:
            self.card_more_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.card_more_menu.grab_release()
        return 'break'

    def update_card_selector(self):
        self.card_list.delete(0, 'end')
        for number, card in enumerate(self.cards, 1):
            checked = sum(item.checked for item in card.items)
            warn = ' ⚠' if (not card.document_code or not card.document_name or duplicate_numbers(card.items)
                            or any(item.warnings for item in card.items if item.checked)) else ''
            self.card_list.insert('end', f"{number}. {card.document_code or 'Без обозначения'} — "
                                         f"{card.document_name or Path(card.source).name}  ({checked}){warn}")
        self.card_list.selection_clear(0, 'end')
        if self.current_card in self.cards:
            index = self.cards.index(self.current_card)
            self.card_list.selection_set(index)
            self.card_list.see(index)

    def refresh_card_list(self):
        self.store_current_card()
        self.update_card_selector()
        self.update_project_title()

    def switch_card(self, event=None):
        selection = self.card_list.curselection()
        if not selection:
            return
        index = int(selection[0])
        if 0 <= index < len(self.cards) and self.cards[index] is not self.current_card:
            self.activate_card(self.cards[index])

    def refresh(self):
        if not any(item.manual for item in self.items):
            self.items.sort(key=item_sort_key)
        self.store_current_card()
        self.update_card_selector()
        if self.current_card is not None:
            self.card_empty.pack_forget()
            self.card_editor.pack(fill='both', expand=True)
        else:
            self.card_editor.pack_forget()
            self.card_empty.pack(fill='both', expand=True)
        self.tree.delete(*self.tree.get_children())
        self.tree.tag_configure("review", background="#fff0d6")
        for item in self.items:
            self.tree.insert("", "end", iid=item.id, tags=("review",) if item.warnings else (), values=("☑" if item.checked else "☐",
                item.number, item.text.replace("\n", " ↵ "), ", ".join(map(str, item.pages))))
        checked = sum(item.checked for item in self.items)
        review = sum(bool(item.warnings) for item in self.items)
        parts = [f'Пунктов: {len(self.items)}', f'отмечено для паспорта: {checked}']
        if review:
            parts.append(f'⚠ сверить с PDF: {review}')
        if self.source:
            parts.append('файл: ' + Path(self.source).name)
        self.card_summary.set('  ·  '.join(parts))
        if not self.busy:
            total = sum(item.checked for card in self.cards for item in card.items)
            self.status.set(f'КК в проекте: {len(self.cards)}  ·  отмечено пунктов для паспорта: {total}'
                            if self.cards else 'Начните с шага 1 или добавьте PDF карт контроля на шаге 2.')
        self.refresh_controls()

    def refresh_controls(self):
        index = next((n for n, card in enumerate(self.cards) if card is self.current_card), -1)
        for button, disabled in ((self.move_up_button, index <= 0),
                                 (self.move_down_button, index < 0 or index >= len(self.cards) - 1),
                                 (self.card_more_button, index < 0)):
            button.state(['disabled'] if self.busy or disabled else ['!disabled'])
        self.update_project_title()

    # --- импорт PDF (можно несколько файлов сразу) ---------------------------

    @staticmethod
    def _file_key(path):
        try:
            return os.path.normcase(str(Path(path).resolve()))
        except (OSError, RuntimeError):
            return os.path.normcase(os.path.abspath(path))

    def open_pdf(self):
        if self.busy:
            return
        selected = filedialog.askopenfilenames(title='Выберите PDF карт контроля (можно несколько)',
                                               filetypes=[("PDF", "*.pdf")])
        if isinstance(selected, str):
            try:
                selected = self.root.tk.splitlist(selected) if selected else ()
            except Exception:  # редкие сборки Tk возвращают одну строку
                selected = (selected,)
        paths, seen = [], set()
        for path in selected or ():
            key = self._file_key(path) if path else None
            if key and key not in seen:
                seen.add(key)
                paths.append(path)
        if not paths:
            return
        known = {self._file_key(card.source): card for card in self.cards if card.source}
        fresh = [path for path in paths if self._file_key(path) not in known]
        existing = [known[self._file_key(path)] for path in paths if self._file_key(path) in known]
        if not fresh:
            self.activate_card(existing[0])
            messagebox.showinfo("КК уже добавлена", "Этот PDF уже есть в списке. Переключился на его чек-лист."
                                if len(paths) == 1 else "Все выбранные PDF уже есть в списке.")
            return
        if existing:
            messagebox.showinfo('Часть КК уже добавлена', 'Уже есть в списке, пропускаю:\n' +
                                '\n'.join(Path(card.source).name for card in existing))
        if not self.ensure_project_name():
            return
        self.pending_pdfs = list(fresh)
        self.pdf_total, self.pdf_done = len(fresh), 0
        self._start_next_pdf()

    def _start_next_pdf(self):
        if not self.pending_pdfs:
            return False
        path = self.pending_pdfs.pop(0)
        self.pdf_done += 1
        counter = f' {self.pdf_done} из {self.pdf_total}' if self.pdf_total > 1 else ''
        self.set_busy(True, f'Считываю PDF{counter}: {Path(path).name}… '
                            'Распознавание значков со скана может занять время.')
        def worker():
            try:
                self.events.put((True, path, extract_pdf_details(path, ocr=True)))
            except Exception as exc:  # любая ошибка чтения показывается пользователю, окно не падает
                log.exception('Не удалось прочитать PDF %s', path)
                self.events.put((False, path, str(exc)))
        threading.Thread(target=worker, daemon=True).start()
        return True

    def set_busy(self, busy, message='', exporting=False):
        self.busy = busy
        try:
            self.root.configure(cursor='watch' if busy else '')
        except tk.TclError:
            pass
        state = 'disabled' if busy else 'normal'
        for widget in (self.open_button, self.word_button):
            widget.configure(state=state)
        if busy:
            self.status.set(message)
            self.progress.pack(side='right', padx=(12, 0), before=self.status_label)
            self.progress.start(12)
            if exporting:
                self.export_progress_text.set(message)
                self.export_progress.pack(fill='x', pady=(14, 0))
                self.export_progressbar.start(12)
        else:
            self.progress.stop()
            self.progress.pack_forget()
            self.export_progressbar.stop()
            self.export_progress.pack_forget()
        self.refresh_controls()

    def poll(self):
        try:
            self._poll_pdf()
            self._poll_export()
        finally:
            self.root.after(100, self.poll)

    def _poll_pdf(self):
        try:
            ok, path, value = self.events.get_nowait()
        except queue.Empty:
            return
        self.set_busy(False)
        if ok:
            card = card_from_extraction(path, value)
            self.remember_action()
            self.cards.append(card)
            if self.current_step != 'cards':
                self.show_step('cards')
            self.activate_card(card)
            if value.get("identity_warnings"):
                messagebox.showwarning("Проверьте сведения карты",
                                       Path(path).name + ":\n\n" + "\n\n".join(value["identity_warnings"]),
                                       parent=self.root)
            duplicates = duplicate_numbers(card.items)
            if duplicates:
                messagebox.showwarning("Повторные номера в PDF", Path(path).name + ": из PDF извлечены повторные номера " +
                                       ", ".join(duplicates) + ". Данные сохранены; проверьте и исправьте их вручную.")
            raw = value["raw"]
            self.refresh()
            if (not any(line.strip() and not PAGE_MARKER.fullmatch(line.strip()) for line in raw.splitlines())
                    and not any(entry["type"] in ("cover", "control") for entry in self.details.get("page_types", []))):
                messagebox.showwarning("В PDF нет текста", Path(path).name + ": в PDF не найден текстовый слой. "
                                       "Вероятно, это скан целиком. Со скана программа распознаёт только отдельные "
                                       "значки и параметры; полностью сканированную карту нужно сначала "
                                       "распознать отдельно.")
            elif not self.items and all(entry["type"] == "cover" for entry in self.details.get("page_types", [])):
                messagebox.showinfo("Только титульный лист", Path(path).name + ": сведения об изделии извлечены, "
                                    "но нумерованных требований на титуле нет. Откройте полный PDF этой КК "
                                    "с листами контроля.")
            elif not self.items:
                self.tabs.select(1)
                messagebox.showinfo("Нет нумерованных пунктов", Path(path).name + ": текст считан, но строки вида "
                                    "«1. Текст» не найдены. Исправьте текст на вкладке «Исходный текст из PDF» "
                                    "или нажмите «Каждая строка — отдельный пункт».")
        else:
            self.status.set('Не удалось прочитать ' + Path(path).name + '; остальные КК не затронуты.')
            messagebox.showerror("Ошибка чтения PDF", f"{Path(path).name}:\n\n{value}")
        self._start_next_pdf()

    # --- шаг 3: проверка ------------------------------------------------------

    def _build_check_page(self):
        page = self._page_frame('check', 3, 'Проверка перед формированием',
                                'Программа собрала всё, что стоит проверить. Двойной щелчок по замечанию '
                                'откроет нужное место. Ниже — схема листов: как примерно лягут таблицы.')
        self._page_nav(page, ('cards', '←  Карты контроля'), ('export', 'Далее: формирование Word  →'))
        body = ttk.Frame(page, padding=(24, 0, 24, 8))
        body.pack(fill='both', expand=True)
        self.check_summary = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.check_summary, style='Strong.TLabel', wraplength=900,
                  justify='left').pack(anchor='w', pady=(0, 8))
        issues = ttk.LabelFrame(body, text='Замечания', padding=10, style='Section.TLabelframe')
        issues.pack(fill='x')
        self.issues_ok = ttk.Label(issues, text='✓ Замечаний нет — можно формировать паспорт.', style='Ok.TLabel')
        self.issues_holder = ttk.Frame(issues)
        self.issue_tree = ttk.Treeview(self.issues_holder, columns=('level', 'place', 'text'), show='headings',
                                       selectmode='browse', height=6, style='Issues.Treeview')
        for name, title, width in (('level', 'Важность', 120), ('place', 'Где', 140), ('text', 'Что не так', 640)):
            self.issue_tree.heading(name, text=title)
            self.issue_tree.column(name, width=width, minwidth=60, stretch=name == 'text')
        self.issue_tree.tag_configure('error', foreground='#B91C1C')
        self.issue_tree.tag_configure('warning', foreground='#92400E')
        self.issue_tree.tag_configure('note', foreground='#374151')
        self.issue_tree.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(self.issues_holder, orient='vertical', command=self.issue_tree.yview)
        scroll.pack(side='right', fill='y')
        self.issue_tree.configure(yscrollcommand=scroll.set)
        self.issue_tree.bind('<Double-1>', self.go_to_issue)
        self.issue_tree.bind('<Return>', self.go_to_issue)
        self.issue_buttons = ttk.Frame(issues)
        ttk.Button(self.issue_buttons, text='Перейти к месту', command=self.go_to_issue).pack(side='left')
        ttk.Label(self.issue_buttons, text='✖ ошибка — исправить обязательно;  ⚠ проверьте — желательно '
                  'сверить;  ⓘ заметка — просто для сведения.', style='Hint.TLabel').pack(side='left', padx=10)
        preview = ttk.LabelFrame(body, text='Как лягут таблицы — примерная схема', padding=10,
                                 style='Section.TLabelframe')
        preview.pack(fill='both', expand=True, pady=(10, 0))
        legend = ttk.Frame(preview)
        legend.pack(fill='x')
        for key, label in LEGEND:
            fill, outline = PREVIEW_COLORS[key]
            swatch = tk.Canvas(legend, width=16, height=12, highlightthickness=0)
            swatch.create_rectangle(1, 1, 15, 11, fill=fill, outline=outline)
            swatch.pack(side='left', padx=(0, 4))
            ttk.Label(legend, text=label, style='Hint.TLabel').pack(side='left', padx=(0, 14))
        self.preview_info = tk.StringVar(value='')
        ttk.Label(legend, textvariable=self.preview_info, style='Hint.TLabel').pack(side='right')
        holder = ttk.Frame(preview)
        holder.pack(fill='both', expand=True, pady=(6, 0))
        self.preview_canvas = tk.Canvas(holder, background='#E5E7EB', highlightthickness=0, height=240)
        self.preview_canvas.pack(side='left', fill='both', expand=True)
        pscroll = ttk.Scrollbar(holder, orient='vertical', command=self.preview_canvas.yview)
        pscroll.pack(side='right', fill='y')
        self.preview_canvas.configure(yscrollcommand=pscroll.set)
        self.preview_canvas.bind('<Configure>', lambda event: self.draw_preview())
        ttk.Label(preview, text='Схема приблизительная: точные переносы строк считает Word. Перенос таблиц '
                  'меняется на шаге 4.', style='Hint.TLabel').pack(anchor='w', pady=(6, 0))

    def update_check_page(self):
        self.store_current_card()
        issues = list(collect_issues(self.cards, self.settings))
        try:
            self.preview = preview_layout(self.cards, self.settings)
        except Exception:  # схема — подсказка; её сбой не должен мешать формированию
            log.warning('Схема листов не построена', exc_info=True)
            self.preview = None
        if self.preview is not None:
            issues.extend(self.preview.issues)
        order = {'error': 0, 'warning': 1, 'note': 2}
        issues.sort(key=lambda issue: order.get(issue.level, 3))
        self.issues = issues
        self.issue_tree.delete(*self.issue_tree.get_children())
        for index, issue in enumerate(issues):
            place = 'Титульный лист' if issue.step == 'title' else (
                f'КК {issue.card_index + 1}' if issue.card_index is not None else 'Паспорт')
            self.issue_tree.insert('', 'end', iid=str(index), tags=(issue.level,),
                                   values=(LEVEL_LABELS.get(issue.level, issue.level), place, issue.text))
        if issues:
            self.issues_ok.pack_forget()
            self.issue_tree.configure(height=min(8, max(3, len(issues))))
            self.issues_holder.pack(fill='x')
            self.issue_buttons.pack(fill='x', pady=(8, 0))
        else:
            self.issues_holder.pack_forget()
            self.issue_buttons.pack_forget()
            self.issues_ok.pack(anchor='w')
        errors = sum(issue.level == 'error' for issue in issues)
        warnings = sum(issue.level == 'warning' for issue in issues)
        groups = sum(any(item.checked for item in card.items) for card in self.cards)
        checked = sum(item.checked for card in self.cards for item in card.items)
        pages = f'  ·  листов: примерно {1 + self.preview.pages} (титульный + {self.preview.pages} с таблицами)' \
            if self.preview is not None and self.preview.pages else ''
        self.check_summary.set(f'КК в паспорте: {groups} из {len(self.cards)}  ·  пунктов: {checked}{pages}\n'
                               f'Ошибок: {errors}  ·  проверить: {warnings}')
        mode = WORD_LAYOUT_MODES.get(self.settings.get('word_layout_mode'), '')
        self.preview_info.set(f'Режим переноса: {mode}' if mode else '')
        self.draw_preview()
        self.update_steps()

    def draw_preview(self):
        if self.current_step != 'check':
            return
        canvas = self.preview_canvas
        preview = self.preview
        if preview is None or not preview.segments:
            canvas.delete('all')
            canvas.create_text(20, 30, anchor='w', fill='#4B5563', font=('Arial', 11),
                               text='Схема появится, когда в картах будут отмечены пункты.')
            canvas.configure(scrollregion=(0, 0, 400, 60))
            return
        height = draw_layout_preview(canvas, preview, canvas.winfo_width())
        canvas.configure(scrollregion=(0, 0, max(int(canvas.winfo_width()), 300), int(height)))

    def go_to_issue(self, event=None):
        selection = self.issue_tree.selection() or (self.issue_tree.focus(),)
        try:
            issue = self.issues[int(selection[0])]
        except (IndexError, ValueError, TypeError):
            return 'break'
        if issue.step == 'title' or issue.card_index is None:
            self.show_step('title' if issue.step == 'title' else 'cards')
            return 'break'
        self.show_step('cards')
        if 0 <= issue.card_index < len(self.cards):
            self.activate_card(self.cards[issue.card_index])
            if issue.item_id and self.tree.exists(issue.item_id):
                self.tabs.select(0)
                self.tree.selection_set(issue.item_id)
                self.tree.see(issue.item_id)
                self.tree.focus(issue.item_id)
                self.tree.focus_set()
            elif 'обозначение' in issue.text.lower():
                self.code_entry.focus_set()
            elif 'наименование' in issue.text.lower():
                self.name_entry.focus_set()
        return 'break'

    # --- шаг 4: формирование Word ---------------------------------------------

    def _build_export_page(self):
        page = self._page_frame('export', 4, 'Формирование паспорта в Word',
                                'Выберите, как переносить таблицы, и нажмите зелёную кнопку. Файл проекта '
                                'при этом не меняется — паспорт можно сформировать сколько угодно раз.')
        self._page_nav(page, ('check', '←  Проверка'), None)
        outer, body = _scroll_area(page)
        outer.pack(fill='both', expand=True, padx=24)
        self.export_summary = tk.StringVar(value='')
        ttk.Label(body, textvariable=self.export_summary, style='Strong.TLabel').pack(anchor='w', pady=(0, 6))
        self.export_problems = tk.StringVar(value='')
        self.export_problems_label = ttk.Label(body, textvariable=self.export_problems, style='Warn.TLabel',
                                               wraplength=820, justify='left')
        self.mode_box = ttk.LabelFrame(body, text='Как переносить таблицы между колонками', padding=12,
                                       style='Section.TLabelframe')
        self.mode_box.pack(fill='x', pady=(6, 10))
        self.word_mode_buttons = []
        for mode, title in WORD_LAYOUT_MODES.items():
            button = ttk.Radiobutton(self.mode_box, text=title, value=mode, variable=self.word_layout_mode,
                                     command=self.change_word_layout_mode)
            button.pack(anchor='w', pady=(4, 0))
            self.word_mode_buttons.append(button)
            ttk.Label(self.mode_box, text=WORD_MODE_HINTS.get(mode, ''), style='Hint.TLabel', wraplength=760,
                      justify='left').pack(anchor='w', padx=24)
        self.word_check_button = ttk.Checkbutton(self.mode_box, text='Проверять переносы через Word',
                                                 variable=self.word_layout_check,
                                                 command=self.change_word_layout_check)
        self.word_check_button.pack(anchor='w', pady=(10, 0))
        ttk.Label(self.mode_box, text='Нужен установленный Microsoft Word: программа проверит реальные переносы строк. '
                  'Без Word — расчёт по оценке.', style='Hint.TLabel', wraplength=760,
                  justify='left').pack(anchor='w', padx=24)
        ttk.Button(body, text='Настроить размеры таблицы…',
                   command=self.open_table_layout_dialog).pack(anchor='w', pady=(0, 8))
        action_box = ttk.Frame(body)
        action_box.pack(fill='x', pady=(4, 0))
        word_outline = tk.Frame(action_box, background='#000000')
        word_outline.pack(fill='x')
        self.word_button = tk.Button(word_outline, text='Сформировать паспорт в Word', command=self.save,
                                     background=GREEN, activebackground=GREEN_ACTIVE, foreground='#000000',
                                     font=('Arial', 14, 'bold'), pady=14, relief='flat', borderwidth=0,
                                     highlightthickness=0, cursor='hand2')
        self.word_button.pack(fill='x', padx=1, pady=1)
        pdf_outline = tk.Frame(action_box, background='#000000')
        pdf_outline.pack(fill='x', pady=(6, 0))
        self.pdf_button = tk.Checkbutton(pdf_outline, text='Сохранять также PDF рядом с файлом Word',
                                         variable=self.pdf_export, command=self.change_pdf_export,
                                         background='#FFCDD2', activebackground='#EF9A9A',
                                         selectcolor='#FFFFFF', foreground='#000000',
                                         font=('Arial', 11), anchor='w', padx=10, pady=7, relief='flat',
                                         borderwidth=0, highlightthickness=0, cursor='hand2')
        self.pdf_button.pack(fill='x', padx=1, pady=1)
        tip(self.word_button, 'Создать файл паспорта .docx из отмеченных пунктов всех КК')
        tip(self.pdf_button, 'PDF сделает Word или LibreOffice. Если ни того, ни другого нет — '
                              'будет только Word.')
        self.export_progress = ttk.Frame(body)
        self.export_progress_text = tk.StringVar(value='')
        ttk.Label(self.export_progress, textvariable=self.export_progress_text, style='Hint.TLabel').pack(anchor='w')
        self.export_progressbar = ttk.Progressbar(self.export_progress, mode='indeterminate')
        self.export_progressbar.pack(fill='x', pady=(4, 0))
        self.result_box = ttk.LabelFrame(body, text='Готово', padding=12, style='Section.TLabelframe')
        self.result_text = tk.StringVar(value='')
        ttk.Label(self.result_box, textvariable=self.result_text, wraplength=820, justify='left').pack(anchor='w')
        buttons = ttk.Frame(self.result_box)
        buttons.pack(anchor='w', pady=(10, 0))
        ttk.Button(buttons, text='Открыть паспорт', command=lambda: self.open_export('docx')).pack(side='left')
        ttk.Button(buttons, text='Показать в папке', command=lambda: self.open_export('folder')).pack(side='left', padx=8)
        self.result_pdf_button = ttk.Button(buttons, text='Открыть PDF', command=lambda: self.open_export('pdf'))
        self.result_pdf_button.pack(side='left')

    def open_table_layout_dialog(self):
        if self.busy:
            self.wait_message()
            return
        win = tk.Toplevel(self.root)
        win.title('Размеры таблицы КК')
        win.transient(self.root)
        win.resizable(False, False)
        win.grab_set()
        frame = ttk.Frame(win, padding=14)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Ширина столбцов, мм', font=('Arial', 11, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w', pady=(0, 6))
        labels = ('№', 'Контролируемые параметры', 'Факт', 'ОК/НОК', 'Комментарии')
        layout = self.settings.get('table_layout', BUILTIN_TABLE_LAYOUT)
        variables = [tk.StringVar(value=f'{n:g}') for n in layout['widths']]
        for index, (label, var) in enumerate(zip(labels, variables), 1):
            ttk.Label(frame, text=label + ':').grid(row=index, column=0, sticky='w', padx=(0, 16), pady=3)
            ttk.Entry(frame, textvariable=var, width=12).grid(row=index, column=1, sticky='w')
        total = tk.StringVar()
        ttk.Label(frame, textvariable=total, style='Hint.TLabel').grid(row=6, column=0, columnspan=2, sticky='w', pady=(6, 8))
        def update_total(*_):
            try:
                result = sum(float(var.get().strip().replace(',', '.')) for var in variables)
                total.set(f'Сумма: {result:g} мм; требуется {TOTAL_WIDTH_MM:g} мм.')
            except ValueError:
                total.set(f'Сумма должна быть {TOTAL_WIDTH_MM:g} мм.')
        for var in variables:
            var.trace_add('write', update_total)
        update_total()
        ttk.Label(frame, text='Минимальная высота строк, мм', font=('Arial', 11, 'bold')).grid(
            row=7, column=0, columnspan=2, sticky='w', pady=(0, 6))
        header = tk.StringVar(value=f"{layout['header_min_mm']:g}")
        rows = tk.StringVar(value=f"{layout['row_min_mm']:g}")
        for row_no, label, var in ((8, 'Шапка', header), (9, 'Строки с данными', rows)):
            ttk.Label(frame, text=label + ':').grid(row=row_no, column=0, sticky='w', pady=3)
            ttk.Entry(frame, textvariable=var, width=12).grid(row=row_no, column=1, sticky='w')
        ttk.Label(frame, text='Текст сохраняет размер. При нехватке места строка увеличивается. '
                  'Длинные слова переносит Word; перенос по слогам зависит от его настроек.',
                  style='Hint.TLabel', wraplength=420, justify='left').grid(
                      row=10, column=0, columnspan=2, sticky='w', pady=(10, 10))
        def put(value):
            for var, number in zip(variables, value['widths']):
                var.set(f'{number:g}')
            header.set(f"{value['header_min_mm']:g}")
            rows.set(f"{value['row_min_mm']:g}")
        def candidate():
            try:
                result = {'widths': [float(var.get().strip().replace(',', '.')) for var in variables],
                          'header_min_mm': float(header.get().strip().replace(',', '.')),
                          'row_min_mm': float(rows.get().strip().replace(',', '.'))}
            except ValueError as exc:
                raise ValueError('Укажите числовые размеры в миллиметрах.') from exc
            return validate_table_layout(result)
        def apply(save_default=False):
            try:
                result = candidate()
                if save_default:
                    save_table_defaults(result)
            except (ValueError, OSError) as exc:
                messagebox.showwarning('Настройки таблицы', str(exc), parent=win)
                return
            if result != self.settings.get('table_layout'):
                self.remember_action()
                self.settings['table_layout'] = result
                self.update_project_title()
                self.update_export_page()
            if save_default:
                self.status.set('Размеры сохранены по умолчанию для новых проектов.')
            win.destroy()
        buttons = ttk.Frame(frame)
        buttons.grid(row=11, column=0, columnspan=2, sticky='ew')
        ttk.Button(buttons, text='Применить', command=apply).pack(side='left', padx=(0, 5))
        ttk.Button(buttons, text='Сохранить по умолчанию',
                   command=lambda: apply(True)).pack(side='left', padx=(0, 5))
        ttk.Button(buttons, text='Отмена', command=win.destroy).pack(side='left')
        reset = ttk.Frame(frame)
        reset.grid(row=12, column=0, columnspan=2, sticky='w', pady=(6, 0))
        ttk.Button(reset, text='Загрузить значения по умолчанию',
                   command=lambda: put(load_table_defaults())).pack(side='left', padx=(0, 5))
        ttk.Button(reset, text='Исходные размеры',
                   command=lambda: put(BUILTIN_TABLE_LAYOUT)).pack(side='left')

    def update_export_page(self):
        self.store_current_card()
        checked = sum(item.checked for card in self.cards for item in card.items)
        groups = sum(any(item.checked for item in card.items) for card in self.cards)
        self.export_summary.set(f"В паспорт пойдут {checked} {_plural(checked, 'пункт', 'пункта', 'пунктов')} "
                                f"из {groups} {_plural(groups, 'карты', 'карт', 'карт')} контроля.")
        errors = blocking(collect_issues(self.cards, self.settings))
        if errors:
            listed = '; '.join(issue.text for issue in errors[:3]) + (' …' if len(errors) > 3 else '')
            self.export_problems.set('⚠ Сначала исправьте: ' + listed + '\nПодробности — на шаге «Проверка».')
            self.export_problems_label.pack(anchor='w', pady=(0, 6), before=self.mode_box)
        else:
            self.export_problems.set('')
            self.export_problems_label.pack_forget()
        self.update_result_box()

    def update_result_box(self):
        result = self.last_export
        if not result or self.busy:
            self.result_box.pack_forget()
            return
        lines = [f"✓ Паспорт сохранён (пунктов: {result['count']}):", result['path']]
        if result.get('pdf'):
            lines.append('PDF: ' + result['pdf'])
        elif result.get('pdf_error'):
            lines.append('PDF не создан: ' + result['pdf_error'])
        if result.get('note'):
            lines.append(result['note'])
        if result.get('requested') and not result.get('checked'):
            lines.append('Переносы через Word не проверены — паспорт собран по расчёту.')
        if result.get('state') != self.project_state():
            lines.append('⚠ После формирования проект менялся — сформируйте паспорт ещё раз, '
                         'чтобы изменения попали в Word.')
        self.result_text.set('\n'.join(lines))
        if result.get('pdf'):
            self.result_pdf_button.pack(side='left')
        else:
            self.result_pdf_button.pack_forget()
        self.result_box.pack(fill='x', pady=(14, 0))

    def open_export(self, kind):
        result = self.last_export or {}
        path = result.get('pdf') if kind == 'pdf' else result.get('path')
        if not path or not Path(path).exists():
            messagebox.showwarning('Файл не найден', 'Файл перемещён или удалён:\n' + str(path or ''), parent=self.root)
            return
        try:
            if kind == 'folder':
                show_in_folder(path)
            else:
                open_path(path)
        except OSError as exc:
            messagebox.showerror('Не удалось открыть', str(exc), parent=self.root)

    def save(self):
        """Сформировать паспорт в Word (в фоне — окно не зависает)."""
        if self.busy:
            self.wait_message()
            return
        if not self.apply_title_form(silent=False):
            self.show_step('title')
            return
        self.store_current_card()
        checked = [item for card in self.cards for item in card.items if item.checked]
        if not checked:
            messagebox.showinfo('Нечего формировать', 'Не отмечено ни одного пункта. Отметьте галочками '
                                'нужные пункты на шаге «Карты контроля».')
            self.show_step('cards')
            return
        errors = blocking(collect_issues(self.cards, self.settings))
        if errors:
            listed = '\n'.join('• ' + issue.text for issue in errors[:6]) + ('\n…' if len(errors) > 6 else '')
            if messagebox.askyesno('Паспорт пока не готов', 'Сначала нужно исправить:\n\n' + listed +
                                   '\n\nПерейти к шагу «Проверка»?'):
                self.show_step('check')
            return
        review = [item for item in checked if item.warnings]
        if review:
            numbers = ', '.join(item.number for item in review[:12]) + (' …' if len(review) > 12 else '')
            if not messagebox.askyesno('Проверьте распознанные символы',
                                       f'Пунктов с пометкой «сверить с PDF»: {len(review)} ({numbers}).\n\n'
                                       'Всё равно сформировать паспорт?'):
                return
        if not self.ensure_project_name():
            return
        options = {'initialfile': passport_filename(self.settings['name'])}
        if self.project_path:
            options['initialdir'] = str(Path(self.project_path).parent)
        path = filedialog.asksaveasfilename(title='Сохранить паспорт', defaultextension='.docx',
                                            filetypes=[('Word', '*.docx')], **options)
        if not path:
            return
        cards = copy.deepcopy(self.cards)
        settings = copy.deepcopy(self.settings)
        state = self.project_state()
        count = len(checked)
        self.result_box.pack_forget()
        self.set_busy(True, f'Формирую паспорт ({count} пунктов)… Проверка через Word и PDF могут занять '
                            'до минуты.', exporting=True)
        def worker():
            try:
                result = export_passport_docx(path, cards, settings)
                self.export_events.put((True, path, result, count, state))
            except Exception as exc:  # ошибка показывается пользователю, окно не падает
                log.exception('Паспорт не сформирован: %s', path)
                self.export_events.put((False, path, exc, count, state))
        threading.Thread(target=worker, daemon=True).start()

    def _poll_export(self):
        try:
            ok, path, value, count, state = self.export_events.get_nowait()
        except queue.Empty:
            return
        self.set_busy(False)
        if not ok:
            text = str(value) if isinstance(value, (OSError, ValueError, RuntimeError)) else \
                f'{type(value).__name__}: {value}'
            self.status.set('Паспорт не сформирован.')
            messagebox.showerror('Паспорт не сформирован', text + '\n\nПодробности в журнале:\n' + str(log_path()))
            self.update_steps()
            return
        result = value if isinstance(value, dict) else {}
        pdf = result.get('pdf_path')
        self.last_export = {'path': path, 'pdf': str(pdf) if pdf else '', 'pdf_error': result.get('pdf_error') or '',
                            'note': result.get('note') or '', 'count': count, 'state': state,
                            'checked': bool(result.get('layout_checked')),
                            'requested': bool(result.get('layout_check_requested'))}
        self.status.set(f'Паспорт сохранён ({count} пунктов): {path}')
        if self.current_step == 'export':
            self.update_export_page()
        self.update_steps()
        problems = []
        if self.last_export['pdf_error']:
            problems.append('PDF не создан: ' + self.last_export['pdf_error'])
        if self.last_export['requested'] and not self.last_export['checked']:
            problems.append('Переносы через Word не проверены (Word недоступен) — паспорт собран по расчёту.')
        if problems:
            messagebox.showwarning('Паспорт сохранён с замечаниями', path + '\n\n' + '\n\n'.join(problems))
        elif self.current_step != 'export':
            messagebox.showinfo('Паспорт сохранён', path)

    # --- навигация по шагам -----------------------------------------------

    def show_step(self, key):
        if key not in self.pages:
            return
        if self.current_step == 'title' and key != 'title' and not self.apply_title_form(silent=True):
            try:
                changed = self._title_snapshot(self._form_candidate()) != self._title_snapshot(self.settings)
            except tk.TclError:
                changed = False
            if changed:
                problem = self.title_error.get().lstrip('⚠ ').strip()
                if messagebox.askyesno('Проверьте титульный лист', problem + '\n\nОстаться и исправить?\n'
                                       'Нет — перейти дальше, а неверные поля вернуть к прежним значениям.',
                                       parent=self.root):
                    return
                self.render_title_form()
        if key != 'start':
            self.store_current_card()
        for page in self.pages.values():
            page.pack_forget()
        self.pages[key].pack(fill='both', expand=True)
        self.current_step = key
        if key == 'start':
            self.update_start_page()
        elif key == 'title':
            self.update_title_preview()
        elif key == 'check':
            self.update_check_page()
        elif key == 'export':
            self.update_export_page()
        self.update_steps()

    def update_steps(self, state=None):
        buttons = getattr(self, 'step_buttons', None)
        if not buttons:
            return
        if state is None:
            state = self.project_state()
        issues = collect_issues(self.cards, self.settings)
        counts = count_by_step(issues)
        checked = sum(item.checked for card in self.cards for item in card.items)
        statuses = {}
        errors, warnings = counts.get('title', (0, 0))
        statuses['title'] = ('warn', '⚠ заполните обязательные поля') if errors else (
            ('warn', f'⚠ проверьте: {warnings}') if warnings else ('ok', '✓ заполнено'))
        errors, warnings = counts.get('cards', (0, 0))
        if not self.cards:
            statuses['cards'] = ('idle', 'добавьте PDF карт контроля')
        elif errors or warnings:
            statuses['cards'] = ('warn', f'⚠ замечаний: {errors + warnings}  ·  пунктов: {checked}')
        else:
            statuses['cards'] = ('ok', f'✓ КК: {len(self.cards)}, пунктов: {checked}')
        total = sum(e + w for e, w in counts.values())
        if not self.cards:
            statuses['check'] = ('idle', 'пока нечего проверять')
        else:
            statuses['check'] = ('warn', f'⚠ замечаний: {total}') if total else ('ok', '✓ замечаний нет')
        if self.last_export and self.last_export.get('state') == state:
            statuses['export'] = ('ok', '✓ паспорт сформирован')
        elif blocking(issues) or not checked:
            statuses['export'] = ('idle', 'сначала исправьте ошибки' if checked else 'отметьте пункты')
        else:
            statuses['export'] = ('idle', 'можно формировать')
        for key, (button, number, title) in buttons.items():
            kind, text = statuses.get(key, ('idle', ''))
            active = key == self.current_step
            button.configure(text=f'{number}. {title}\n      {text}', foreground=STEP_COLORS[kind],
                             background=SIDEBAR_ACTIVE if active else SIDEBAR_BG,
                             font=('Arial', 11, 'bold' if active else 'normal'))

    # --- справка ------------------------------------------------------------

    def show_help(self, key=None):
        key = key if key in HELP else (self.current_step or 'start')
        win = tk.Toplevel(self.root)
        win.title('Как пользоваться программой')
        win.geometry('680x560')
        win.minsize(480, 360)
        win.transient(self.root)
        footer = ttk.Frame(win, padding=(12, 8, 12, 12))
        footer.pack(side='bottom', fill='x')
        ttk.Button(footer, text='Закрыть', command=win.destroy).pack(side='right')
        ttk.Button(footer, text='Показать журнал', command=self.open_log_folder).pack(side='left')
        frame = ttk.Frame(win, padding=(12, 12, 12, 0))
        frame.pack(side='top', fill='both', expand=True)
        text = tk.Text(frame, wrap='word', font=('Arial', 11), relief='flat', padx=10, pady=8,
                       background='#FFFFFF')
        scroll = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
        scroll.pack(side='right', fill='y')
        text.pack(side='left', fill='both', expand=True)
        text.configure(yscrollcommand=scroll.set)
        text.tag_configure('head', font=('Arial', 13, 'bold'), spacing3=6)
        title, body = HELP[key]
        text.insert('end', title + '\n', 'head')
        text.insert('end', body + '\n\n')
        text.insert('end', 'Коротко о всей программе\n', 'head')
        text.insert('end', HELP_GENERAL)
        text.configure(state='disabled')
        win.bind('<Escape>', lambda event: win.destroy())
        win.focus_set()
        return win

    def open_log_folder(self):
        try:
            show_in_folder(log_path())
        except OSError as exc:
            messagebox.showerror('Журнал', f'Не удалось открыть папку журнала:\n{log_path()}\n\n{exc}',
                                 parent=self.root)

    def wait_message(self):
        messagebox.showinfo('Подождите', 'Дождитесь окончания загрузки PDF или формирования паспорта.',
                            parent=self.root)

    # --- горячие клавиши (работают и в русской раскладке) -------------------

    def _bind_shortcuts(self, root):
        root.bind('<Control-KeyPress>', self._shortcut)
        root.bind('<F1>', lambda event: (self.show_help(), 'break')[1])

    def _shortcut(self, event):
        try:
            if event.widget.winfo_toplevel() is not self.root:
                return None  # в диалогах — их собственные клавиши
        except (AttributeError, tk.TclError):
            pass
        if sys.platform == 'win32':
            key = SHORTCUT_KEYCODES.get(getattr(event, 'keycode', None))
        else:
            key = SHORTCUT_KEYSYMS.get(str(getattr(event, 'keysym', '')).lower())
        if key is None:
            return None
        shift = bool(getattr(event, 'state', 0) & 1)
        try:
            in_text = event.widget.winfo_class() == 'Text'
        except (AttributeError, tk.TclError):
            in_text = False
        if key in ('z', 'y') and in_text:
            return None  # у многострочного поля своя отмена
        if key == 's':
            self.save_project_as() if shift else self.save_project()
        elif key == 'o':
            self.start_open()
        elif key == 'n':
            self.start_new()
        elif key == 'z':
            self.redo_action() if shift else self.undo_action()
        elif key == 'y':
            self.redo_action()
        elif key in '1234':
            self.show_step(STEPS[int(key) - 1][0])
        return 'break'

    # --- отмена и повтор -----------------------------------------------------

    def undo_action(self):
        self.apply_title_form()
        if self.busy or not self.undo_stack:
            return
        current = self.project_state()
        selected_now = self.current_card.id if self.current_card else None
        while self.undo_stack:
            cards, settings, selected = self.undo_stack.pop()
            if project_payload(cards, settings) != current:
                self.redo_stack.append(copy.deepcopy((self.cards, self.settings, selected_now)))
                self.redo_stack = self.redo_stack[-30:]
                self.cards, self.settings = cards, settings
                self._restore_view(selected)
                self.status.set('Действие отменено.')
                return
        self.update_project_title()

    def redo_action(self):
        self.apply_title_form()
        self.update_project_title()  # новая правка после отмены сбрасывает «Повторить»
        if self.busy or not self.redo_stack:
            return
        selected_now = self.current_card.id if self.current_card else None
        self.undo_stack.append(copy.deepcopy((self.cards, self.settings, selected_now)))
        self.undo_stack = self.undo_stack[-30:]
        cards, settings, selected = self.redo_stack.pop()
        self.cards, self.settings = cards, settings
        self._restore_view(selected)
        self.status.set('Действие повторено.')

    def _restore_view(self, selected):
        self.redo_base = None
        self.reset_project_view(selected)
        self.redo_base = self.project_state()
        self.update_project_title()
        if self.current_step == 'check':
            self.update_check_page()
        elif self.current_step == 'export':
            self.update_export_page()
        elif self.current_step == 'title':
            self.update_title_preview()

    # --- заголовок окна и состояние сохранения ---------------------------

    def update_project_title(self):
        self._sync_word_controls()
        state = self.project_state()
        dirty = state != self.saved_state
        if self.redo_base is not None and state != self.redo_base:
            self.redo_stack = []
            self.redo_base = None
        name = self.settings['name'] or 'Новый проект'
        self.project_label.set(name)
        self.root.title(('* ' if dirty else '') + name + ' — Паспорт КК')
        if dirty:
            text, color = '● Есть несохранённые изменения', '#FDE68A'
            if self.last_autosave_at is not None:
                text += ' · автокопия ' + self.last_autosave_at.strftime('%H:%M')
        elif self.last_saved_at is not None:
            text, color = '✓ Сохранено в ' + self.last_saved_at.strftime('%H:%M'), '#A7F3D0'
        elif self.project_path:
            text, color = '✓ Все изменения сохранены', '#A7F3D0'
        else:
            text, color = '', '#A7F3D0'
        self.save_state.set(text)
        self.save_state_label.configure(foreground=color)
        self.undo_button.configure(state='normal' if self.undo_stack and not self.busy else 'disabled')
        self.redo_button.configure(state='normal' if self.redo_stack and not self.busy else 'disabled')
        if self.title_form is not None and self._title_snapshot(self.settings) != self._title_form_source:
            self.render_title_form()  # настройки сменились снаружи (отмена, восстановление)
            if self.current_step == 'title':
                self.update_title_preview()
        if self.current_step == 'export':
            self.update_result_box()
        self.update_steps(state)

    # --- новый / открыть проект -----------------------------------------------

    def _reset_history(self):
        self.undo_stack = []
        self.redo_stack = []
        self.redo_base = None
        self.last_export = None
        self.last_saved_at = None

    def new_project(self):
        if not self.confirm_leave_project():
            return False
        self.cards = []
        self.settings = default_project_settings()
        self.project_path = ''
        self._reset_history()
        self.saved_state = project_payload([], self.settings)
        self.reset_project_view()
        self.render_title_form()
        self.autosave_now()
        self.update_project_title()
        self.status.set('Новый проект: заполните титульный лист и добавьте PDF карт контроля.')
        return True

    def open_project(self, path=None):
        if not self.confirm_leave_project():
            return False
        if not path:
            options = {}
            if self.recent:
                options['initialdir'] = str(Path(self.recent[0]).parent)
            path = filedialog.askopenfilename(title='Открыть проект',
                                              filetypes=[('Проект паспорта', '*.kkproject')], **options)
        if not path:
            return False
        try:
            cards, settings = read_project_file(path)
        except (OSError, ValueError, TypeError) as exc:
            log.warning('Проект не открыт: %s', path, exc_info=True)
            messagebox.showerror('Проект не открыт', str(exc))
            return False
        self.cards, self.settings = cards, settings
        self.project_path = path
        self._reset_history()
        self.saved_state = project_payload(cards, settings)
        self.reset_project_view()
        self.saved_state = self.project_state()
        self.render_title_form()
        self.autosave_now()
        try:
            add_recent(path)
        except OSError:
            log.warning('Список недавних проектов не обновлён', exc_info=True)
        self.update_project_title()
        self.status.set('Проект открыт: ' + path)
        return True


    @staticmethod
    def _set_window_icon(root):
        icon = ASSETS_DIR / 'app.ico'
        if os.name == 'nt' and icon.is_file():
            try:
                root.iconbitmap(default=str(icon))
            except tk.TclError:
                log.debug('Значок окна не установлен', exc_info=True)

    # --- автосохранение и восстановление ------------------------------------------
    def autosave_now(self):
        """Записать копию несохранённого проекта; если правок нет — убрать копию."""
        session = self.autosave
        if session is None or not session.enabled:
            return False
        try:
            state = self.project_state()
            if state == self.saved_state:
                session.discard()
                return False
            saved = session.save(self.cards, self.settings, self.project_path, payload=state)
            if saved:
                self.last_autosave_at = datetime.datetime.now()
            return saved
        except (OSError, ValueError, TypeError):
            log.warning('Автосохранение не удалось', exc_info=True)
            return False

    def autosave_tick(self):
        try:
            if not self.busy:
                self.apply_title_form()
                self.autosave_now()
                self.update_project_title()
        finally:
            self.root.after(AUTOSAVE_SECONDS * 1000, self.autosave_tick)

    def keep_discarded_copy(self):
        """«Нет» в вопросе о сохранении: последняя такая копия открывается из меню «Проект»."""
        if self.autosave is None:
            return
        try:
            self.autosave.keep_last_closed(self.cards, self.settings, self.project_path)
        except (OSError, ValueError, TypeError):
            log.warning('Не удалось сохранить копию закрытого проекта', exc_info=True)

    def apply_recovered(self, cards, settings, project_path, description):
        """Показать восстановленный проект; он считается несохранённым, пока его не сохранят."""
        self.cards, self.settings = cards, settings
        self.project_path = project_path if project_path and Path(project_path).is_file() else ''
        self._reset_history()
        self.saved_state = None
        self.reset_project_view()
        self.render_title_form()
        saved = self.autosave_now()
        self.update_project_title()
        self.status.set('Проект восстановлен из копии: ' + description.splitlines()[0] +
                        '. Сохраните его кнопкой «Сохранить проект».')
        self.show_step('cards')
        return saved

    def restore_entry(self, entry):
        try:
            cards, settings = entry.load()
        except (OSError, ValueError, TypeError) as exc:
            log.warning('Копия не восстановлена: %s', entry.project_file, exc_info=True)
            mark_broken(entry)
            messagebox.showerror('Копия повреждена', f'Не удалось восстановить проект: {exc}\n\n'
                                 f'Файл оставлен для разбора:\n{entry.project_file}.broken')
            return False
        try:
            if self.apply_recovered(cards, settings, entry.project_path, entry.describe()):
                remove_entry(entry)  # данные уже в этом окне и в его автосохранении
            else:
                set_aside(entry)  # автосохранение не записалось — старую копию не теряем
        except OSError:
            log.warning('Не удалось убрать восстановленную копию', exc_info=True)
        return True

    def offer_recovery(self):
        """При запуске: предложить проект окна, закрытого аварийно."""
        if self.autosave is None or not self.autosave.enabled:
            return
        try:
            entries = find_recoverable(self.autosave.folder)
        except OSError:
            log.warning('Не удалось проверить автосохранения', exc_info=True)
            return
        if not entries:
            return
        entry = entries[0]
        more = (f'\n\nЕсть ещё копий: {len(entries) - 1} — о них программа спросит при следующем запуске.'
                if len(entries) > 1 else '')
        answer = messagebox.askyesnocancel('Восстановление проекта',
            'Программа была закрыта аварийно, остался несохранённый проект:\n\n'
            + entry.describe() + '\n\nВосстановить его?\n\n'
            'Да — открыть копию.\n'
            'Нет — не открывать (копия останется в меню «Проект → '
            'Последний несохранённый проект…»).\n'
            'Отмена — спросить при следующем запуске.' + more)
        if answer is None:
            return
        if answer:
            self.restore_entry(entry)
            return
        try:
            set_aside(entry)
        except OSError:
            log.warning('Не удалось отложить копию %s', entry.project_file, exc_info=True)

    def open_last_closed(self):
        """Открыть проект, закрытый без сохранения (или отложенную аварийную копию)."""
        if self.busy:
            self.wait_message()
            return
        entry = last_closed_entry(self.autosave.folder) if self.autosave is not None else None
        if entry is None:
            messagebox.showinfo('Нет копии', 'Нет проекта, закрытого без сохранения.')
            return
        try:
            cards, settings = entry.load()  # до вопроса: «Нет» ниже перезапишет эту копию
        except (OSError, ValueError, TypeError) as exc:
            log.warning('Копия не открыта: %s', entry.project_file, exc_info=True)
            messagebox.showerror('Копия не открыта', str(exc))
            return
        if not messagebox.askyesno('Последний несохранённый проект',
                                   entry.describe() + '\n\nОткрыть эту копию?'):
            return
        if not self.confirm_leave_project():
            return
        self.apply_recovered(cards, settings, entry.project_path, entry.describe())

    def _sync_word_controls(self):
        """Переключатели Word следуют за настройками проекта (открытие, отмена)."""
        self.word_layout_mode.set(self.settings.get('word_layout_mode', 'free'))
        self.word_layout_check.set(self.settings.get('word_layout_check', False))
        self.pdf_export.set(self.settings.get('pdf_export', False))

    def change_word_layout_mode(self):
        mode = self.word_layout_mode.get()
        previous = self.settings.get('word_layout_mode', 'free')
        if self.busy or mode not in WORD_LAYOUT_MODES:
            self.word_layout_mode.set(previous)
            return
        if mode == previous:
            return
        self.remember_action()
        self.settings['word_layout_mode'] = mode
        self.update_project_title()

    def change_word_layout_check(self):
        enabled = bool(self.word_layout_check.get())
        previous = self.settings.get('word_layout_check', False)
        if self.busy:
            self.word_layout_check.set(previous)
            return
        if enabled == previous:
            return
        self.remember_action()
        self.settings['word_layout_check'] = enabled
        self.update_project_title()

    def change_pdf_export(self):
        enabled = bool(self.pdf_export.get())
        previous = self.settings.get('pdf_export', False)
        if self.busy:
            self.pdf_export.set(previous)
            return
        if enabled == previous:
            return
        self.remember_action()
        self.settings['pdf_export'] = enabled
        self.update_project_title()

    def project_state(self):
        self.store_current_card()
        return project_payload(self.cards, self.settings)

    def remember_action(self):
        self.store_current_card()
        selected = self.current_card.id if self.current_card else None
        if self.undo_stack:
            cards, settings, _ = self.undo_stack[-1]
            if project_payload(cards, settings) == project_payload(self.cards, self.settings):
                # Repeated focus/clicks without edits must not evict real undo actions.
                self.undo_stack[-1] = (cards, settings, selected)
                return False
        self.undo_stack.append(copy.deepcopy((self.cards, self.settings,
                                              selected)))
        self.undo_stack = self.undo_stack[-30:]
        return True

    def reset_project_view(self, selected=None):
        self._sync_word_controls()
        self.current_card = None
        self.items = []
        self.source = ''
        self.details = {'control_card': False, 'graphic_rows': {}, 'warnings': [], 'document_code': ''}
        self.document_code.set('')
        self.document_name.set('')
        self.raw.delete('1.0', 'end')
        self.raw.edit_reset()
        if self.cards:
            self.activate_card(next((card for card in self.cards if card.id == selected), self.cards[0]))
        else:
            self.refresh()

    def ensure_project_name(self):
        if self.settings['name'].strip():
            return True
        return self.rename_project()

    def rename_project(self):
        if self.busy:
            self.wait_message()
            return False
        suggestion = self.settings['name'] or (self.document_name.get().strip() + ' ' + self.document_code.get().strip()).strip()
        name = simpledialog.askstring('Название проекта',
            'Наименование и обозначение изделия\n(это название проекта и строка титульника).\nНапример, Рама ТМ10.1133.000:',
            initialvalue=suggestion, parent=self.root)
        if name is None or not name.strip():
            return False
        try:
            candidate = validate_project_settings({**self.settings, 'name': name.strip(),
                'passport_name': self.settings['passport_name'] if name.strip() == self.settings['name'] else ''})
        except ValueError as exc:
            messagebox.showwarning('Название проекта', str(exc), parent=self.root)
            return False
        if candidate != self.settings:
            self.remember_action()
            self.settings = candidate
            self.update_project_title()
        return True

    def save_project(self, save_as=False):
        if self.busy:
            self.wait_message()
            return False
        self.apply_title_form()
        if not self.ensure_project_name():
            return False
        path = self.project_path
        if save_as or not path:
            options = {'initialfile': project_filename(self.settings['name'], '.kkproject')}
            if self.project_path:
                # «Сохранить как» открывает папку и имя текущего файла проекта.
                current = Path(self.project_path)
                options = {'initialdir': str(current.parent), 'initialfile': current.name}
            path = filedialog.asksaveasfilename(
                title='Сохранить проект как' if save_as else 'Сохранить проект',
                defaultextension='.kkproject', filetypes=[('Проект паспорта', '*.kkproject')],
                **options)
        if not path:
            return False
        self.store_current_card()
        try:
            write_project_file(path, self.cards, self.settings)
        except (OSError, ValueError, TypeError) as exc:
            log.warning('Проект не сохранён: %s', path, exc_info=True)
            messagebox.showerror('Проект не сохранён', str(exc))
            return False
        self.project_path = path
        self.saved_state = self.project_state()
        self.autosave_now()  # правок больше нет — копия удаляется
        self.last_saved_at = datetime.datetime.now()
        try:
            add_recent(path)
        except OSError:
            log.warning('Список недавних проектов не обновлён', exc_info=True)
        self.update_project_title()
        self.status.set('Проект сохранён: ' + path)
        return True

    def save_project_as(self):
        """Сохранить копию проекта под другим именем; дальше работа идёт в новом файле."""
        return self.save_project(True)

    def confirm_leave_project(self):
        if self.busy:
            self.wait_message()
            return False
        self.apply_title_form()
        if self.project_state() == self.saved_state:
            return True
        answer = messagebox.askyesnocancel('Несохранённый проект',
            'Сохранить изменения проекта перед выходом или сменой проекта?\n\n'
            'Да — сохранить, Нет — продолжить без сохранения, Отмена — остаться.')
        if answer is None:
            return False
        if answer:
            return self.save_project()
        self.keep_discarded_copy()
        return True

    def close_app(self):
        if self.confirm_leave_project():
            if self.autosave is not None:
                self.autosave.close()
            self.root.destroy()

    def store_current_card(self):
        if self.current_card is not None:
            self.current_card.items = self.items
            self.current_card.document_code = self.document_code.get().strip()
            self.current_card.document_name = self.document_name.get().strip()
            self.current_card.raw = self.raw.get("1.0", "end-1c")

    def activate_card(self, card):
        self.check_anchor = None
        self.check_anchor_value = True
        self.store_current_card()
        self.current_card = card
        self.items, self.source, self.details = card.items, card.source, card.details
        self.document_code.set(card.document_code)
        self.document_name.set(card.document_name)
        self.raw.delete("1.0", "end")
        self.raw.insert("1.0", card.raw)
        self.raw.edit_reset()
        self.refresh()

    def move_current_card(self, offset=0, position=None):
        if self.busy or self.current_card is None:
            return False
        index = next((n for n, card in enumerate(self.cards) if card is self.current_card), None)
        if index is None:
            return False
        target = index + offset if position is None else position - 1
        if not isinstance(target, int) or not 0 <= target < len(self.cards) or target == index:
            return False
        self.remember_action()
        self.store_current_card()
        card = self.cards.pop(index)
        self.cards.insert(target, card)
        self.refresh()
        return True

    def choose_card_position(self):
        if self.busy or self.current_card is None or len(self.cards) < 2:
            return
        position = simpledialog.askinteger('Порядок карт контроля',
            f'На какое место поставить текущую КК?\nВведите число от 1 до {len(self.cards)}.',
            initialvalue=self.cards.index(self.current_card) + 1,
            minvalue=1, maxvalue=len(self.cards), parent=self.root)
        if position is not None:
            self.move_current_card(position=position)

    def remove_current_card(self):
        if self.busy:
            self.wait_message()
            return
        card = self.current_card
        if card is None:
            return
        index = next((n for n, value in enumerate(self.cards) if value is card), None)
        if index is None:
            return
        label = (self.document_code.get().strip() + " " +
                 self.document_name.get().strip()).strip() or Path(card.source).name
        if not messagebox.askyesno(
                "Удалить КК из списка?",
                f"Удалить «{label}» из рассмотрения?\n\n"
                "Галочки и правки этой КК будут удалены из программы. "
                "Остальные КК сохранятся.\nИсходный PDF на диске не удаляется."):
            return
        self.remember_action()
        self.store_current_card()
        self.cards.pop(index)
        self.current_card = None
        if self.cards:
            self.activate_card(self.cards[min(index, len(self.cards) - 1)])
        else:
            self.items = []
            self.source = ""
            self.details = {"control_card": False, "graphic_rows": {},
                            "document_code": "", "warnings": []}
            self.document_code.set("")
            self.document_name.set("")
            self.raw.delete("1.0", "end")
            self.raw.edit_reset()
            self.refresh()

    def selected_item(self):
        selection = self.tree.selection()
        return next((i for i in self.items if selection and i.id == selection[0]), None)

    def set_all(self, value):
        if any(item.checked != value for item in self.items):
            self.remember_action()
        for item in self.items:
            item.checked = value
        self.refresh()

    def deselect_welds(self):
        if any(item.checked and WELD_WORD.search(item.text) for item in self.items):
            self.remember_action()
        deselect_weld_items(self.items)
        self.refresh()

    def choose_numbers(self):
        if self.busy or self.current_card is None:
            return
        expression = simpledialog.askstring('Быстрый выбор пунктов текущей КК',
            'Введите номера: например, 1-10 или 1,2,3-10,15.\n'
            'Это номера пунктов, не позиции строк. Остальные галочки не меняются.\n'
            'Подпункты можно указать отдельно: 1.2,1.3.', parent=self.root)
        if expression is None:
            return
        try:
            self.mark_number_expression(expression)
        except ValueError as exc:
            messagebox.showwarning('Проверьте номера', str(exc), parent=self.root)

    def mark_number_expression(self, expression):
        if self.busy or self.current_card is None:
            return False
        selected, missing_count, missing = select_number_expression(expression, self.items)
        ids = set(selected)
        if any(item.id in ids and not item.checked for item in self.items):
            self.remember_action()
            for item in self.items:
                if item.id in ids:
                    item.checked = True
            self.refresh()
        if missing_count:
            messagebox.showwarning('Не все номера найдены',
                ('Отмечены существующие пункты текущей КК.\n' if selected else
                 'Подходящих пунктов нет. Галочки не изменены.\n') +
                f'Не найдено номеров: {missing_count}.\n' + ', '.join(missing) +
                (' …' if missing_count > len(missing) else ''), parent=self.root)
        self.check_anchor = None
        return bool(selected)

    def toggle(self, iid, extend=False):
        if self.busy:
            return
        item = next((i for i in self.items if i.id == iid), None)
        if item is None:
            return
        order = [i.id for i in self.items]
        anchor = getattr(self, 'check_anchor', None)
        if extend and anchor in order:
            a, b = sorted((order.index(anchor), order.index(iid)))
            affected = self.items[a:b + 1]
            value = getattr(self, 'check_anchor_value', True)
        else:
            affected, value = [item], not item.checked
            self.check_anchor = iid
            self.check_anchor_value = value
        if any(i.checked != value for i in affected):
            self.remember_action()
            for i in affected:
                i.checked = value
            self.refresh()
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.tree.see(iid)

    def click(self, event):
        if self.tree.identify_region(event.x, event.y) == "cell" and self.tree.identify_column(event.x) == "#1":
            self.toggle(self.tree.identify_row(event.y), extend=bool(getattr(event, 'state', 0) & 1))
            return "break"

    def toggle_focus(self, event):
        item = self.selected_item()
        if item:
            self.toggle(item.id, extend=bool(getattr(event, 'state', 0) & 1))
        return "break"

    def double_click(self, event):
        if self.tree.identify_column(event.x) != "#1":
            iid = self.tree.identify_row(event.y)
            if iid:
                self.tree.selection_set(iid)
                self.edit_item(self.selected_item())

    def add_item(self):
        if self.current_card is None:
            messagebox.showinfo("Нет КК", "Сначала добавьте карту контроля.")
            return
        anchor = self.selected_item()
        self.edit_item(Item(number="", text="", checked=True, manual=True), new=True,
                       anchor_id=anchor.id if anchor else None)

    def delete_item(self):
        item = self.selected_item()
        if item and messagebox.askyesno("Удалить", "Удалить выбранную строку из чек-листа?"):
            self.remember_action()
            self.items.remove(item)
            self.refresh()

    def edit_item(self, item, new=False, anchor_id=None):
        if not item:
            return
        win = tk.Toplevel(self.root)
        win.title("Добавить свой пункт" if new and item.manual else "Текст пункта")
        win.geometry("720x500" if new and item.manual else "720x440")
        win.minsize(520, 360)
        win.transient(self.root)
        win.grab_set()
        # Pack footer FIRST: it always reserves space at the bottom of the window.
        footer = ttk.Frame(win, padding=(12, 8, 12, 12))
        footer.pack(side="bottom", fill="x")
        frame = ttk.Frame(win, padding=(12, 12, 12, 0))
        frame.pack(side="top", fill="both", expand=True)
        if item.warnings:
            ttk.Label(frame, text="\n".join(item.warnings), wraplength=670).pack(anchor="w", pady=(0, 8))
        if item.source_crop:
            def show_crop():
                webbrowser.open(Path(item.source_crop).resolve().as_uri())
            ttk.Button(frame, text="Показать исходный пункт из PDF", command=show_crop).pack(anchor="w")
        if item.original_text:
            ttk.Label(frame, text="Исходный текст (для сверки):").pack(anchor="w", pady=(6, 0))
            original = tk.Text(frame, height=3, wrap="word", font=("Arial", 10))
            original.insert("1.0", item.original_text)
            original.configure(state="disabled")
            original.pack(fill="x", pady=(2, 6))
        position = tk.StringVar(value='after')
        if new and item.manual:
            if anchor_id is not None:
                placement = ttk.Frame(frame)
                placement.pack(fill='x', pady=(0, 8))
                ttk.Label(placement, text="Место:").pack(side='left', padx=(0, 8))
                ttk.Radiobutton(placement, text="Перед выбранным пунктом", variable=position,
                                value='before').pack(side='left', padx=(0, 10))
                ttk.Radiobutton(placement, text="После выбранного пункта", variable=position,
                                value='after').pack(side='left')
            else:
                ttk.Label(frame, text="Пункт будет добавлен в конец списка.",
                          style='Hint.TLabel').pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="Номер или обозначение (необязательно):" if item.manual else "Номер:").pack(anchor="w")
        number = ttk.Entry(frame)
        number.insert(0, item.number)
        number.pack(fill="x", pady=(4, 12))
        text = tk.Text(frame, height=6, wrap="word", font=("Arial", 12), undo=True)
        text.insert("1.0", item.text)
        text.pack(fill="both", expand=True)
        ocr_state = {'text': item.ocr_text, 'status': item.ocr_status}
        if item.ocr_crops or item.ocr_text or item.ocr_status:
            ttk.Button(footer, text="OCR…", command=lambda: open_ocr_dialog(win, item, number, text, ocr_state)).pack(side="left")
        def accept():
            recorded = self.remember_action()
            try:
                insert_index = None
                if new and item.manual and anchor_id is not None:
                    anchor_index = next((i for i, existing in enumerate(self.items)
                                         if existing.id == anchor_id), None)
                    if anchor_index is None:
                        raise ValueError("Выбранный пункт больше не найден.")
                    insert_index = anchor_index + (position.get() == 'after')
                commit_item(self.items, item, number.get(), text.get("1.0", "end-1c"),
                            new, insert_index=insert_index)
            except ValueError as exc:
                if recorded:
                    self.undo_stack.pop()
                messagebox.showwarning("Пункт не сохранён", str(exc), parent=win)
                return
            item.ocr_text = ocr_state['text']
            item.ocr_status = ocr_state['status']
            self.refresh()
            if self.tree.exists(item.id):
                self.tree.selection_set(item.id)
                self.tree.focus(item.id)
                self.tree.see(item.id)
            win.destroy()
        ttk.Button(footer, text="Сохранить", command=accept).pack(side="right")

    def rebuild(self, every_line):
        if self.items and not messagebox.askyesno("Пересборка", "Заменить список? Правки строк и галочки будут сброшены."):
            return
        if self.current_card is None:
            return
        candidate = build_card_items(self.details, self.raw.get("1.0", "end-1c"), every_line)
        duplicates = duplicate_numbers(candidate)
        if duplicates:
            messagebox.showwarning("Повторные номера", "Пересборка отменена: номера " + ", ".join(duplicates) +
                                   " повторяются. Исправьте исходный текст; прежний список сохранён.")
            return
        self.remember_action()
        self.items = candidate
        self.refresh()
        self.tabs.select(0)

def _report_callback_exception(exc_type, exc, tb, app=None):
    """Непредвиденная ошибка в обработчике кнопки: копия проекта, журнал и понятное сообщение."""
    log.error('Ошибка в окне программы', exc_info=(exc_type, exc, tb))
    where = log_path()
    details = f'\n\nПодробности записаны в журнал:\n{where}' if where else ''
    if app is not None:
        try:
            if app.autosave_now():
                details += '\n\nНесохранённый проект записан в автосохранение.'
        except Exception:
            log.warning('Автосохранение после ошибки не удалось', exc_info=True)
    messagebox.showerror('Непредвиденная ошибка', f'{exc_type.__name__}: {exc}{details}')


def run_gui() -> None:
    if tk is None:
        raise SystemExit("Нужен tkinter. В Windows установите Python с компонентом Tcl/Tk; "
                         "в Ubuntu/Debian: sudo apt install python3-tk")
    root = tk.Tk()
    root.report_callback_exception = _report_callback_exception
    install_clipboard_support(root)
    app = App(root)
    root.report_callback_exception = functools.partial(_report_callback_exception, app=app)
    root.mainloop()


def main() -> None:
    """Точка входа: журнал, уборка старых временных папок, окно."""
    setup_logging()
    removed = cleanup_stale()
    log.info('Запуск (Python %s); удалено старых временных папок: %s', sys.version.split()[0], removed)
    run_gui()
