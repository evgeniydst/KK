"""Подготовка и пропорциональное размещение изображения изделия."""
from __future__ import annotations

import base64
import binascii
import io
import warnings
from pathlib import Path

from PIL import Image, ImageGrab, ImageOps, UnidentifiedImageError

# Ограничения защищают память приложения и размер файла проекта.
MAX_SOURCE_BYTES = 20 * 1024 * 1024
MAX_PIXELS = 40_000_000
MAX_STORED_BYTES = 12 * 1024 * 1024
MAX_SIDE = 3000
PICTURE_WIDTH_MM = 124
PICTURE_PADDING_X_MM = 4  # запас с учётом внутренних полей таблицы Word
PICTURE_PADDING_Y_MM = 2


def fit_size(width, height, max_width, max_height):
    """Пропорционально вписать изображение; размеры заданы в любой общей единице."""
    if min(width, height, max_width, max_height) <= 0:
        raise ValueError('Некорректные размеры изображения или рамки.')
    factor = min(max_width / width, max_height / height)
    return width * factor, height * factor


def _prepare(image):
    if image.width * image.height > MAX_PIXELS:
        raise ValueError('Изображение слишком велико (максимум 40 мегапикселей).')
    image = ImageOps.exif_transpose(image)
    image.load()
    image.thumbnail((MAX_SIDE, MAX_SIDE), Image.Resampling.LANCZOS)
    if image.mode not in ('RGB', 'RGBA'):
        image = image.convert('RGBA' if 'transparency' in image.info else 'RGB')
    output = io.BytesIO()
    image.save(output, format='PNG', optimize=True)
    data = output.getvalue()
    if len(data) > MAX_STORED_BYTES:
        raise ValueError('Изображение слишком велико для проекта. Уменьшите его и попробуйте снова.')
    return base64.b64encode(data).decode('ascii')


def from_file(path):
    try:
        if Path(path).stat().st_size > MAX_SOURCE_BYTES:
            raise ValueError('Файл изображения слишком велик (максимум 20 МБ).')
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(path) as image:
                return _prepare(image)
    except (OSError, UnidentifiedImageError, Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise ValueError('Не удалось прочитать изображение. Выберите корректный файл PNG, JPEG, BMP, GIF или TIFF.') from exc


def from_clipboard():
    try:
        image = ImageGrab.grabclipboard()
        if isinstance(image, list):
            if not image:
                raise ValueError('В буфере обмена нет изображения.')
            return from_file(image[0])
        if not isinstance(image, Image.Image):
            raise ValueError('В буфере обмена нет изображения. Скопируйте рисунок или файл изображения.')
        return _prepare(image)
    except (OSError, NotImplementedError) as exc:
        raise ValueError('Не удалось получить изображение из буфера обмена. Попробуйте загрузить его с диска.') from exc


def decode_image(encoded):
    if not isinstance(encoded, str) or len(encoded) > ((MAX_STORED_BYTES + 2) // 3) * 4:
        raise ValueError('Некорректное изображение титульного листа.')
    try:
        data = base64.b64decode(encoded, validate=True)
        if not data.startswith(b'\x89PNG\r\n\x1a\n'):
            raise ValueError('Изображение титульного листа должно быть PNG.')
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(data)) as image:
                if image.width * image.height > MAX_PIXELS:
                    raise ValueError('Изображение титульного листа слишком велико.')
                image.load()
                return image.copy()
    except (binascii.Error, OSError, Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise ValueError('Повреждено изображение титульного листа.') from exc
