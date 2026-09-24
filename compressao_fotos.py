"""Compressão JPEG compartilhada pelo renomeador e pelo utilitário de fotos."""

import os
import tempfile

from PIL import Image, ImageOps


PHOTO_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic"}


def compressed_name(filename):
    """Mantém o nome de JPEGs e troca apenas a extensão dos demais formatos."""
    stem, extension = os.path.splitext(filename)
    if extension.lower() not in PHOTO_EXTENSIONS:
        raise ValueError(f"Formato não suportado: {filename}")
    return filename if extension.lower() in {".jpg", ".jpeg"} else f"{stem}.JPG"


def save_compressed_jpeg(image, target):
    """Usa exatamente os parâmetros de compressão do renomeador."""
    image.thumbnail((800, 800), Image.Resampling.BILINEAR)
    image.convert("RGB").save(target, "JPEG", optimize=False, quality=60)


def compress_photo(source, target):
    """Grava primeiro um temporário para não deixar um resultado parcial."""
    output_dir = os.path.dirname(target)
    fd, temporary = tempfile.mkstemp(prefix=".ecowave-", suffix=".jpg", dir=output_dir)
    os.close(fd)
    try:
        with Image.open(source) as image:
            upright = ImageOps.exif_transpose(image)
            save_compressed_jpeg(upright, temporary)
        os.replace(temporary, target)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)
