"""Render PDF pages to PNG images in document order."""
import re
from pathlib import Path

import pymupdf

from .service import MAX_FILE

MAX_PAGES = 200
MAX_PIXELS = 50_000_000
DPI_CHOICES = (150, 300)
PART = re.compile(r'^(\d+)(?:\s*-\s*(\d+))?$')
RANGE_HINT = '页码范围格式不正确，请使用如 1-5,8 的写法'


def parse_page_spec(spec, total):
    spec = (spec or '').strip()
    if not spec:
        if total > MAX_PAGES:
            raise ValueError(f'单次最多导入 {MAX_PAGES} 页，请填写页码范围')
        return list(range(1, total + 1))
    pages = set()
    for part in re.split(r'[,，;；]', spec):
        part = part.strip()
        if not part:
            continue
        match = PART.match(part)
        if not match:
            raise ValueError(RANGE_HINT)
        start = int(match.group(1))
        end = int(match.group(2) or start)
        if start < 1 or end > total or start > end:
            raise ValueError(f'页码超出范围，当前 PDF 共 {total} 页')
        if end - start + 1 > MAX_PAGES:
            raise ValueError(f'单次最多导入 {MAX_PAGES} 页，请填写页码范围')
        pages.update(range(start, end + 1))
    if not pages:
        raise ValueError(RANGE_HINT)
    return sorted(pages)


def render_pdf_pages(data, document_name, page_spec='', dpi=150):
    if len(data) > MAX_FILE:
        raise ValueError('PDF 文件不能超过 50 MB')
    if dpi not in DPI_CHOICES:
        raise ValueError('清晰度仅支持 150 或 300 DPI')
    try:
        document = pymupdf.open(stream=data, filetype='pdf')
    except Exception as exc:
        raise ValueError('文件损坏或不是有效的 PDF') from exc
    try:
        if document.needs_pass:
            raise ValueError('请先解除 PDF 密码保护')
        total = document.page_count
        if not total:
            raise ValueError('PDF 没有可导入的页面')
        pages = parse_page_spec(page_spec, total)
        stem = Path(document_name or 'PDF').stem
        scale = dpi / 72
        for number in pages:
            page = document[number - 1]
            if page.rect.width * scale * page.rect.height * scale > MAX_PIXELS:
                raise ValueError(f'第 {number} 页尺寸过大，请降低清晰度')
            pixmap = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
            context = {
                'document': document_name or 'PDF', 'page': number, 'kind': 'pdf',
                'paragraph': None, 'inline_before': '', 'inline_after': '',
                'above': [], 'below': [], 'alt_text': [], 'candidates': [],
            }
            yield f'{stem}_p{number:03d}.png', context, pixmap.tobytes('png')
    finally:
        document.close()
