import shutil
from pathlib import Path
from tempfile import SpooledTemporaryFile

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PyPdfError

from .service import safe_name


def merge_to_folder(files, name, choose_folder):
    if len(files) < 2:
        raise ValueError('请至少选择两份 PDF 文件')
    if sum(file.size or 0 for file in files) > 250 * 1024 * 1024:
        raise ValueError('PDF 文件总大小不能超过 250 MB')
    name = name.strip()
    stem = safe_name(name[:-4] if name.lower().endswith('.pdf') else name)
    writer = PdfWriter()
    for index, file in enumerate(files, 1):
        try:
            reader = PdfReader(file.file)
            if reader.is_encrypted:
                raise ValueError('请先解除 PDF 密码保护')
            if not reader.pages:
                raise ValueError('文件没有可合并的页面')
            if reader.get_fields():
                reader.add_form_topname(f'文档{index}')
            writer.append(reader)
        except (PyPdfError, ValueError, OSError, KeyError, TypeError) as exc:
            reason = str(exc) if isinstance(exc, ValueError) else '文件损坏或不是有效的 PDF'
            raise ValueError(f'无法合并“{file.filename}”：{reason}') from exc

    with SpooledTemporaryFile(max_size=8 * 1024 * 1024) as merged:
        writer.write(merged)
        folder = choose_folder()
        if not folder:
            return {'cancelled': True}
        for index in range(10000):
            suffix = f'_{index}' if index else ''
            target = Path(folder) / f'{stem}{suffix}.pdf'
            try:
                output = target.open('xb')
            except FileExistsError:
                continue
            try:
                with output:
                    merged.seek(0)
                    shutil.copyfileobj(merged, output)
            except OSError:
                target.unlink(missing_ok=True)
                raise
            return {'name': target.name, 'pages': len(writer.pages), 'folder': Path(folder).name}
        raise ValueError('同名文件过多，请修改输出名称')
