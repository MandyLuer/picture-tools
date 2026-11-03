"""Convert legacy DOC files in an isolated, bounded temporary job."""
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import zipfile

MAX_DOCUMENT = 50 * 1024 * 1024


def find_libreoffice():
    found = shutil.which('soffice') or shutil.which('libreoffice')
    if found:
        return found
    candidates = [Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')]
    for variable in ('PROGRAMFILES', 'PROGRAMFILES(X86)'):
        base = os.environ.get(variable)
        if base:
            candidates.append(Path(base) / 'LibreOffice/program/soffice.exe')
    return next((str(path) for path in candidates if path.is_file()), None)


def word_available():
    if sys.platform != 'win32':
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Word.Application\CLSID'):
            return True
    except OSError:
        return False


def convert_doc(data, timeout=90):
    if not data or len(data) > MAX_DOCUMENT:
        raise ValueError('DOC 文件为空或超过 50 MB')
    # Some exporters use a .doc suffix for OOXML; accept the real ZIP content.
    if data.startswith(b'PK\x03\x04') and zipfile.is_zipfile(io.BytesIO(data)):
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if 'word/document.xml' in archive.namelist():
                return data
        raise ValueError('DOC 文件内容不是有效的 Word 文档')
    if not (data.startswith(bytes.fromhex('d0cf11e0a1b11ae1')) or data.lstrip().startswith(bytes.fromhex('7b5c727466'))):
        raise ValueError('DOC 文件内容无效，请选择 Word 97–2003 文档或另存为 DOCX')
    provider = 'word' if word_available() else find_libreoffice()
    if not provider:
        raise ValueError('导入 .doc 需要本机安装 Microsoft Word 或 LibreOffice；也可先另存为 .docx 再导入')
    with tempfile.TemporaryDirectory(prefix='shitu-doc-') as folder:
        root = Path(folder)
        (root / 'input.doc').write_bytes(data)
        request = root / 'request.json'
        request.write_text(json.dumps({'provider': provider}), encoding='utf-8')
        command = ([sys.executable, '--document-worker', str(request)] if getattr(sys, 'frozen', False)
                   else [sys.executable, '-m', 'backend.document_worker', str(request)])
        try:
            completed = subprocess.run(command, timeout=timeout, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                       creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        except subprocess.TimeoutExpired as exc:
            raise ValueError('DOC 转换超时，已停止转换任务。请确认文档可正常打开且转换软件已激活') from exc
        result_path = root / 'result.json'
        if not result_path.is_file():
            raise ValueError('DOC 转换进程未完成，请检查 Word / LibreOffice 是否可以正常启动')
        result = json.loads(result_path.read_text(encoding='utf-8'))
        if completed.returncode or result.get('error'):
            raise ValueError(result.get('error') or 'DOC 转换失败')
        output = root / 'output.docx'
        if not output.is_file() or output.stat().st_size > MAX_DOCUMENT:
            raise ValueError('DOC 转换未生成有效结果或转换后的文件超过 50 MB')
        raw = output.read_bytes()
        if not zipfile.is_zipfile(io.BytesIO(raw)):
            raise ValueError('DOC 转换结果不是有效 DOCX 文件')
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            if 'word/document.xml' not in archive.namelist():
                raise ValueError('DOC 转换结果缺少 Word 正文，请尝试另存为 DOCX 后导入')
        return raw
