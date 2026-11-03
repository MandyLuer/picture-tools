import io
import json
from pathlib import Path
import subprocess
from types import SimpleNamespace
import zipfile

import pytest
from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.document_conversion import convert_doc


def docx():
    image = io.BytesIO()
    Image.new('RGB', (30, 20), 'green').save(image, 'PNG')
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w') as archive:
        archive.writestr('word/_rels/document.xml.rels', '<Relationships><Relationship Id="a" Target="media/a.png"/></Relationships>')
        archive.writestr('word/document.xml', '<root xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>图 1 测试说明</w:t></w:r></w:p><w:p><blip r:embed="a"/></w:p></root>')
        archive.writestr('word/media/a.png', image.getvalue())
    return output.getvalue()


LEGACY = bytes.fromhex('d0cf11e0a1b11ae1') + b'fixture'


def test_missing_doc_converter_reports_actionable_error(monkeypatch):
    monkeypatch.setattr('backend.document_conversion.word_available', lambda: False)
    monkeypatch.setattr('backend.document_conversion.find_libreoffice', lambda: None)
    with pytest.raises(ValueError, match='安装 Microsoft Word 或 LibreOffice'):
        convert_doc(LEGACY)


def test_rtf_doc_uses_the_converter(monkeypatch):
    monkeypatch.setattr('backend.document_conversion.word_available', lambda: False)
    monkeypatch.setattr('backend.document_conversion.find_libreoffice', lambda: None)
    with pytest.raises(ValueError, match='安装 Microsoft Word 或 LibreOffice'):
        convert_doc(bytes.fromhex('7b5c72746631') + b' test}')


def test_doc_conversion_preserves_original_and_validates_worker_output(monkeypatch):
    monkeypatch.setattr('backend.document_conversion.word_available', lambda: True)
    def complete(command, **kwargs):
        root = Path(command[-1]).parent
        assert (root / 'input.doc').read_bytes() == LEGACY
        assert json.loads((root / 'request.json').read_text(encoding='utf-8'))['provider'] == 'word'
        (root / 'output.docx').write_bytes(docx())
        (root / 'result.json').write_text('{"converted":true}', encoding='utf-8')
        return SimpleNamespace(returncode=0)
    monkeypatch.setattr(subprocess, 'run', complete)
    assert convert_doc(LEGACY) == docx()
    assert LEGACY.endswith(b'fixture')


def test_doc_conversion_timeout_is_clear(monkeypatch):
    monkeypatch.setattr('backend.document_conversion.word_available', lambda: True)
    def timeout(command, **kwargs):
        raise subprocess.TimeoutExpired(command, 1)
    monkeypatch.setattr(subprocess, 'run', timeout)
    with pytest.raises(ValueError, match='DOC 转换超时'):
        convert_doc(LEGACY, timeout=1)


@pytest.mark.parametrize('data', [b'', b'broken', b'x' * (50 * 1024 * 1024 + 1)], ids=['empty','invalid','oversize'])
def test_invalid_doc_is_rejected_without_launching_office(data, monkeypatch):
    called = []
    monkeypatch.setattr(subprocess, 'run', lambda *a, **kw: called.append(True))
    with pytest.raises(ValueError):
        convert_doc(data)
    assert not called


def test_docx_with_doc_suffix_needs_no_converter(monkeypatch):
    monkeypatch.setattr('backend.document_conversion.word_available', lambda: pytest.fail('No conversion required'))
    assert convert_doc(docx()) == docx()


def test_binary_doc_with_embedded_theme_zip_is_not_mistaken_for_docx(monkeypatch):
    theme = io.BytesIO()
    with zipfile.ZipFile(theme, 'w') as archive:
        archive.writestr('theme/theme1.xml', '<theme/>')
    data = LEGACY + theme.getvalue()
    assert zipfile.is_zipfile(io.BytesIO(data))
    monkeypatch.setattr('backend.document_conversion.word_available', lambda: False)
    monkeypatch.setattr('backend.document_conversion.find_libreoffice', lambda: None)
    with pytest.raises(ValueError, match='安装 Microsoft Word 或 LibreOffice'):
        convert_doc(data)


def test_doc_and_docx_import_use_same_image_and_context_pipeline(monkeypatch):
    monkeypatch.setattr('backend.document_conversion.convert_doc', lambda data: docx())
    with TestClient(create_app()) as client:
        response = client.post('/api/import', files=[('files', ('legacy.DOC', LEGACY)), ('files', ('modern.docx', docx()))])
        assert response.status_code == 200
        images = response.json()
        assert len(images) == 2
        assert [item['contexts'][0]['document'] for item in images] == ['legacy.DOC','modern.docx']
        assert images[0]['contexts'][0]['above'] == ['图 1 测试说明']


def test_failed_doc_rolls_back_batch_and_retains_existing_images(monkeypatch):
    def failure(data):
        raise ValueError('导入 .doc 需要本机安装 Microsoft Word 或 LibreOffice')
    monkeypatch.setattr('backend.document_conversion.convert_doc', failure)
    with TestClient(create_app()) as client:
        existing = client.post('/api/import', files=[('files', ('existing.docx', docx()))]).json()
        response = client.post('/api/import', files=[('files', ('new.docx', docx())), ('files', ('legacy.doc', LEGACY))])
        assert response.status_code == 400
        assert 'Microsoft Word 或 LibreOffice' in response.json()['detail']
        assert client.get('/api/images').json() == existing
