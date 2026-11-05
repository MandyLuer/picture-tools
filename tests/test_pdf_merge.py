import io

from fastapi.testclient import TestClient
from pypdf import PdfReader, PdfWriter

from backend.app import create_app


def test_pdf_merge_keeps_order_and_existing_files(tmp_path, monkeypatch):
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: tmp_path)
    files = []
    for name, width in [('second.pdf', 300), ('first.pdf', 200)]:
        pdf = PdfWriter()
        pdf.add_blank_page(width, 100)
        buffer = io.BytesIO()
        pdf.write(buffer)
        files.append(('files', (name, buffer.getvalue(), 'application/pdf')))

    (tmp_path / 'merged.pdf').write_bytes(files[0][1][1])
    with TestClient(create_app(), client=('127.0.0.1', 12345)) as client:
        response = client.post('/api/pdf/merge', files=files, data={'name': 'merged.PDF'})
        assert response.status_code == 200
        assert response.json()['name'] == 'merged_1.pdf'
        assert [int(page.mediabox.width) for page in PdfReader(tmp_path / 'merged_1.pdf').pages] == [300, 200]
        assert (tmp_path / 'merged.pdf').read_bytes() == files[0][1][1]

        invalid = client.post('/api/pdf/merge', files=files[:1] + [('files', ('broken.pdf', b'invalid'))])
        assert invalid.status_code == 400
        assert 'broken.pdf' in invalid.json()['detail']
        assert len(list(tmp_path.iterdir())) == 2

        source_folder = tmp_path / 'source'
        source_folder.mkdir()
        paths = []
        for _, (filename, content, _) in files:
            path = source_folder / filename
            path.write_bytes(content)
            paths.append(path)
        monkeypatch.setattr('backend.folder_picker.choose_pdf_files', lambda: list(map(str, paths)))

        def choose_destination(initial):
            assert initial == str(source_folder)
            return initial

        monkeypatch.setattr('backend.folder_picker.choose_export_folder', choose_destination)
        chosen = client.post('/api/pdf/choose').json()['files']
        response = client.post('/api/pdf/merge', data={'source_ids': [file['id'] for file in reversed(chosen)],
                                                      'name': 'combined'})
        assert response.status_code == 200
        assert [int(page.mediabox.width) for page in PdfReader(source_folder / 'combined.pdf').pages] == [200, 300]
