import io
from pathlib import Path

from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app


def add_image(client):
    output = io.BytesIO()
    Image.new('RGB', (12, 8), 'green').save(output, 'PNG')
    raw = output.getvalue()
    assert client.post('/api/import', files=[('files', ('export.png', raw))]).status_code == 200
    return raw


def test_local_folder_export_uses_system_chooser_and_preserves_bytes(monkeypatch, tmp_path):
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: str(tmp_path))
    with TestClient(create_app(), client=('127.0.0.1', 50000)) as client:
        raw = add_image(client)
        result = client.post('/api/export/folder', json={'folder': 'ignored-client-path'})
        assert result.status_code == 200
        assert result.json()['count'] == 1
        assert (tmp_path / 'export.png').read_bytes() == raw
        assert len(client.get('/api/images').json()) == 1


def test_cancelled_system_chooser_writes_nothing(monkeypatch, tmp_path):
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: None)
    with TestClient(create_app(), client=('127.0.0.1', 50000)) as client:
        add_image(client)
        assert client.post('/api/export/folder').json() == {'cancelled': True}
        assert list(tmp_path.iterdir()) == []


def test_existing_file_is_not_overwritten(monkeypatch, tmp_path):
    (tmp_path / 'export.png').write_bytes(b'existing')
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: str(tmp_path))
    with TestClient(create_app(), client=('127.0.0.1', 50000)) as client:
        add_image(client)
        response = client.post('/api/export/folder')
        assert response.status_code == 400 and '同名' in response.json()['detail']
        assert (tmp_path / 'export.png').read_bytes() == b'existing'


def test_remote_clients_cannot_open_a_server_folder_chooser(monkeypatch):
    called = []
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: called.append(True))
    with TestClient(create_app(), client=('192.0.2.1', 50000)) as client:
        assert client.post('/api/export/folder').status_code == 403
    assert not called


def test_workspace_changed_during_chooser_does_not_export_stale_images(monkeypatch, tmp_path):
    app = create_app()
    def choose():
        with app.state.lock:
            app.state.workspace.clear()
        return str(tmp_path)
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', choose)
    with TestClient(app, client=('127.0.0.1', 50000)) as client:
        add_image(client)
        response = client.post('/api/export/folder')
        assert response.status_code == 400 and '工作区已变化' in response.json()['detail']
        assert list(tmp_path.iterdir()) == []
