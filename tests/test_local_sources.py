import io

import pytest
from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.local_sources import import_folder, rename_source, export_and_remove
from backend.service import Workspace


@pytest.fixture
def local(tmp_path):
    folder = tmp_path / 'originals'
    folder.mkdir()
    data = io.BytesIO()
    Image.new('RGB', (10, 8), 'green').save(data, 'PNG')
    source = folder / 'original.png'
    source.write_bytes(data.getvalue())
    work = Workspace()
    import_folder(work, folder)
    yield work, source, next(iter(work.images)), tmp_path / 'exported'
    work.close()


def test_rename_preserves_bytes_and_supports_repeated_rename(local):
    work, source, key, _ = local
    raw = source.read_bytes()
    rename_source(work, key, 'new')
    assert not source.exists()
    assert source.with_name('new.png').read_bytes() == raw
    rename_source(work, key, 'again')
    assert len(list(source.parent.iterdir())) == 1
    assert source.with_name('again.png').read_bytes() == raw
    assert work.listing()[0]['has_source']
    assert 'source' not in work.listing()[0]


def test_conflict_does_not_overwrite_or_change_workspace(local):
    work, source, key, _ = local
    target = source.with_name('occupied.png')
    target.write_bytes(b'existing')
    with pytest.raises(ValueError, match='已存在'):
        rename_source(work, key, 'occupied')
    assert source.exists()
    assert target.read_bytes() == b'existing'
    assert work.get(key)['name'] == 'original.png'


def test_external_changes_and_edits_prevent_source_rename(local):
    work, source, key, _ = local
    work.edit(key, 'left')
    with pytest.raises(ValueError, match='已有编辑'):
        rename_source(work, key, 'edited')
    source.write_bytes(b'changed')
    with pytest.raises(ValueError, match='已变化'):
        rename_source(work, key, 'changed')
    assert source.read_bytes() == b'changed'


def test_export_verifies_edited_result_then_deletes_only_registered_source(local):
    work, source, key, folder = local
    unrelated = source.with_name('unrelated.txt')
    unrelated.write_text('keep')
    work.edit(key, 'left')
    result = export_and_remove(work, folder)
    assert result['removed'] == 1
    assert result['errors'] == []
    assert not source.exists()
    assert unrelated.exists()
    assert (folder / 'original.png').read_bytes() == work.get(key)['path'].read_bytes()
    assert not work.get(key).get('source')


def test_export_conflict_keeps_source(local):
    work, source, key, folder = local
    folder.mkdir()
    (folder / 'original.png').write_bytes(b'keep')
    with pytest.raises(ValueError):
        export_and_remove(work, folder)
    assert source.exists()
    assert (folder / 'original.png').read_bytes() == b'keep'


def test_failed_export_keeps_source(local, monkeypatch):
    work, source, key, folder = local
    monkeypatch.setattr(type(folder), 'open', lambda *args, **kwargs: (_ for _ in ()).throw(OSError('disk full')))
    with pytest.raises(OSError):
        export_and_remove(work, folder)
    assert source.exists()


def test_http_import_cancel_and_source_rename(local, monkeypatch):
    _, source, _, _ = local
    monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: str(source.parent))
    with TestClient(create_app(), client=('127.0.0.1', 12345)) as client:
        result = client.post('/api/import/source-folder').json()
        key = result['images'][0]['id']
        assert result['count'] == 1
        assert client.post(f'/api/images/{key}/source-name', json={'name':'api-name'}).status_code == 200
        assert source.with_name('api-name.png').exists()
        monkeypatch.setattr('backend.folder_picker.choose_export_folder', lambda: None)
        assert client.post('/api/import/source-folder').json() == {'cancelled':True}
