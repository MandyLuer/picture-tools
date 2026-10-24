import io
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.service import Workspace


def png(color='red', size=(40, 20)):
    output = io.BytesIO()
    Image.new('RGB', size, color).save(output, 'PNG')
    return output.getvalue()


@pytest.fixture
def workspace():
    work = Workspace()
    yield work
    work.close()


def test_document_order_and_deduplication(workspace):
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w') as archive:
        archive.writestr('word/_rels/document.xml.rels', '<Relationships><Relationship Id="a" Target="media/a.png"/><Relationship Id="b" Target="media/b.png"/></Relationships>')
        archive.writestr('word/document.xml', '<root xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><blip r:embed="b"/><imagedata r:id="a"/><blip r:embed="b"/></root>')
        archive.writestr('word/media/a.png', png('red'))
        archive.writestr('word/media/b.png', png('blue'))
    assert workspace.extract(output.getvalue()) == 2
    items = list(workspace.images.values())
    assert [i['name'] for i in items] == ['001.png', '002.png']
    with Image.open(items[0]['path']) as image:
        assert image.getpixel((0, 0)) == (0, 0, 255)


def test_rename_crop_export(workspace, tmp_path):
    key = workspace.add(png(), '001.png')
    workspace.rename(key, '新名称')
    cropped = workspace.edit(key, 'crop', [2, 3, 20, 15])
    assert cropped == key
    assert workspace.get(cropped)['name'] == '新名称.png'
    assert workspace.listing()[0]['width'] == 18
    assert workspace.listing()[0]['height'] == 12
    other = workspace.add(png(), 'other.png')
    with pytest.raises(ValueError):
        workspace.rename(other, '新名称')
    with pytest.raises(ValueError):
        workspace.rename(key, '../outside')
    with pytest.raises(ValueError):
        workspace.edit(key, 'crop', [-1, 0, 20, 10])
    workspace.export_folder(tmp_path)
    with pytest.raises(ValueError):
        workspace.export_folder(tmp_path)
    with zipfile.ZipFile(io.BytesIO(workspace.export_zip())) as archive:
        assert archive.namelist() == ['新名称.png', 'other.png']


@pytest.mark.parametrize('operation', ['left', 'right', 'horizontal', 'vertical', 'compress'])
def test_edits_update_same_image(workspace, operation):
    image = Image.new('RGB', (40, 20), 'red')
    image.paste('blue', (20, 0, 40, 20))
    source = io.BytesIO()
    image.save(source, 'PNG')
    key = workspace.add(source.getvalue(), 'source.png')
    result = workspace.edit(key, operation)
    assert result == key
    assert len(workspace.images) == 1
    assert workspace.listing()[0]['revision'] == 1
    expected = {'left': Image.Transpose.ROTATE_90, 'right': Image.Transpose.ROTATE_270,
                'horizontal': Image.Transpose.FLIP_LEFT_RIGHT, 'vertical': Image.Transpose.FLIP_TOP_BOTTOM}
    transformed = image.transpose(expected[operation]) if operation != 'compress' else image
    with Image.open(workspace.get(result)['path']) as edited:
        assert edited.size == transformed.size
        assert edited.convert('RGB').tobytes() == transformed.tobytes()


def test_api_auth_import_and_rollback():
    app = create_app('secret')
    with TestClient(app) as client:
        assert client.get('/api/images').status_code == 403
        client.headers['x-app-token'] = 'secret'
        assert client.post('/api/import', files=[('files', ('a.png', png())), ('files', ('bad.docx', b'broken'))]).status_code == 400
        assert client.get('/api/images').json() == []
        result = client.post('/api/import', files={'files': ('a.png', png())})
        key = result.json()[0]['id']
        assert client.patch(f'/api/images/{key}', json={'name': 'renamed'}).status_code == 200
        assert client.get(f'/api/images/{key}/file').content == png()
        preview = client.get(f'/api/images/{key}/file?preview=true')
        assert preview.headers['content-type'] == 'image/png'
        with Image.open(io.BytesIO(preview.content)) as rendered:
            assert rendered.size == (40, 20)
        edit = client.post(f'/api/images/{key}/edit', json={'operation': 'crop', 'box': [0, 0, 10, 10]})
        assert edit.status_code == 200
        assert edit.json()['selected'] == key
        assert len(edit.json()['images']) == 1
        assert edit.json()['images'][0]['revision'] == 1
        assert client.get('/api/export').headers['content-type'] == 'application/zip'
        assert client.delete('/api/images').json() == []
        assert client.get(f'/api/images/{key}/file').status_code == 404
        assert client.delete('/api/images').json() == []
        assert client.post('/api/import', files={'files': ('again.png', png())}).status_code == 200


def test_delete_one_keeps_other_images_and_export(workspace):
    first = workspace.add(png(), 'first.png')
    second = workspace.add(png('blue'), 'second.png')
    path = workspace.get(first)['path']
    workspace.remove(first)
    assert not path.exists()
    assert [item['id'] for item in workspace.listing()] == [second]
    with zipfile.ZipFile(io.BytesIO(workspace.export_zip())) as archive:
        assert archive.namelist() == ['second.png']


def test_failed_edit_leaves_image_intact(workspace, monkeypatch):
    key = workspace.add(png(), 'source.png')
    def fail_replace(*args):
        raise PermissionError('locked')
    monkeypatch.setattr(Path, 'replace', fail_replace)
    with pytest.raises(PermissionError):
        workspace.edit(key, 'left')
    assert workspace.get(key)['path'].read_bytes() == png()
    assert workspace.get(key)['revision'] == 0
    assert len(list(workspace.root.iterdir())) == 1


def test_clear_locked_storage_does_not_leave_stale_active_images(workspace, monkeypatch):
    key = workspace.add(png(), 'old.png')
    old_temp = workspace.temp
    with monkeypatch.context() as patch:
        def locked():
            raise PermissionError('locked')
        patch.setattr(old_temp, 'cleanup', locked)
        workspace.clear()
        assert workspace.listing() == []
        with pytest.raises(KeyError):
            workspace.get(key)
        new_key = workspace.add(png(), 'new.png')
        assert workspace.get(new_key)['path'].parent == workspace.root
    workspace.clear()
    assert not Path(old_temp.name).exists()


def test_single_delete_api():
    with TestClient(create_app()) as client:
        items = client.post('/api/import', files=[('files', ('a.png', png())), ('files', ('b.png', png('blue')))]).json()
        assert len(client.delete(f"/api/images/{items[0]['id']}").json()) == 1
        assert client.get(f"/api/images/{items[0]['id']}/file?preview=true").status_code == 404
        assert client.delete(f"/api/images/{items[0]['id']}").status_code == 404
        assert client.get(f"/api/images/{items[1]['id']}/file").content == png('blue')
