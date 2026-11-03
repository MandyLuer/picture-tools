import io
from concurrent.futures import ThreadPoolExecutor
from threading import Event
from unittest.mock import patch

from fastapi.testclient import TestClient
from PIL import Image

from backend.app import create_app
from backend.service import Workspace


def image_bytes(fmt='PNG'):
    output = io.BytesIO()
    Image.new('RGB', (80, 40), 'green').save(output, fmt)
    return output.getvalue()


def test_listing_and_rename_do_not_reopen_image_files():
    workspace = Workspace()
    try:
        key = workspace.add(image_bytes(), 'original.png')
        with patch('backend.service.Image.open', side_effect=AssertionError('unnecessary disk image read')):
            workspace.rename(key, 'renamed')
            item = workspace.listing()[0]
        assert (item['name'], item['width'], item['height'], item['size']) == ('renamed.png', 80, 40, len(image_bytes()))
    finally:
        workspace.close()


def test_previews_are_cached_across_rename_and_invalidated_after_edit():
    app = create_app()
    with TestClient(app) as client:
        key = client.post('/api/import', files=[('files', ('image.jpg', image_bytes('JPEG')))]).json()[0]['id']
        url = f'/api/images/{key}/file?preview=true&revision=0'
        with patch.object(app.state.workspace, 'render_preview', wraps=app.state.workspace.render_preview) as render:
            first = client.get(url)
            client.patch(f'/api/images/{key}', json={'name':'renamed'})
            second = client.get(url)
            assert render.call_count == 1
        assert first.content == second.content
        assert 'immutable' in first.headers['cache-control']
        assert client.get(url, headers={'if-none-match':first.headers['etag']}).status_code == 304
        client.post(f'/api/images/{key}/edit', json={'operation':'right'})
        updated = client.get(f'/api/images/{key}/file?preview=true&revision=1', headers={'if-none-match':first.headers['etag']})
        assert updated.status_code == 200 and updated.headers['etag'] != first.headers['etag']
        with Image.open(io.BytesIO(updated.content)) as image:
            assert image.size == (40, 80)
        assert client.get(url).headers['cache-control'] == 'no-store'


def test_compact_rename_returns_only_the_current_name_without_listing_the_workspace():
    app = create_app()
    with TestClient(app) as client:
        key = client.post('/api/import', files=[('files', ('image.png', image_bytes()))]).json()[0]['id']
        with patch.object(app.state.workspace, 'listing', side_effect=AssertionError('full list not needed')):
            result = client.patch(f'/api/images/{key}?compact=true', json={'name':'updated'})
        assert result.json() == {'id':key, 'name':'updated.png', 'revision':0}
        assert len(client.patch(f'/api/images/{key}', json={'name':'compatible'}).json()) == 1


def test_preview_cache_is_bounded_and_does_not_keep_removed_or_stale_revisions(monkeypatch):
    monkeypatch.setattr('backend.service.MAX_PREVIEW_CACHE', 10)
    workspace = Workspace()
    try:
        first = workspace.add(image_bytes(), 'a.png')
        second = workspace.add(image_bytes(), 'b.png')
        workspace.cache_preview(first, 0, False, b'123456')
        workspace.cache_preview(second, 0, False, b'abcdef')
        assert workspace.cached_preview(first, 0, False) is None
        assert workspace._preview_bytes == 6
        workspace.cache_preview(second, 99, False, b'stale')
        assert workspace.cached_preview(second, 99, False) is None
        workspace.remove(second)
        assert workspace._preview_bytes == 0
        workspace.cache_preview(first, 0, True, b'thumb')
        workspace.clear()
        assert workspace._preview_bytes == 0 and not workspace._preview_cache
    finally:
        workspace.close()


def test_cold_preview_render_does_not_block_rename_or_cache_an_old_edit():
    app = create_app()
    entered, finish = Event(), Event()
    original = app.state.workspace.render_preview
    def slow_render(*args):
        entered.set()
        assert finish.wait(5)
        return original(*args)
    with TestClient(app) as client, ThreadPoolExecutor(max_workers=2) as pool:
        key = client.post('/api/import', files=[('files', ('image.jpg', image_bytes('JPEG')))]).json()[0]['id']
        with patch.object(app.state.workspace, 'render_preview', side_effect=slow_render):
            pending = pool.submit(client.get, f'/api/images/{key}/file?preview=true&revision=0')
            try:
                assert entered.wait(2)
                renamed = pool.submit(client.patch, f'/api/images/{key}', json={'name':'renamed'}).result(timeout=2)
                assert renamed.status_code == 200
                assert client.post(f'/api/images/{key}/edit', json={'operation':'right'}).status_code == 200
            finally:
                finish.set()
            assert pending.result(timeout=2).status_code == 200
        assert app.state.workspace.cached_preview(key, 0, False) is None
