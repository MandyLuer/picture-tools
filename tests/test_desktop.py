"""Exercise pywebview's actual bridge injection without a native GUI."""
import json
import threading
from types import SimpleNamespace
from unittest.mock import Mock
import io
import zipfile
import pytest

pytest.importorskip("webview", reason="Desktop bridge tests require desktop dependencies")
from desktop import DesktopBridge


def test_bridge_injection_exposes_only_folder_actions(monkeypatch):
    from webview.util import inject_pywebview

    class PrivateWindow:
        def __dir__(self):
            raise AssertionError("Native window must never be traversed by bridge injection")

    bridge = DesktopBridge(object())
    bridge._window = PrivateWindow()
    expected = ['capture_certificate', 'choose_certificate', 'export_folder', 'export_zip', 'import_folder']
    assert [name for name in dir(bridge) if not name.startswith('_')] == expected
    # Production uses an explicit method allowlist, never js_api=bridge.
    scripts = []
    window = SimpleNamespace(
        _js_api=None,
        _functions={name: getattr(bridge, name) for name in expected},
        _expose_lock=threading.Lock(),
        events=SimpleNamespace(before_load=Mock(), _pywebviewready=threading.Event(), loaded=threading.Event()),
        run_js=scripts.append,
    )
    monkeypatch.setattr('webview.util.load_js_files', lambda *args: ('initial', '%(functions)s'))
    inject_pywebview('edgechromium', window)
    assert window.events.loaded.wait(2), 'Bridge injection hung before the loaded event'
    assert window.events._pywebviewready.is_set()
    assert sorted(f['func'] for f in json.loads(scripts[-1])) == expected


@pytest.mark.parametrize('action', ['import_folder', 'export_folder', 'export_zip'])
def test_cancelled_dialog_is_not_reported_as_success(action):
    bridge = DesktopBridge(object())
    bridge._window = SimpleNamespace(create_file_dialog=Mock(return_value=None))
    assert getattr(bridge, action)() == {'cancelled': True}


def test_zip_export_reports_success_after_writing_file(tmp_path):
    from backend.service import Workspace
    from PIL import Image
    work = Workspace()
    try:
        output = io.BytesIO()
        Image.new('RGB', (10, 10), 'red').save(output, 'PNG')
        work.add(output.getvalue(), 'image.png')
        bridge = DesktopBridge(SimpleNamespace(state=SimpleNamespace(workspace=work, lock=threading.RLock())))
        destination = tmp_path / 'pictures.zip'
        bridge._window = SimpleNamespace(create_file_dialog=Mock(return_value=[str(destination)]))
        assert bridge.export_zip() == {'count': 1}
        with zipfile.ZipFile(destination) as archive:
            assert archive.namelist() == ['image.png']
    finally:
        work.close()
