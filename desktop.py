"""Local desktop host. Never exposes arbitrary filesystem paths over HTTP."""
import socket
import threading
import time
import secrets
import re
import sys
import urllib.request
import os
import logging
import faulthandler
from pathlib import Path

import uvicorn
import webview

from backend.app import create_app


class DesktopBridge:
    def __init__(self, app):
        # pywebview recursively inspects public attributes during bridge injection.
        # Keep server/window internals private and expose only explicit methods.
        self._app = app
        self._window = None
        self._certificate = None
        self._capture_lock = threading.Lock()

    def choose_certificate(self):
        if sys.platform != 'win32':
            return {'error': '原生证书截图仅支持 Windows 桌面客户端'}
        result = self._window.create_file_dialog(webview.FileDialog.OPEN, file_types=('Certificate (*.crt;*.cer;*.pem)',))
        if not result:
            return {'cancelled': True}
        try:
            from backend.certificates import read_certificate, MAX_CERTIFICATE
            path = Path(result[0])
            if path.stat().st_size > MAX_CERTIFICATE:
                raise ValueError('证书文件不能超过 2 MB')
            der, metadata = read_certificate(path.read_bytes(), path.name)
            self._certificate = (der, metadata)
            return {'certificate': metadata}
        except Exception as exc:
            return {'error': str(exc)}

    def capture_certificate(self, kinds):
        if not self._capture_lock.acquire(blocking=False):
            return {'error': '已有证书截图任务正在运行'}
        try:
            if self._certificate is None:
                raise ValueError('请先选择一张证书')
            from backend.certificates import screenshot_certificate
            der, metadata = self._certificate
            with self._app.state.lock:
                generation = self._app.state.workspace.root
            pictures = screenshot_certificate(der, metadata, kinds)
            with self._app.state.lock:
                workspace = self._app.state.workspace
                if workspace.root != generation:
                    raise ValueError('工作区已变更，本次截图未导入，请重试')
                added = []
                try:
                    for raw, name in pictures:
                        added.append(workspace.add(raw, name))
                    return {'count': len(added), 'images': workspace.listing(), 'selected': added[0]}
                except Exception:
                    for key in added:
                        workspace.remove(key)
                    raise
        except Exception as exc:
            return {'error': str(exc)}
        finally:
            self._capture_lock.release()

    def import_folder(self):
        result = self._window.create_file_dialog(webview.FileDialog.FOLDER)
        if not result:
            return {"cancelled": True}
        workspace = self._app.state.workspace
        with self._app.state.lock:
            before = set(workspace.images)
            try:
                for path in sorted(Path(result[0]).iterdir(), key=lambda p: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', p.name)]):
                    if path.is_file():
                        from PIL import Image
                        try:
                            with Image.open(path) as image:
                                image.verify()
                        except (OSError, ValueError):
                            continue
                        if path.stat().st_size > 50 * 1024 * 1024:
                            raise ValueError("单张图片不能超过 50 MB")
                        workspace.add(path.read_bytes(), path.name)
            except Exception as exc:
                for key in set(workspace.images) - before:
                    workspace.images.pop(key)["path"].unlink(missing_ok=True)
                return {"error": str(exc)}
        count = len(workspace.images) - len(before)
        return {"count": count} if count else {"error": "所选文件夹中没有找到支持的图片"}

    def export_folder(self):
        result = self._window.create_file_dialog(webview.FileDialog.FOLDER)
        if not result:
            return {"cancelled": True}
        try:
            with self._app.state.lock:
                workspace = self._app.state.workspace
                workspace.export_folder(result[0])
                return {"count": len(workspace.images)}
        except Exception as exc:
            return {"error": str(exc)}

    def export_zip(self):
        result = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename='拾图图片.zip', file_types=('ZIP (*.zip)',))
        if not result:
            return {"cancelled": True}
        try:
            with self._app.state.lock:
                workspace = self._app.state.workspace
                target = Path(result[0])
                if target.suffix.lower() != '.zip':
                    target = target.with_name(target.name + '.zip')
                target.write_bytes(workspace.export_zip())
                return {"count": len(workspace.images)}
        except Exception as exc:
            return {"error": str(exc)}


def main():
    gui_test = "--gui-smoke-test" in sys.argv
    log_root = (Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Shitu" / "logs"
                if getattr(sys, "frozen", False) else Path(__file__).resolve().parent / "logs")
    log_root.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=log_root / "desktop.log", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(name)s: %(message)s", encoding="utf-8")
    logger = logging.getLogger("shitu.desktop")
    logger.info("Starting desktop; gui_test=%s", gui_test)
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shitu.Desktop.2")
    webview.settings['ALLOW_DOWNLOADS'] = True
    token = secrets.token_urlsafe(32)
    app = create_app(token)
    context_test = gui_test and "--test-document" in sys.argv
    if context_test:
        test_document = Path(sys.argv[sys.argv.index("--test-document") + 1])
        app.state.workspace.extract(test_document.read_bytes(), test_document.name)
    certificate_test = gui_test and '--test-certificate' in sys.argv
    if certificate_test:
        from backend.certificates import read_certificate, screenshot_certificate
        cert_path = Path(sys.argv[sys.argv.index('--test-certificate') + 1])
        der, metadata = read_certificate(cert_path.read_bytes(), cert_path.name)
        for raw, name in screenshot_certificate(der, metadata, ['general','issuer','public_key','path']):
            app.state.workspace.add(raw, name)
    list_test = gui_test and "--test-image-list" in sys.argv
    if list_test:
        import io
        from PIL import Image
        sample = io.BytesIO()
        Image.new('RGB', (80, 40), '#126c62').save(sample, 'PNG')
        for number in range(80):
            app.state.workspace.add(sample.getvalue(), f'list-test-{number:03d}.png')
    sock = socket.socket()
    sock.bind(("127.0.0.1", 0))
    port = sock.getsockname()[1]
    server = uvicorn.Server(uvicorn.Config(app, log_config=None, access_log=False))
    worker = threading.Thread(target=server.run, kwargs={"sockets": [sock]}, daemon=True)
    worker.start()
    deadline = time.monotonic() + 15
    while not server.started:
        if not worker.is_alive() or time.monotonic() > deadline:
            raise RuntimeError("本地服务启动失败")
        time.sleep(.05)
    logger.info("Local service ready")
    if "--smoke-test" in sys.argv:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=10) as response:
                assert b'<div id="app">' in response.read()
            request = urllib.request.Request(f"http://127.0.0.1:{port}/api/images", headers={"x-app-token": token})
            with urllib.request.urlopen(request, timeout=10) as response:
                assert response.read() == b'[]'
        finally:
            server.should_exit = True
            worker.join(timeout=10)
            sock.close()
        return
    bridge = DesktopBridge(app)
    compact_test = gui_test and '--compact-window' in sys.argv
    window = webview.create_window("拾图 Shitu · Word 图片工作台", f"http://127.0.0.1:{port}/?token={token}", width=800 if compact_test else 1440, height=650 if compact_test else 960, min_size=(800, 650))
    bridge._window = window
    window.expose(bridge.import_folder, bridge.export_folder, bridge.export_zip,
                  bridge.choose_certificate, bridge.capture_certificate)
    loaded = threading.Event()
    test_success = threading.Event()

    def on_loaded():
        logger.info("WebView loaded and bridge ready")
        loaded.set()
        if gui_test:
            try:
                deadline = time.monotonic() + 10
                result = None
                while time.monotonic() < deadline:
                    result = window.evaluate_js("""({
                    title: document.title,
                    uiReady: !!document.querySelector('.workspace-heading'),
                    bridge: Object.keys(window.pywebview.api).sort(),
                    desktopReady: document.querySelector('.badge')?.textContent.includes('桌面版'),
                    menuCount: document.querySelectorAll('.feature-menu [role=tab]').length,
                    sidebarImports: document.querySelectorAll('aside input[type=file]').length,
                    certificateEntry: !!document.querySelector('[data-menu=certificate]'),
                    contextReady: !!document.querySelector('.document-context .name-candidates button'),
                    imageCount: document.images.length,
                    imagesReady: Array.from(document.images).filter(i => i.loading !== 'lazy').every(i => i.complete && i.naturalWidth > 0),
                    listCount: document.querySelectorAll('.image-row').length,
                    listHeight: document.querySelector('.image-list')?.clientHeight,
                    listScrollHeight: document.querySelector('.image-list')?.scrollHeight,
                    listOverflow: getComputedStyle(document.querySelector('.image-list')).overflowY,
                    previewWidth: document.querySelector('.stage')?.clientWidth,
                    previewHeight: document.querySelector('.stage')?.clientHeight,
                    pageHeight: document.documentElement.scrollHeight,
                    viewportHeight: innerHeight,
                    pageWidth: document.documentElement.scrollWidth,
                    viewportWidth: innerWidth
                    })""")
                    if (result and result['uiReady'] and result['desktopReady'] and result['imagesReady']
                            and (not context_test or result['contextReady'])
                            and (not list_test or result['listCount'] >= 80)):
                        break
                    time.sleep(.1)
                if context_test:
                    if not result or not result['contextReady']:
                        raise RuntimeError("Document reference descriptions did not render")
                    logger.info("Document context panel and name candidates rendered")
                logger.info("GUI smoke result: %s", result)
                if list_test and not (result['listCount'] >= 80 and result['listHeight'] > 0
                                      and result['listScrollHeight'] > result['listHeight']
                                      and result['listOverflow'] == 'auto'
                                      and result['pageHeight'] <= result['viewportHeight'] + 1):
                    raise RuntimeError('Image list must scroll independently without growing the page')
                if certificate_test and not (result['certificateEntry'] and result['listCount'] == 4):
                    raise RuntimeError('Certificate screenshots did not reach the desktop image list')
                if not (result and result["uiReady"] and result["desktopReady"] and result["imagesReady"]
                        and result["previewWidth"] > 0 and result["previewHeight"] > 0
                        and result['menuCount'] == 3 and result['sidebarImports'] == 0
                        and result['pageWidth'] <= result['viewportWidth'] + 1
                        and result["bridge"] == ["capture_certificate", "choose_certificate", "export_folder", "export_zip", "import_folder"]):
                    raise RuntimeError("Desktop UI/bridge readiness check failed")
                test_success.set()
            except Exception:
                logger.exception("GUI smoke failed")
            finally:
                window.destroy()

    window.events.loaded += on_loaded

    def closing():
        if not gui_test and app.state.workspace.images and not window.create_confirmation_dialog("退出工作台", "退出将清除临时工作区。请确认已导出需要的图片，是否退出？"):
            return False
    window.events.closing += closing
    def monitor_startup():
        if not loaded.wait(30):
            logger.error("WebView initialization timed out")
            with (log_root / "startup-threads.log").open("w") as dump:
                faulthandler.dump_traceback(file=dump, all_threads=True)
            if gui_test:
                os._exit(2)
    threading.Thread(target=monitor_startup, daemon=True).start()
    try:
        resource_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        logger.info("Starting WebView event loop")
        webview.start(gui="edgechromium" if sys.platform == "win32" else None,
                      icon=str(resource_root / "assets/branding/shitu.ico"))
    finally:
        server.should_exit = True
        worker.join(timeout=10)
        sock.close()
        logger.info("Desktop stopped")
    if gui_test and not test_success.is_set():
        raise RuntimeError("GUI smoke test did not complete successfully")


if __name__ == "__main__":
    if '--folder-picker' in sys.argv:
        from backend.folder_picker import main as folder_picker
        sys.exit(folder_picker(sys.argv[sys.argv.index('--folder-picker') + 1]))
    if '--folder-picker' in sys.argv:
        from backend.folder_picker import main as folder_picker
        sys.exit(folder_picker(sys.argv[sys.argv.index('--folder-picker') + 1]))
    if '--certificate-worker' in sys.argv:
        from backend.certificate_worker import main as certificate_worker
        sys.exit(certificate_worker(sys.argv[sys.argv.index('--certificate-worker') + 1]))
    try:
        main()
    except Exception:
        logging.getLogger("shitu.desktop").exception("Desktop startup failed")
        if "--gui-smoke-test" in sys.argv or "--smoke-test" in sys.argv:
            sys.exit(1)
        raise
