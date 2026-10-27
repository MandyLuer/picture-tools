"""Product screenshot worker: operates only its own read-only CryptUI dialog.

No certificate-store writes, input-file changes or network certificate retrieval.
The parent process owns the timeout and terminates this worker on a hang.
"""
import ctypes
from ctypes import wintypes as wt
import json
import os
from pathlib import Path
import sys
import threading
import time


class ViewCertificate(ctypes.Structure):
    _fields_ = [
        ('dwSize', wt.DWORD), ('hwndParent', wt.HWND), ('dwFlags', wt.DWORD),
        ('szTitle', wt.LPCWSTR), ('pCertContext', ctypes.c_void_p),
        ('rgszPurposes', ctypes.c_void_p), ('cPurposes', wt.DWORD),
        ('provider', ctypes.c_void_p), ('trusted', wt.BOOL),
        ('idxSigner', wt.DWORD), ('idxCert', wt.DWORD), ('counterSigner', wt.BOOL),
        ('idxCounterSigner', wt.DWORD), ('cStores', wt.DWORD),
        ('rghStores', ctypes.c_void_p), ('cPropSheetPages', wt.DWORD),
        ('rgPropSheetPages', ctypes.c_void_p), ('nStartPage', wt.DWORD),
    ]


def open_viewer(der, title):
    crypt32 = ctypes.WinDLL('crypt32', use_last_error=True)
    cryptui = ctypes.WinDLL('cryptui', use_last_error=True)
    crypt32.CertCreateCertificateContext.argtypes = [wt.DWORD, ctypes.c_void_p, wt.DWORD]
    crypt32.CertCreateCertificateContext.restype = ctypes.c_void_p
    crypt32.CertFreeCertificateContext.argtypes = [ctypes.c_void_p]
    cryptui.CryptUIDlgViewCertificateW.argtypes = [ctypes.POINTER(ViewCertificate), ctypes.POINTER(wt.BOOL)]
    cryptui.CryptUIDlgViewCertificateW.restype = wt.BOOL
    buffer = ctypes.create_string_buffer(der)
    context = crypt32.CertCreateCertificateContext(0x10001, buffer, len(der))
    if not context:
        raise ctypes.WinError(ctypes.get_last_error())
    info = ViewCertificate()
    info.dwSize = ctypes.sizeof(info)
    # Disable editing/install and use cached chain data only (cryptuiapi.h).
    info.dwFlags = 0x4 | 0x10 | 0x10000 | 0x20000 | 0x40000
    info.szTitle, info.pCertContext = title, context
    changed = wt.BOOL()
    try:
        if not cryptui.CryptUIDlgViewCertificateW(ctypes.byref(info), ctypes.byref(changed)):
            error = ctypes.get_last_error()
            if error and error != 1223:  # Closing our read-only viewer is normal cancellation.
                raise ctypes.WinError(error)
    finally:
        crypt32.CertFreeCertificateContext(context)


def capture_window(dialog, path):
    """Print this GDI dialog, without including other desktop windows."""
    import win32gui
    import win32ui
    from PIL import Image, ImageStat

    rect = dialog.rectangle()
    width, height = rect.width(), rect.height()
    if width < 250 or height < 300:
        raise ValueError('证书窗口大小异常，无法截图')
    window_dc = win32gui.GetWindowDC(dialog.handle)
    source_dc = win32ui.CreateDCFromHandle(window_dc)
    memory_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source_dc, width, height)
    previous = memory_dc.SelectObject(bitmap)
    try:
        print_window = ctypes.windll.user32.PrintWindow
        print_window.argtypes = [wt.HWND, wt.HDC, wt.UINT]
        print_window.restype = wt.BOOL
        if not print_window(dialog.handle, memory_dc.GetSafeHdc(), 2):
            raise ValueError('Windows 无法捕获证书窗口，请确认当前桌面未锁定')
        image = Image.frombuffer('RGB', (width, height), bitmap.GetBitmapBits(True), 'raw', 'BGRX', 0, 1)
        # GetWindowRect includes invisible resize margins on Windows 10/11.
        frame = wt.RECT()
        dwm = ctypes.WinDLL('dwmapi')
        dwm.DwmGetWindowAttribute.argtypes = [wt.HWND, wt.DWORD, ctypes.c_void_p, wt.DWORD]
        if dwm.DwmGetWindowAttribute(dialog.handle, 9, ctypes.byref(frame), ctypes.sizeof(frame)) == 0:
            left, top = max(0, frame.left - rect.left), max(0, frame.top - rect.top)
            right, bottom = min(width, frame.right - rect.left), min(height, frame.bottom - rect.top)
            if right > left and bottom > top:
                image = image.crop((left, top, right, bottom))
        if max(ImageStat.Stat(image).stddev) < 4:
            raise ValueError('截图为空白，未导入结果，请在可交互桌面重试')
        image.save(path, 'PNG')
    finally:
        memory_dc.SelectObject(previous)
        win32gui.DeleteObject(bitmap.GetHandle())
        memory_dc.DeleteDC()
        source_dc.DeleteDC()
        win32gui.ReleaseDC(dialog.handle, window_dc)


def capture_job(root, kinds):
    from .certificates import find_field_index, ScreenshotFailure
    from pywinauto import Application
    from pywinauto.timings import wait_until

    title = '证书'
    viewer_error = []

    def show():
        try:
            open_viewer((root / 'certificate.der').read_bytes(), title)
        except Exception as exc:
            viewer_error.append(str(exc))

    thread = threading.Thread(target=show, daemon=True)
    thread.start()
    app = Application(backend='win32').connect(process=os.getpid(), timeout=15)
    spec = app.window(title=title, class_name='#32770')
    spec.wait('exists visible ready', timeout=20)
    dialog = spec.wrapper_object()
    active_kind = None
    try:
        dialog.restore()
        dialog.set_focus()
        tabs = spec.child_window(class_name='SysTabControl32').wrapper_object()
        if tabs.tab_count() < 3:
            raise ValueError('证书查看器未提供常规、详细信息和证书路径三个页签')
        records = []
        for number, kind in enumerate(kinds, 1):
            active_kind = kind
            (root / 'progress.json').write_text(json.dumps({'kind': kind}), encoding='utf-8')
            page = 0 if kind == 'general' else 2 if kind == 'path' else 1
            tabs.select(page)
            wait_until(8, .1, lambda: tabs.get_selected_tab() == page)
            selected_field = ''
            if page == 1:
                details = spec.child_window(class_name='SysListView32', visible_only=True).wrapper_object()
                wait_until(8, .1, lambda: details.item_count() > 0)
                # Use the available value-column width, so common names are readable.
                details.send_message(0x101E, 1, max(180, details.client_rect().width() - details.column_widths()[0] - 28))
                index = find_field_index([details.get_item(i, 0).text() for i in range(details.item_count())], kind)
                details.get_item(index).ensure_visible()
                details.select(index)
                details.set_focus()
                wait_until(5, .1, lambda: details.get_item(index).is_selected())
                selected_field = details.get_item(index, 0).text()
            elif kind == 'path':
                tree = spec.child_window(class_name='SysTreeView32', visible_only=True).wrapper_object()
                wait_until(10, .1, lambda: bool(tree.roots()))
            # Allow native property-sheet text and highlights to finish painting.
            dialog.send_message(0x000F)
            time.sleep(.3)
            if tabs.get_selected_tab() != page:
                raise ValueError('截图期间证书页签被切换，请避免操作证书窗口并重试')
            if page == 1 and not details.get_item(index).is_selected():
                raise ValueError('截图期间证书字段被切换，请重试')
            filename = f'{number:02d}.png'
            capture_window(dialog, root / filename)
            records.append({'kind': kind, 'file': filename, 'selected_field': selected_field,
                            'tab': tabs.get_tab_text(page), 'capture_method': 'PrintWindow'})
        return {'screenshots': records, 'chain_source': 'Windows local stores/cache; no online retrieval'}
    except Exception as exc:
        raise ScreenshotFailure(str(exc), active_kind) from exc
    finally:
        if spec.exists():
            dialog.close()
        thread.join(timeout=3)
        if viewer_error and sys.exc_info()[0] is None:
            raise ValueError(viewer_error[0])


def main(request_file):
    root = Path(request_file).resolve().parent
    try:
        if sys.platform != 'win32':
            raise ValueError('原生证书截图仅支持 Windows')
        ctypes.windll.user32.SetProcessDPIAware()
        from .certificates import validate_kinds
        request = json.loads(Path(request_file).read_text(encoding='utf-8'))
        result = capture_job(root, validate_kinds(request['kinds']))
        status = 0
    except Exception as exc:
        result, status = {'error': f'证书截图失败：{exc}',
                          'failed_kind': getattr(exc, 'failed_kind', ''),
                          'failed_field': getattr(exc, 'failed_field', '')}, 1
    (root / 'result.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    return status


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
