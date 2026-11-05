"""Isolated Windows folder chooser for local browser exports."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile


def choose_export_folder(initial=None):
    return run_picker('folder', initial).get('folder')


def choose_pdf_files():
    return run_picker('pdf').get('files', [])


def run_picker(mode, initial=None):
    if sys.platform != 'win32':
        raise ValueError('系统文件夹选择仅支持 Windows 本机服务，请改用 ZIP 压缩包导出')
    with tempfile.TemporaryDirectory(prefix='shitu-export-picker-') as folder:
        result_file = Path(folder) / 'result.json'
        result_file.with_suffix('.request').write_text(json.dumps({'mode': mode, 'initial': initial}), encoding='utf-8')
        command = ([sys.executable, '--folder-picker', str(result_file)] if getattr(sys, 'frozen', False)
                   else [sys.executable, '-m', 'backend.folder_picker', str(result_file)])
        try:
            result = subprocess.run(command, timeout=180, stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired as exc:
            raise ValueError('选择文件夹超时，已关闭选择框，请重试') from exc
        if not result_file.exists():
            raise ValueError('无法打开系统文件夹选择框，请确认 Windows 桌面未锁定')
        data = json.loads(result_file.read_text(encoding='utf-8'))
        if result.returncode or data.get('error'):
            raise ValueError(data.get('error') or '系统文件夹选择失败')
        return data


def select_pdf_files():
    import pywintypes
    import win32con
    import win32gui

    try:
        names, _, _ = win32gui.GetOpenFileNameW(
            Title='拾图 Shitu：请选择 PDF 文件', Filter='PDF 文件\0*.pdf\0', FilterIndex=1,
            Flags=win32con.OFN_ALLOWMULTISELECT | win32con.OFN_EXPLORER | win32con.OFN_FILEMUSTEXIST,
        )
    except pywintypes.error as exc:
        if exc.winerror == 0:
            return []
        raise
    parts = names.rstrip('\0').split('\0')
    return parts if len(parts) == 1 else [str(Path(parts[0]) / name) for name in parts[1:]]


def select_folder(initial=None):
    import ctypes
    from ctypes import wintypes as wt

    callback_type = ctypes.WINFUNCTYPE(ctypes.c_int, wt.HWND, wt.UINT, wt.LPARAM, wt.LPARAM)

    class BrowseInfo(ctypes.Structure):
        _fields_ = [('owner', wt.HWND), ('root', ctypes.c_void_p), ('display', wt.LPWSTR),
                    ('title', wt.LPCWSTR), ('flags', wt.UINT), ('callback', callback_type),
                    ('param', wt.LPARAM), ('image', ctypes.c_int)]

    shell = ctypes.WinDLL('shell32', use_last_error=True)
    ole = ctypes.WinDLL('ole32')
    user = ctypes.WinDLL('user32')
    ole.CoInitializeEx.argtypes = [ctypes.c_void_p, wt.DWORD]
    ole.CoInitializeEx.restype = ctypes.c_long
    ole.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    shell.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BrowseInfo)]
    shell.SHBrowseForFolderW.restype = ctypes.c_void_p
    shell.SHGetPathFromIDListEx.argtypes = [ctypes.c_void_p, wt.LPWSTR, wt.DWORD, wt.DWORD]
    shell.SHGetPathFromIDListEx.restype = wt.BOOL
    user.SetForegroundWindow.argtypes = [wt.HWND]
    user.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
    user.SetWindowPos.argtypes = [wt.HWND, wt.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.UINT]

    @callback_type
    def activate(hwnd, message, param, data):
        if message == 1:  # BFFM_INITIALIZED; raise only this process's chooser.
            if initial:
                selection = ctypes.create_unicode_buffer(initial)
                user.SendMessageW(hwnd, 0x467, 1, ctypes.addressof(selection))
            user.SetWindowPos(hwnd, ctypes.c_void_p(-1), 0, 0, 0, 0, 0x3)
            user.SetForegroundWindow(hwnd)
        return 0

    initialized = ole.CoInitializeEx(None, 2)
    if initialized < 0:
        raise ValueError('无法初始化系统文件夹选择框')
    pidl = None
    try:
        display = ctypes.create_unicode_buffer(32768)
        info = BrowseInfo(None, None, ctypes.cast(display, wt.LPWSTR), '拾图 Shitu：请选择图片文件夹', 0x41, activate, 0, 0)
        pidl = shell.SHBrowseForFolderW(ctypes.byref(info))
        if not pidl:
            return None
        path = ctypes.create_unicode_buffer(32768)
        if not shell.SHGetPathFromIDListEx(pidl, path, len(path), 0):
            raise ValueError('请选择本地文件系统中的文件夹')
        return path.value
    finally:
        if pidl:
            ole.CoTaskMemFree(pidl)
        ole.CoUninitialize()


def main(result_file):
    try:
        request_path = Path(result_file).with_suffix('.request')
        options = json.loads(request_path.read_text(encoding='utf-8')) if request_path.exists() else {}
        if options.get('mode') == 'pdf':
            result = {'files': select_pdf_files()}
        else:
            result = {'folder': select_folder(options.get('initial'))}
        status = 0
    except Exception as exc:
        result, status = {'error': str(exc)}, 1
    Path(result_file).write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    return status


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
