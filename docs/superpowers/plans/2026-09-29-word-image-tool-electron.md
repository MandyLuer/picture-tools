# Word 图片提取与重命名工具 — Electron 客户端实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把现有单文件 Tkinter 程序 `image-editer.py` 重构为「Python 后端 + Vue 前端 + Electron 客户端」的三层应用，保留全部图片能力并修掉已记录的缺陷。

**Architecture:** Python 后端是无 GUI 依赖的独立进程，通过 stdio 行分隔 JSON-RPC 与 Electron 主进程通信；同一份 `rpc.py` 分发器另有开发用 HTTP 传输。渲染进程是 Vue 3 三栏工作台，经 `transport.js` 统一访问后端；本地图片由主进程注册的 `appimg://` 协议按白名单提供，不走 base64。图片编辑采用「源文件 + 操作栈」的非破坏性模型，只在保存时编码一次。

**Tech Stack:** Python 3.10（Pillow、pytest）、Node 24、Vue 3、Vite 5、Element Plus、Pinia、Electron 31、electron-builder、Vitest、Playwright。

**Spec:** `docs/superpowers/specs/2026-09-29-word-image-tool-electron-design.md`

## Global Constraints

- 目标平台仅 Windows；不实现也不验证 macOS / Linux 构建。
- Python 运行时为 3.10.8（`venv\Scripts\python.exe`）。**不得使用 3.11+ 语法**（无 `Self`、无 `except*`、无 `tomllib`）。
- 后端依赖只有 `pillow`；打包依赖 `pyinstaller`；开发期依赖 `pytest`。**不得引入 `lxml` 与 `python-docx`**。
- 后端三个服务模块**不得 import tkinter 或任何 GUI 库**。
- 所有文件读写显式传 `encoding="utf-8"`；子进程参数一律以数组传递，绝不拼接 shell 字符串。
- 后端绝不生成 `_cropped` 或 `_compressed` 文件；绝不删除用户输出目录中的既有文件。
- 图片编辑只有四个原子操作：`rotate`（90 的整数倍，`expand=True`）、`flip`（`horizontal` | `vertical`）、`crop`（源图像坐标系 `{left, top, right, bottom}`）、`compress`（`{quality, format}`）。缩放是视图状态，不进操作栈。
- 冲突策略取值只有三个：`reject`（默认）、`autoSuffix`、`overwrite`。
- Electron 必须 `contextIsolation: true`、`sandbox: true`、`nodeIntegration: false`；禁止加载任何远程内容；CSP 禁止 `unsafe-eval`。
- 后端路径一律先 `Path.resolve()` 再做前缀校验；白名单外路径必须拒绝。
- 提交信息用中文，格式 `feat: ...` / `fix: ...` / `test: ...` / `chore: ...` / `docs: ...`。
- 每个任务结束时工作区必须可运行且测试全绿，不存在半成品文件。

## 文件结构

**后端（`backend/`）——每文件一个职责**

| 文件 | 职责 |
|---|---|
| `__main__.py` | 命令行入口，解析 `--stdio` / `--serve` / `--selftest` |
| `errors.py` | 错误码常量和 `AppError` 异常 |
| `security.py` | 路径白名单：`PathGuard` |
| `cache.py` | 预览渲染缓存：LRU + 磁盘文件 + `cacheKey` |
| `transport_stdio.py` | stdin/stdout 行分隔 JSON-RPC 读写循环 |
| `transport_http.py` | 开发用 HTTP 适配器（127.0.0.1 + 随机端口 + token） |
| `rpc.py` | 方法名到函数的静态映射表与分发 |
| `services/docx_extract.py` | docx 媒体清单、保序、提取写盘 |
| `services/image_ops.py` | 格式嗅探、操作栈重放、签名、编码保存 |
| `services/library.py` | 目录扫描排序、单张/批量重命名 |

**Electron 主进程（`app/main/`）**

| 文件 | 职责 |
|---|---|
| `main.js` | 窗口创建、应用生命周期、退出清理 |
| `backend.js` | Python 子进程的启动、请求转发、崩溃重启 |
| `ipc.js` | 注册 IPC handler，把渲染进程请求转发给后端 |
| `protocol.js` | 注册并处理 `appimg://` 协议 |

**渲染进程（`app/renderer/src/`）**

| 文件 | 职责 |
|---|---|
| `api/transport.js` | 统一 `request` / `onEvent`，内部选择 stdio 或 HTTP |
| `stores/library.js` | 文件列表、选中集合、搜索、排序、磁盘指纹 |
| `stores/editor.js` | 当前图片、`ops` 栈、`historyIndex`、已保存基线签名 |
| `stores/tasks.js` | 长任务 id、进度、取消标志 |
| `components/Workspace.vue` | 三栏骨架与栏宽拖拽 |
| `components/pane/ListPane.vue` | 左栏列表（虚拟滚动、缩略图懒加载） |
| `components/pane/PreviewPane.vue` | 中栏画布（缩放、平移、裁剪遮罩） |
| `components/pane/InspectorPane.vue` | 右栏折叠分区（文件信息、重命名、编辑参数） |
| `components/StatusBar.vue` | 底部状态栏 |
| `components/Toolbar.vue` | 顶部工具栏 |
| `composables/useCropCoords.js` | 画布坐标 ↔ 源图像坐标换算 |
| `utils/sortImages.js` | 纯排序比较函数（可单测） |
| `styles/tokens.css` | 设计令牌（浅色/深色） |

---

## 里程碑 M1：项目骨架与工具链

### Task 1: 建立项目骨架、依赖清单与后端运行入口

**Files:**
- Create: `package.json`
- Create: `requirements.txt`（覆盖现有内容）
- Create: `requirements-dev.txt`
- Create: `backend/__init__.py`
- Create: `backend/errors.py`
- Create: `backend/__main__.py`
- Create: `.gitignore`（追加新条目）
- Test: `backend/tests/__init__.py`, `backend/tests/test_cli.py`

**Interfaces:**
- Consumes: 无（首个任务）
- Produces: `backend.__main__.main(argv: list[str]) -> int`；`backend.errors.AppError`；错误码常量 `E_INTERNAL`、`E_BACKEND_DOWN`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/__init__.py`（空文件）和 `backend/tests/test_cli.py`：

```python
import pytest

from backend.__main__ import main


def test_no_args_returns_usage_error(capsys):
    exit_code = main([])
    captured = capsys.readouterr()
    assert exit_code == 2
    assert "--stdio" in captured.err


def test_unknown_arg_returns_usage_error(capsys):
    exit_code = main(["--nope"])
    captured = capsys.readouterr()
    assert exit_code == 2
    assert "unknown" in captured.err.lower()
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_cli.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend'`（因为 `backend/__main__.py` 还不存在）

- [ ] **Step 3: 建立依赖清单**

覆盖写入 `requirements.txt`：

```
pillow==10.4.0
```

创建 `requirements-dev.txt`：

```
-r requirements.txt
pytest==8.3.3
```

创建 `package.json`：

```json
{
  "name": "picture-tools",
  "version": "2.0.0",
  "private": true,
  "description": "Word 图片提取与重命名工具",
  "main": "app/main/main.js",
  "type": "module",
  "scripts": {
    "dev": "concurrently -k \"npm:dev:renderer\" \"npm:dev:electron\"",
    "dev:renderer": "vite",
    "dev:electron": "wait-on tcp:5173 && cross-env ELECTRON_DEV=1 electron .",
    "test:backend": "venv\\Scripts\\python.exe -m pytest backend/tests -v",
    "test:unit": "vitest run",
    "test:e2e": "playwright test",
    "backend": "pyinstaller --noconfirm --onedir --console --name backend --distpath build/backend backend/__main__.py",
    "build:renderer": "vite build",
    "build": "npm run build:renderer && npm run backend",
    "dist": "npm run build && electron-builder --win"
  },
  "devDependencies": {
    "@playwright/test": "^1.48.0",
    "@vitejs/plugin-vue": "^5.1.4",
    "concurrently": "^9.0.1",
    "cross-env": "^7.0.3",
    "electron": "^31.7.0",
    "electron-builder": "^25.1.8",
    "vite": "^5.4.10",
    "vitest": "^2.1.4",
    "wait-on": "^8.0.1"
  },
  "dependencies": {
    "element-plus": "^2.8.8",
    "pinia": "^2.2.6",
    "vue": "^3.5.13"
  }
}
```

- [ ] **Step 4: 创建错误码模块**

创建 `backend/__init__.py`（空文件），创建 `backend/errors.py`：

```python
"""错误码与异常类型。前端据 code 字段决定提示文案与建议动作。"""

E_DOCX_NOT_FOUND = "E_DOCX_NOT_FOUND"
E_DOCX_NOT_ZIP = "E_DOCX_NOT_ZIP"
E_DOCX_NO_MEDIA = "E_DOCX_NO_MEDIA"
E_DOCX_ORDER_GUESSED = "E_DOCX_ORDER_GUESSED"
E_OUTDIR_NOT_WRITABLE = "E_OUTDIR_NOT_WRITABLE"
E_NAME_EXISTS = "E_NAME_EXISTS"
E_NAME_INVALID = "E_NAME_INVALID"
E_IMAGE_DECODE = "E_IMAGE_DECODE"
E_FORMAT_MISMATCH = "E_FORMAT_MISMATCH"
E_DISK_FULL = "E_DISK_FULL"
E_TASK_CANCELLED = "E_TASK_CANCELLED"
E_BACKEND_DOWN = "E_BACKEND_DOWN"
E_INTERNAL = "E_INTERNAL"

# 文件名中 Windows 不允许出现的字符
ILLEGAL_NAME_CHARS = '\\/:*?"<>|'


class AppError(Exception):
    """带错误码的应用异常。message 面向用户，detail 面向诊断。"""

    def __init__(self, code, message, detail=None, hint=None, context=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.detail = detail
        self.hint = hint
        self.context = context or {}

    def to_dict(self):
        return {
            "code": self.code,
            "message": self.message,
            "detail": self.detail,
            "hint": self.hint,
            "context": self.context,
        }
```

- [ ] **Step 5: 创建命令行入口**

创建 `backend/__main__.py`：

```python
"""后端入口：--stdio（生产）| --serve（开发）| --selftest（自检）。"""

import sys

USAGE = "usage: python -m backend --stdio | --serve | --selftest"


def main(argv):
    if not argv:
        print(USAGE, file=sys.stderr)
        return 2

    mode = argv[0]
    if mode == "--stdio":
        from backend.transport_stdio import serve_stdio

        return serve_stdio()
    if mode == "--serve":
        from backend.transport_http import serve_http

        return serve_http()
    if mode == "--selftest":
        from backend.selftest import run_selftest

        return run_selftest()

    print("unknown argument: %s" % mode, file=sys.stderr)
    print(USAGE, file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

- [ ] **Step 6: 补齐 `__main__` 的延迟导入占位**

`--stdio` / `--serve` / `--selftest` 指向的模块要到 Task 5、Task 17、Task 10 才创建。为让本任务可独立验证，先创建三个最小占位模块，后续任务会替换其内容。创建 `backend/transport_stdio.py`：

```python
def serve_stdio():
    raise NotImplementedError("implemented in Task 5")
```

创建 `backend/transport_http.py`：

```python
def serve_http():
    raise NotImplementedError("implemented in Task 5")
```

创建 `backend/selftest.py`：

```python
def run_selftest():
    raise NotImplementedError("implemented in Task 10")
```

- [ ] **Step 7: 追加 .gitignore 条目**

在现有 `.gitignore` 末尾追加：

```
# Node 与前端构建产物
node_modules/
app/renderer/dist/
dist/
release/

# Python 与打包产物
__pycache__/
*.pyc
.pytest_cache/
build/backend/
build/*.spec
```

注意：现有仓库已有 `build/` 与 `venv/` 相关条目，追加时不要重复，若已存在则跳过对应行。

- [ ] **Step 8: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_cli.py -v`
Expected: 2 passed

- [ ] **Step 9: 确认后端可被当作模块调用**

Run: `venv\Scripts\python.exe -m backend`
Expected: 退出码 2，stderr 输出 `usage: python -m backend --stdio | --serve | --selftest`

- [ ] **Step 10: 提交**

```bash
git add package.json requirements.txt requirements-dev.txt .gitignore backend/
git commit -m "chore: 建立后端骨架、依赖清单与命令行入口"
```

---

### Task 2: 实现路径白名单安全模块

**Files:**
- Create: `backend/security.py`
- Test: `backend/tests/test_security.py`

**Interfaces:**
- Consumes: `backend.errors.AppError`（Task 1）
- Produces: `PathGuard` 类——`PathGuard(roots: Iterable[str] | None = None)`、`.allow(path: str) -> Path`、`.allow_root(path: str) -> Path`、`.is_allowed(path: str) -> bool`、`.roots -> list[Path]`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_security.py`：

```python
import os

import pytest

from backend.errors import AppError
from backend.security import PathGuard


def test_unregistered_path_is_rejected(tmp_path):
    guard = PathGuard([tmp_path / "allowed"])
    with pytest.raises(AppError) as excinfo:
        guard.allow(str(tmp_path / "other" / "a.png"))
    assert excinfo.value.code == "E_PATH_DENIED"


def test_path_inside_root_is_allowed(tmp_path):
    root = tmp_path / "allowed"
    root.mkdir()
    guard = PathGuard([root])
    target = root / "sub" / "a.png"
    target.parent.mkdir()
    target.write_bytes(b"x")
    assert guard.allow(str(target)) == target.resolve()


def test_parent_traversal_is_rejected(tmp_path):
    root = tmp_path / "allowed"
    root.mkdir()
    guard = PathGuard([root])
    with pytest.raises(AppError):
        guard.allow(str(root / ".." / "outside.png"))


def test_prefix_sibling_directory_is_rejected(tmp_path):
    """allowed_evil 与 allowed 共享字符串前缀，但不该被放行。"""
    root = tmp_path / "allowed"
    root.mkdir()
    evil = tmp_path / "allowed_evil"
    evil.mkdir()
    guard = PathGuard([root])
    with pytest.raises(AppError):
        guard.allow(str(evil / "a.png"))


def test_root_itself_is_allowed(tmp_path):
    root = tmp_path / "allowed"
    root.mkdir()
    guard = PathGuard([root])
    assert guard.allow(str(root)) == root.resolve()


def test_allow_root_registers_new_directory(tmp_path):
    guard = PathGuard()
    target = tmp_path / "picked"
    target.mkdir()
    guard.allow_root(str(target))
    assert guard.is_allowed(str(target / "a.png"))
    assert len(guard.roots) == 1


def test_is_allowed_never_raises(tmp_path):
    guard = PathGuard([tmp_path / "allowed"])
    assert guard.is_allowed(str(tmp_path / "elsewhere" / "a.png")) is False
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_security.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend.security'`

- [ ] **Step 3: 实现 PathGuard**

创建 `backend/security.py`：

```python
"""路径白名单。后端只允许访问用户显式选中或已授权目录下的文件。"""

import os
from pathlib import Path

from backend.errors import AppError

E_PATH_DENIED = "E_PATH_DENIED"

# Windows 长路径前缀，超过 260 字符时使用
_LONG_PATH_PREFIX = "\\\\?\\"


class PathGuard:
    """维护一组已授权根目录，拒绝集合之外的任何路径。"""

    def __init__(self, roots=None):
        self._roots = []
        if roots:
            for root in roots:
                self.allow_root(str(root))

    @property
    def roots(self):
        return list(self._roots)

    def allow_root(self, path):
        """登记一个授权根目录（用户在原生对话框中选中的目录）。"""
        resolved = self._resolve(path)
        if resolved not in self._roots:
            self._roots.append(resolved)
        return resolved

    def is_allowed(self, path):
        """判断路径是否在白名单内。绝不抛异常。"""
        try:
            candidate = self._resolve(path)
        except (OSError, ValueError):
            return False
        for root in self._roots:
            if candidate == root or root in candidate.parents:
                return True
        return False

    def allow(self, path):
        """校验路径并返回解析后的 Path；不在白名单内则抛 AppError。"""
        resolved = self._resolve(path)
        if not self.is_allowed(resolved):
            raise AppError(
                E_PATH_DENIED,
                "拒绝访问授权目录之外的路径",
                detail="path=%s roots=%s" % (resolved, [str(r) for r in self._roots]),
                hint="请先通过打开文件或打开文件夹授权该目录",
                context={"path": str(resolved)},
            )
        return resolved

    @staticmethod
    def _resolve(path):
        """解析为绝对真实路径，展开长路径前缀并消解符号链接。"""
        text = os.fspath(path)
        if text.startswith(_LONG_PATH_PREFIX):
            text = text[len(_LONG_PATH_PREFIX):]
        return Path(text).expanduser().resolve()
```

- [ ] **Step 4: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_security.py -v`
Expected: 7 passed

- [ ] **Step 5: 补充符号链接逃逸测试（Windows 下需要权限，不满足则跳过）**

在 `backend/tests/test_security.py` 末尾追加：

```python
def test_symlink_escape_is_rejected(tmp_path):
    """目录内的符号链接指向外部时，必须因为 resolve() 而落在白名单外。"""
    root = tmp_path / "allowed"
    root.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "secret.png").write_bytes(b"x")

    link = root / "escape"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except (OSError, NotImplementedError):
        pytest.skip("当前环境不支持创建符号链接（需开发者模式或管理员权限）")

    guard = PathGuard([root])
    with pytest.raises(AppError):
        guard.allow(str(link / "secret.png"))
```

- [ ] **Step 6: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_security.py -v`
Expected: 8 passed（或 7 passed, 1 skipped，若环境不允许创建符号链接）

- [ ] **Step 7: 提交**

```bash
git add backend/security.py backend/tests/test_security.py
git commit -m "feat: 新增后端路径白名单安全模块"
```

---

### Task 3: 实现图片格式嗅探与解码

**Files:**
- Create: `backend/services/__init__.py`
- Create: `backend/services/image_ops.py`
- Test: `backend/tests/test_format_sniff.py`

**Interfaces:**
- Consumes: `backend.errors.AppError`（Task 1）
- Produces: `sniff_format(data: bytes) -> str | None`（返回 `"png"`/`"jpeg"`/`"gif"`/`"bmp"`/`"tiff"`/`"webp"`）；`EXTENSION_BY_FORMAT: dict[str, str]`；`open_image(path: str) -> PIL.Image.Image`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_format_sniff.py`：

```python
import io

import pytest
from PIL import Image

from backend.errors import AppError
from backend.services.image_ops import (
    EXTENSION_BY_FORMAT,
    open_image,
    sniff_format,
)

FIXTURES = [
    ("png", "PNG"),
    ("jpeg", "JPEG"),
    ("gif", "GIF"),
    ("bmp", "BMP"),
    ("tiff", "TIFF"),
    ("webp", "WEBP"),
]


def _encode(fmt, size=(4, 4), mode="RGB"):
    buf = io.BytesIO()
    Image.new(mode, size, (10, 20, 30)).save(buf, format=fmt)
    return buf.getvalue()


@pytest.mark.parametrize("expected,pil_format", FIXTURES)
def test_sniff_each_supported_format(expected, pil_format):
    assert sniff_format(_encode(pil_format)) == expected


def test_sniff_rejects_non_image():
    assert sniff_format(b"this is definitely not an image") is None
    assert sniff_format(b"") is None


def test_sniff_needs_only_leading_bytes():
    """只需文件头即可判定，模拟从 zip 流中只读前若干字节。"""
    data = _encode("PNG")
    assert sniff_format(data[:16]) == "png"


def test_extension_mapping_covers_all_formats():
    for fmt in ("png", "jpeg", "gif", "bmp", "tiff", "webp"):
        assert fmt in EXTENSION_BY_FORMAT
    assert EXTENSION_BY_FORMAT["jpeg"] == ".jpg"


def test_open_image_loads_real_file(tmp_path):
    path = tmp_path / "a.png"
    path.write_bytes(_encode("PNG", size=(7, 9)))
    img = open_image(str(path))
    assert img.size == (7, 9)


def test_open_image_raises_structured_error_on_garbage(tmp_path):
    path = tmp_path / "broken.png"
    path.write_bytes(b"not an image at all")
    with pytest.raises(AppError) as excinfo:
        open_image(str(path))
    assert excinfo.value.code == "E_IMAGE_DECODE"
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_format_sniff.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend.services'`

- [ ] **Step 3: 实现嗅探与解码**

创建 `backend/services/__init__.py`（空文件），创建 `backend/services/image_ops.py`：

```python
"""图片格式嗅探、解码与操作栈重放。

替代原 ZoomableImage 中的编辑方法。只依赖 Pillow，不依赖任何 GUI 库。
"""

from PIL import Image, UnidentifiedImageError

from backend.errors import E_IMAGE_DECODE, AppError

# 格式名 -> 输出扩展名
EXTENSION_BY_FORMAT = {
    "png": ".png",
    "jpeg": ".jpg",
    "gif": ".gif",
    "bmp": ".bmp",
    "tiff": ".tiff",
    "webp": ".webp",
}

# 不支持透明通道的格式
OPAQUE_FORMATS = {"jpeg", "bmp"}


def sniff_format(data):
    """仅凭文件头魔数判断图片格式。

    替代在 Python 3.13 中已被移除的 imghdr 模块。
    返回小写格式名，无法判定时返回 None。
    """
    if not data or len(data) < 12:
        return None

    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return "gif"
    if data.startswith(b"BM"):
        return "bmp"
    if data.startswith(b"II*\x00") or data.startswith(b"MM\x00*"):
        return "tiff"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"

    return None


def open_image(path):
    """打开并完整加载图片。失败时抛结构化错误而非 Pillow 原始异常。"""
    try:
        image = Image.open(path)
        image.load()
        return image
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise AppError(
            E_IMAGE_DECODE,
            "图片已损坏或格式不受支持",
            detail="%s: %s" % (type(exc).__name__, exc),
            hint="可跳过该文件，或先用其他工具转换格式",
            context={"path": str(path)},
        ) from exc
```

- [ ] **Step 4: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_format_sniff.py -v`
Expected: 12 passed

- [ ] **Step 5: 提交**

```bash
git add backend/services/ backend/tests/test_format_sniff.py
git commit -m "feat: 新增图片格式嗅探与解码，替换已移除的 imghdr"
```

---

### Task 4: 实现操作栈重放与裁剪坐标映射

**Files:**
- Modify: `backend/services/image_ops.py`（追加）
- Test: `backend/tests/test_image_ops.py`

**Interfaces:**
- Consumes: `open_image`、`EXTENSION_BY_FORMAT`、`OPAQUE_FORMATS`（Task 3）
- Produces: `apply_ops(source_path: str, ops: list[dict]) -> PIL.Image.Image`；`normalize_ops(ops: list[dict]) -> list[dict]`；操作字典形状固定为 `{"type": "rotate", "angle": int}` / `{"type": "flip", "axis": "horizontal"|"vertical"}` / `{"type": "crop", "rect": {"left","top","right","bottom"}}` / `{"type": "compress", "quality": int, "format": str}`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_image_ops.py`：

```python
import pytest
from PIL import Image

from backend.errors import AppError
from backend.services.image_ops import apply_ops, normalize_ops


def _make(path, size=(40, 20), color=(200, 100, 50)):
    Image.new("RGB", size, color).save(path)
    return str(path)


def test_empty_ops_returns_original_size(tmp_path):
    path = _make(tmp_path / "a.png")
    assert apply_ops(path, []).size == (40, 20)


def test_rotate_90_twice_gives_180(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    result = apply_ops(path, [{"type": "rotate", "angle": 90}, {"type": "rotate", "angle": 90}])
    assert result.size == (40, 20)


def test_rotate_90_once_swaps_dimensions(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    assert apply_ops(path, [{"type": "rotate", "angle": 90}]).size == (20, 40)


def test_rotate_four_times_restores_original_size(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    ops = [{"type": "rotate", "angle": 90}] * 4
    assert apply_ops(path, ops).size == (40, 20)


def test_rotate_rejects_non_right_angle(tmp_path):
    path = _make(tmp_path / "a.png")
    with pytest.raises(AppError) as excinfo:
        apply_ops(path, [{"type": "rotate", "angle": 45}])
    assert excinfo.value.code == "E_INVALID_OP"


def test_flip_horizontal_swaps_left_and_right_pixels(tmp_path):
    img = Image.new("RGB", (2, 1))
    img.putpixel((0, 0), (255, 0, 0))
    img.putpixel((1, 0), (0, 0, 255))
    path = str(tmp_path / "a.png")
    img.save(path)

    result = apply_ops(path, [{"type": "flip", "axis": "horizontal"}])
    assert result.getpixel((0, 0)) == (0, 0, 255)
    assert result.getpixel((1, 0)) == (255, 0, 0)


def test_flip_vertical_swaps_top_and_bottom_pixels(tmp_path):
    img = Image.new("RGB", (1, 2))
    img.putpixel((0, 0), (255, 0, 0))
    img.putpixel((0, 1), (0, 0, 255))
    path = str(tmp_path / "a.png")
    img.save(path)

    result = apply_ops(path, [{"type": "flip", "axis": "vertical"}])
    assert result.getpixel((0, 0)) == (0, 0, 255)


def test_crop_uses_source_image_coordinates(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    result = apply_ops(
        path,
        [{"type": "crop", "rect": {"left": 10, "top": 5, "right": 30, "bottom": 15}}],
    )
    assert result.size == (20, 10)


def test_crop_then_rotate_composes_in_order(tmp_path):
    """裁剪后再旋转：先裁成 20x10，再转 90 度得 10x20。"""
    path = _make(tmp_path / "a.png", size=(40, 20))
    ops = [
        {"type": "crop", "rect": {"left": 0, "top": 0, "right": 20, "bottom": 10}},
        {"type": "rotate", "angle": 90},
    ]
    assert apply_ops(path, ops).size == (10, 20)


def test_rotate_then_crop_composes_in_order(tmp_path):
    """旋转后再裁剪：先转成 20x40，再裁 10x20 区域。"""
    path = _make(tmp_path / "a.png", size=(40, 20))
    ops = [
        {"type": "rotate", "angle": 90},
        {"type": "crop", "rect": {"left": 0, "top": 0, "right": 10, "bottom": 20}},
    ]
    assert apply_ops(path, ops).size == (10, 20)


def test_crop_rect_is_clamped_to_image_bounds(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    result = apply_ops(
        path,
        [{"type": "crop", "rect": {"left": -5, "top": -5, "right": 999, "bottom": 999}}],
    )
    assert result.size == (40, 20)


def test_crop_with_inverted_coordinates_is_normalized(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    result = apply_ops(
        path,
        [{"type": "crop", "rect": {"left": 30, "top": 15, "right": 10, "bottom": 5}}],
    )
    assert result.size == (20, 10)


def test_crop_too_small_is_rejected(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    with pytest.raises(AppError) as excinfo:
        apply_ops(
            path,
            [{"type": "crop", "rect": {"left": 0, "top": 0, "right": 5, "bottom": 5}}],
        )
    assert excinfo.value.code == "E_CROP_TOO_SMALL"


def test_compress_op_does_not_change_pixels_or_size(tmp_path):
    path = _make(tmp_path / "a.png", size=(40, 20))
    result = apply_ops(path, [{"type": "compress", "quality": 50, "format": "jpeg"}])
    assert result.size == (40, 20)


def test_apply_ops_is_repeatable(tmp_path):
    """同一操作栈重复执行必须得到相同尺寸与相同像素。"""
    path = _make(tmp_path / "a.png", size=(40, 20))
    ops = [
        {"type": "flip", "axis": "horizontal"},
        {"type": "crop", "rect": {"left": 5, "top": 5, "right": 35, "bottom": 15}},
    ]
    first = apply_ops(path, ops)
    second = apply_ops(path, ops)
    assert first.size == second.size
    assert list(first.getdata()) == list(second.getdata())


def test_unknown_op_is_rejected(tmp_path):
    path = _make(tmp_path / "a.png")
    with pytest.raises(AppError) as excinfo:
        apply_ops(path, [{"type": "resize", "width": 10, "height": 10}])
    assert excinfo.value.code == "E_INVALID_OP"


def test_normalize_ops_strips_compress_and_preserves_geometry_order(tmp_path):
    ops = [
        {"type": "rotate", "angle": 90},
        {"type": "compress", "quality": 50, "format": "jpeg"},
        {"type": "flip", "axis": "vertical"},
    ]
    normalized = normalize_ops(ops)
    assert [op["type"] for op in normalized] == ["rotate", "flip"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_image_ops.py -v`
Expected: FAIL — `ImportError: cannot import name 'apply_ops'`

- [ ] **Step 3: 追加错误码**

在 `backend/errors.py` 的错误码常量区追加：

```python
E_INVALID_OP = "E_INVALID_OP"
E_CROP_TOO_SMALL = "E_CROP_TOO_SMALL"
```

- [ ] **Step 4: 实现操作栈重放**

在 `backend/services/image_ops.py` 末尾追加：

```python
# 裁剪区域的宽高下限，与原 Tkinter 版保持一致
MIN_CROP_EDGE = 10

# 参与像素运算的操作类型；compress 只影响编码，不改变像素
GEOMETRY_OPS = ("rotate", "flip", "crop")


def normalize_ops(ops):
    """去掉 compress，只保留会改变像素的几何操作，保持原有顺序。"""
    return [op for op in (ops or []) if op.get("type") in GEOMETRY_OPS]


def _apply_rotate(image, op):
    angle = op.get("angle")
    if not isinstance(angle, int) or angle % 90 != 0:
        raise AppError(
            E_INVALID_OP,
            "旋转角度必须是 90 的整数倍",
            detail="angle=%r" % (angle,),
            context={"op": op},
        )
    return image.rotate(angle, expand=True)


def _apply_flip(image, op):
    axis = op.get("axis")
    if axis == "horizontal":
        return image.transpose(Image.FLIP_LEFT_RIGHT)
    if axis == "vertical":
        return image.transpose(Image.FLIP_TOP_BOTTOM)
    raise AppError(
        E_INVALID_OP,
        "翻转方向只能是 horizontal 或 vertical",
        detail="axis=%r" % (axis,),
        context={"op": op},
    )


def _apply_crop(image, op):
    rect = op.get("rect") or {}
    try:
        left = int(rect["left"])
        top = int(rect["top"])
        right = int(rect["right"])
        bottom = int(rect["bottom"])
    except (KeyError, TypeError, ValueError) as exc:
        raise AppError(
            E_INVALID_OP,
            "裁剪区域参数不完整",
            detail=str(exc),
            context={"op": op},
        ) from exc

    width, height = image.size
    left, right = sorted((left, right))
    top, bottom = sorted((top, bottom))
    left = max(0, min(width, left))
    right = max(0, min(width, right))
    top = max(0, min(height, top))
    bottom = max(0, min(height, bottom))

    if right - left < MIN_CROP_EDGE or bottom - top < MIN_CROP_EDGE:
        raise AppError(
            E_CROP_TOO_SMALL,
            "裁剪区域太小，请至少保留 %d 像素" % MIN_CROP_EDGE,
            detail="rect=%s" % ({"left": left, "top": top, "right": right, "bottom": bottom},),
            hint="重新拉框，或直接取消裁剪",
            context={"op": op},
        )

    return image.crop((left, top, right, bottom))


_DISPATCH = {
    "rotate": _apply_rotate,
    "flip": _apply_flip,
    "crop": _apply_crop,
}


def apply_ops(source_path, ops):
    """把操作栈按顺序重放到源图片上，返回结果图像。

    「编辑后的图片长什么样」的唯一真相来源：预览与保存都经过这个函数。
    """
    image = open_image(source_path)
    for op in ops or []:
        op_type = op.get("type")
        handler = _DISPATCH.get(op_type)
        if handler is None:
            raise AppError(
                E_INVALID_OP,
                "不支持的操作类型：%s" % (op_type,),
                hint="当前仅支持 rotate / flip / crop / compress",
                context={"op": op},
            )
        image = handler(image, op)
    return image
```

- [ ] **Step 5: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_image_ops.py -v`
Expected: 17 passed

- [ ] **Step 6: 提交**

```bash
git add backend/errors.py backend/services/image_ops.py backend/tests/test_image_ops.py
git commit -m "feat: 实现图片操作栈重放与裁剪坐标映射"
```

---

### Task 5: 实现操作栈签名与脏标记

**Files:**
- Modify: `backend/services/image_ops.py`（追加）
- Test: `backend/tests/test_signature.py`

**Interfaces:**
- Consumes: `normalize_ops`（Task 4）
- Produces: `ops_signature(source_path: str, ops: list[dict]) -> str`（返回 64 位十六进制 sha256）

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_signature.py`：

```python
from PIL import Image

from backend.services.image_ops import ops_signature


def _make(path, color=(200, 100, 50)):
    Image.new("RGB", (10, 10), color).save(path)
    return str(path)


def test_same_ops_give_same_signature(tmp_path):
    path = _make(tmp_path / "a.png")
    ops = [{"type": "rotate", "angle": 90}]
    assert ops_signature(path, ops) == ops_signature(path, ops)


def test_signature_is_hex_sha256(tmp_path):
    path = _make(tmp_path / "a.png")
    sig = ops_signature(path, [])
    assert len(sig) == 64
    int(sig, 16)  # 不抛异常即为合法十六进制


def test_op_order_is_significant(tmp_path):
    path = _make(tmp_path / "a.png")
    first = [{"type": "rotate", "angle": 90}, {"type": "flip", "axis": "horizontal"}]
    second = [{"type": "flip", "axis": "horizontal"}, {"type": "rotate", "angle": 90}]
    assert ops_signature(path, first) != ops_signature(path, second)


def test_rotating_back_to_original_restores_clean_signature(tmp_path):
    """旋转四次 90 度等价于原图，签名必须回到基线。"""
    path = _make(tmp_path / "a.png")
    baseline = ops_signature(path, [])
    spun = ops_signature(path, [{"type": "rotate", "angle": 90}] * 4)
    assert baseline == spun


def test_different_source_files_give_different_signatures(tmp_path):
    first = _make(tmp_path / "a.png", color=(200, 100, 50))
    second = _make(tmp_path / "b.png", color=(10, 20, 30))
    assert ops_signature(first, []) != ops_signature(second, [])


def test_editing_file_on_disk_changes_signature(tmp_path):
    path = _make(tmp_path / "a.png", color=(200, 100, 50))
    before = ops_signature(path, [])
    _make(tmp_path / "a.png", color=(1, 2, 3))
    assert ops_signature(path, []) != before


def test_compress_alone_does_not_dirty_geometry(tmp_path):
    """只改压缩质量不该改变几何签名，但仍需与空操作栈不同。"""
    path = _make(tmp_path / "a.png")
    baseline = ops_signature(path, [])
    compressed = ops_signature(path, [{"type": "compress", "quality": 50, "format": "jpeg"}])
    assert baseline != compressed


def test_flip_twice_restores_clean_signature(tmp_path):
    path = _make(tmp_path / "a.png")
    baseline = ops_signature(path, [])
    flipped = ops_signature(path, [{"type": "flip", "axis": "horizontal"}] * 2)
    assert baseline == flipped
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_signature.py -v`
Expected: FAIL — `ImportError: cannot import name 'ops_signature'`

- [ ] **Step 3: 实现签名**

在 `backend/services/image_ops.py` 顶部把 `import` 区改为：

```python
import hashlib
import json
import os

from PIL import Image, UnidentifiedImageError

from backend.errors import (
    E_CROP_TOO_SMALL,
    E_IMAGE_DECODE,
    E_INVALID_OP,
    AppError,
)
```

在文件末尾追加：

```python
def _source_fingerprint(source_path):
    """源文件指纹：路径 + 大小 + 修改时间。避免为了签名而读取整个文件。"""
    stat = os.stat(source_path)
    return {
        "path": os.path.abspath(source_path),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def ops_signature(source_path, ops):
    """计算「源文件 + 操作栈」的稳定哈希，用于前端脏标记。

    注意：compress 参与签名（改变输出编码），但不参与像素几何。
    四次 90 度旋转虽然像素上等价于原图，语义上仍是独立操作栈，
    因此先归一化再序列化，保证「转回原状即可回到干净」。
    """
    payload = {
        "source": _source_fingerprint(source_path),
        "ops": normalize_ops(ops),
        "encoding": _encoding_of(ops),
    }
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()
```

同时追加编码归一化辅助函数，并修正上面签名中调用它的位置：

```python
def _encoding_of(ops):
    """提取 compress 操作描述的最终编码，未指定则为 None。"""
    for op in reversed(ops or []):
        if op.get("type") == "compress":
            return {"quality": op.get("quality"), "format": op.get("format")}
    return None
```

**注意**：上一步 `ops_signature` 的 payload 里 `"ops": normalize_ops(ops)` 只含几何操作，`"encoding"` 单独承载 compress。这满足两条测试：只改压缩质量会改变签名；四次旋转互相抵消不可能自动成立——因此必须再追加一步几何规范化，见 Step 4。

- [ ] **Step 4: 让「转回原状」真正回到基线**

`normalize_ops` 只过滤 compress，不会把「旋转四次」约简成「无操作」。追加一个几何约简函数，并在 `ops_signature` 中使用它。

在 `backend/services/image_ops.py` 末尾追加：

```python
def reduce_geometry(ops):
    """在保持像素结果等价的前提下约简几何操作。

    只处理两个可安全抵消的情形：
      - 四次同向 90 度旋转 == 无操作
      - 两次同向翻转 == 无操作
      - 连续同向旋转的角度先累加
    裁剪无法与任何操作抵消，遇裁剪即停止约简的合并，但不影响其余部分。
    """
    stack = []
    for op in normalize_ops(ops):
        op_type = op.get("type")

        if op_type == "rotate" and stack and stack[-1].get("type") == "rotate":
            combined = (stack[-1]["angle"] + op["angle"]) % 360
            stack.pop()
            if combined % 360 != 0:
                stack.append({"type": "rotate", "angle": combined})
            continue

        if (
            op_type == "flip"
            and stack
            and stack[-1].get("type") == "flip"
            and stack[-1].get("axis") == op.get("axis")
        ):
            stack.pop()
            continue

        stack.append({"type": op_type, **{k: v for k, v in op.items() if k != "type"}})

    return stack
```

把 `ops_signature` 的 payload 改为使用 `reduce_geometry`：

```python
    payload = {
        "source": _source_fingerprint(source_path),
        "ops": reduce_geometry(ops),
        "encoding": _encoding_of(ops),
    }
```

注意 `{"type": op_type, **{...}}` 中键顺序可能导致 `type` 被覆盖，实际实现请写成先复制再设 type：

```python
        merged = {k: v for k, v in op.items() if k != "type"}
        merged["type"] = op_type
        stack.append(merged)
```

- [ ] **Step 5: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_signature.py -v`
Expected: 8 passed

- [ ] **Step 6: 运行全部后端测试确认无回归**

Run: `venv\Scripts\python.exe -m pytest backend/tests -v`
Expected: 全部通过（test_cli 2、test_security 8、test_format_sniff 12、test_image_ops 17、test_signature 8）

- [ ] **Step 7: 提交**

```bash
git add backend/services/image_ops.py backend/tests/test_signature.py
git commit -m "feat: 实现操作栈签名与几何约简，支持脏标记回洁"
```

---

### Task 6: 实现 docx 媒体清单与保序解析

**Files:**
- Create: `backend/services/docx_extract.py`
- Test: `backend/tests/test_docx_order.py`
- Test fixtures: `backend/tests/fixtures/make_fixtures.py`（生成三个样例 docx 的脚本）

**Interfaces:**
- Consumes: `sniff_format`（Task 3）、`AppError`（Task 1）
- Produces: `inspect_document(path: str) -> dict`，返回 `{"media": [{"partName","relId","bytes","format"}], "order": [partName...], "orderSource": "xml"|"regex"|"filename"|"guess"}`

- [ ] **Step 1: 编写样例 docx 生成脚本**

创建 `backend/tests/fixtures/make_fixtures.py`：

```python
"""生成三个样例 docx，用于验证图片保序逻辑。

不依赖 python-docx：直接按 OOXML 最小结构写 zip。
运行方式：venv\\Scripts\\python.exe backend/tests/fixtures/make_fixtures.py
"""

import io
import os
import zipfile

from PIL import Image

FIXTURE_DIR = os.path.dirname(os.path.abspath(__file__))

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{relationships}
</Relationships>"""

DRAWING_TEMPLATE = """<w:p><w:r><w:drawing>
<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
<wp:extent cx="100" cy="100"/>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>
<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:blipFill><a:blip r:embed="{rel_id}"/></pic:blipFill>
</pic:pic></a:graphicData></a:graphic>
</wp:inline></w:drawing></w:r></w:p>"""

PICT_TEMPLATE = """<w:p><w:r><w:pict>
<v:shape xmlns:v="urn:schemas-microsoft-com:vml">
<v:imagedata r:id="{rel_id}"/>
</v:shape></w:pict></w:r></w:p>"""

DOCUMENT_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>{body}</w:body></w:document>"""

VML_NS = ' xmlns:v="urn:schemas-microsoft-com:vml"'


def _png_bytes(color):
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), color).save(buf, format="PNG")
    return buf.getvalue()


def _write_docx(path, body, media, relationships):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT_TEMPLATE.format(body=body))
        z.writestr(
            "word/_rels/document.xml.rels",
            DOC_RELS_TEMPLATE.format(relationships="\n".join(relationships)),
        )
        for name, data in media.items():
            z.writestr(name, data)


def _rel(rel_id, target):
    return (
        '  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/image" Target="%s"/>' % (rel_id, target)
    )


def build_plain(path):
    """普通图文文档：插入顺序与 media 文件名顺序一致，图片 1、2、3。"""
    media = {
        "word/media/image1.png": _png_bytes((255, 0, 0)),
        "word/media/image2.png": _png_bytes((0, 255, 0)),
        "word/media/image3.png": _png_bytes((0, 0, 255)),
    }
    body = (
        DRAWING_TEMPLATE.format(rel_id="rId1")
        + DRAWING_TEMPLATE.format(rel_id="rId2")
        + DRAWING_TEMPLATE.format(rel_id="rId3")
    )
    rels = [_rel("rId1", "media/image1.png"), _rel("rId2", "media/image2.png"), _rel("rId3", "media/image3.png")]
    _write_docx(path, body, media, rels)


def build_reordered(path):
    """关键样例：插入顺序与 media 文件名顺序故意不一致。

    文档中依次插入 image3、image1、image2，若解析器偷懒按文件名排序，
    得到的结果会与 document.xml 中的顺序不同，从而暴露缺陷。
    """
    media = {
        "word/media/image1.png": _png_bytes((255, 0, 0)),
        "word/media/image2.png": _png_bytes((0, 255, 0)),
        "word/media/image3.png": _png_bytes((0, 0, 255)),
    }
    body = (
        DRAWING_TEMPLATE.format(rel_id="rId3")
        + DRAWING_TEMPLATE.format(rel_id="rId1")
        + DRAWING_TEMPLATE.format(rel_id="rId2")
    )
    rels = [_rel("rId1", "media/image1.png"), _rel("rId2", "media/image2.png"), _rel("rId3", "media/image3.png")]
    _write_docx(path, body, media, rels)


def build_vml(path):
    """旧式 VML 文档：图片用 w:pict + v:imagedata 引用，需走 pict 分支。"""
    media = {
        "word/media/image1.png": _png_bytes((255, 0, 0)),
        "word/media/image2.png": _png_bytes((0, 255, 0)),
    }
    body = PICT_TEMPLATE.format(rel_id="rId2") + PICT_TEMPLATE.format(rel_id="rId1")
    rels = [_rel("rId1", "media/image1.png"), _rel("rId2", "media/image2.png")]
    _write_docx(path, body, media, rels)


def main():
    build_plain(os.path.join(FIXTURE_DIR, "plain.docx"))
    build_reordered(os.path.join(FIXTURE_DIR, "reordered.docx"))
    build_vml(os.path.join(FIXTURE_DIR, "vml.docx"))
    print("fixtures written to %s" % FIXTURE_DIR)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 生成样例文件并确认落盘**

Run: `venv\Scripts\python.exe backend/tests/fixtures/make_fixtures.py`
Expected: 输出 `fixtures written to ...fixtures`

Run: `Get-ChildItem backend/tests/fixtures/*.docx | Select-Object Name, Length`
Expected: 三个 docx 文件，长度均大于 0

- [ ] **Step 3: 把样例文件纳入版本管理**

样例 docx 是测试资产，必须入库（否则新克隆的仓库测试无法运行）。在 `.gitignore` 中确认没有 `*.docx` 之类会误伤它们的规则；若有则追加例外：

```
!backend/tests/fixtures/*.docx
```

- [ ] **Step 4: 写失败的测试**

创建 `backend/tests/test_docx_order.py`：

```python
import os

import pytest

from backend.errors import AppError
from backend.services.docx_extract import inspect_document

FIXTURE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


def _fixture(name):
    return os.path.join(FIXTURE_DIR, name)


def test_reordered_document_follows_document_order_not_filename_order():
    """最关键的一条：必须按 document.xml 的插入顺序，而不是 media 文件名顺序。"""
    result = inspect_document(_fixture("reordered.docx"))
    assert result["orderSource"] == "xml"
    assert result["order"] == [
        "word/media/image3.png",
        "word/media/image1.png",
        "word/media/image2.png",
    ]


def test_plain_document_order_matches_insertion():
    result = inspect_document(_fixture("plain.docx"))
    assert result["order"] == [
        "word/media/image1.png",
        "word/media/image2.png",
        "word/media/image3.png",
    ]


def test_vml_pict_images_are_recognized():
    result = inspect_document(_fixture("vml.docx"))
    assert result["order"] == ["word/media/image2.png", "word/media/image1.png"]


def test_media_entries_report_format_and_size():
    result = inspect_document(_fixture("plain.docx"))
    assert len(result["media"]) == 3
    for entry in result["media"]:
        assert entry["format"] == "png"
        assert entry["bytes"] > 0
        assert entry["partName"].startswith("word/media/")


def test_media_carries_the_relationship_id():
    result = inspect_document(_fixture("plain.docx"))
    rel_ids = sorted(entry["relId"] for entry in result["media"])
    assert rel_ids == ["rId1", "rId2", "rId3"]


def test_missing_file_raises_structured_error(tmp_path):
    with pytest.raises(AppError) as excinfo:
        inspect_document(str(tmp_path / "nope.docx"))
    assert excinfo.value.code == "E_DOCX_NOT_FOUND"


def test_non_zip_file_raises_structured_error(tmp_path):
    path = tmp_path / "fake.docx"
    path.write_bytes(b"this is not a zip archive")
    with pytest.raises(AppError) as excinfo:
        inspect_document(str(path))
    assert excinfo.value.code == "E_DOCX_NOT_ZIP"


def test_zip_without_media_raises_structured_error(tmp_path):
    import zipfile

    path = tmp_path / "empty.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", "<w:document/>")
    with pytest.raises(AppError) as excinfo:
        inspect_document(str(path))
    assert excinfo.value.code == "E_DOCX_NO_MEDIA"
```

- [ ] **Step 5: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_docx_order.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend.services.docx_extract'`

- [ ] **Step 6: 实现清单与保序解析**

创建 `backend/services/docx_extract.py`：

```python
"""从 docx 读取媒体清单，并按图片在文档中的插入顺序排序。

替代原 WordImageExtractorApp.get_image_order_from_docx 与 extract_images。
只依赖标准库 zipfile 与 xml.etree.ElementTree。
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET

from backend.errors import (
    E_DOCX_NOT_FOUND,
    E_DOCX_NOT_ZIP,
    E_DOCX_NO_MEDIA,
    AppError,
)
from backend.services.image_ops import sniff_format

MEDIA_PREFIX = "word/media/"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
}
RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"

# 读取文件头所需的字节数，足以判定所有受支持格式
SNIFF_BYTES = 32


def inspect_document(path):
    """返回媒体清单、插入顺序与顺序来源。"""
    if not os.path.isfile(path):
        raise AppError(
            E_DOCX_NOT_FOUND,
            "找不到这个 Word 文件",
            detail="path=%s" % (path,),
            hint="重新选择文件",
            context={"path": path},
        )

    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            media_parts = [n for n in names if n.startswith(MEDIA_PREFIX)]

            rel_mapping = _read_rel_mapping(archive, names)
            order, order_source = _read_document_order(archive, names, rel_mapping)

            media = []
            for part_name in media_parts:
                with archive.open(part_name) as handle:
                    head = handle.read(SNIFF_BYTES)
                    handle.seek(0)
                    size = len(handle.read())
                fmt = sniff_format(head)
                if fmt is None:
                    continue
                media.append(
                    {
                        "partName": part_name,
                        "relId": _rel_id_for(rel_mapping, part_name),
                        "bytes": size,
                        "format": fmt,
                    }
                )

            if not media:
                raise AppError(
                    E_DOCX_NO_MEDIA,
                    "文档中没有嵌入图片",
                    hint="确认文档中确实插入了图片",
                    context={"path": path},
                )

            known = {entry["partName"] for entry in media}
            order = [p for p in order if p in known]
            # 文档中未被引用的媒体也要出现在末尾，避免丢图
            for entry in media:
                if entry["partName"] not in order:
                    order.append(entry["partName"])
                    order_source = "guess"

            return {"media": media, "order": order, "orderSource": order_source}
    except AppError:
        raise
    except zipfile.BadZipFile as exc:
        raise AppError(
            E_DOCX_NOT_ZIP,
            "该文件不是有效的 .docx（.doc 旧格式不支持）",
            detail=str(exc),
            hint="用 Word 打开并另存为 .docx",
            context={"path": path},
        ) from exc


def _read_rel_mapping(archive, names):
    """建立 rId -> 媒体部件路径 的映射。"""
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in names:
        return {}
    with archive.open(rels_path) as handle:
        root = ET.fromstring(handle.read())
    mapping = {}
    for rel in root.findall(RELS_NS):
        target = rel.get("Target") or ""
        if target.startswith("media/"):
            mapping[rel.get("Id")] = "word/" + target
    return mapping


def _rel_id_for(rel_mapping, part_name):
    for rel_id, target in rel_mapping.items():
        if target == part_name:
            return rel_id
    return None


def _order_from_xml(archive, names, rel_mapping):
    """第一优先级：解析 document.xml，按出现次序收集 drawing 与 pict 引用。"""
    doc_path = "word/document.xml"
    if doc_path not in names or not rel_mapping:
        return None
    with archive.open(doc_path) as handle:
        root = ET.fromstring(handle.read())

    order = []
    embed_attr = "{%s}embed" % NS["r"]
    id_attr = "{%s}id" % NS["r"]

    for node in root.iter():
        if node.tag == "{%s}blip" % NS["a"]:
            _append_once(order, rel_mapping, node.get(embed_attr))
        elif node.tag == "{%s}imagedata" % NS["v"]:
            _append_once(order, rel_mapping, node.get(id_attr))

    return order or None


def _append_once(order, rel_mapping, rel_id):
    if rel_id and rel_id in rel_mapping:
        part = rel_mapping[rel_id]
        if part not in order:
            order.append(part)


def _order_from_regex(archive, rel_mapping):
    """第二优先级：正则扫描 document.xml 中的 r:embed 与 r:id。"""
    try:
        with archive.open("word/document.xml") as handle:
            xml_text = handle.read().decode("utf-8")
    except (KeyError, UnicodeDecodeError):
        return None

    refs = re.findall(r'<a:blip[^>]*r:embed="([^"]+)"', xml_text)
    refs += re.findall(r'<v:imagedata[^>]*r:id="([^"]+)"', xml_text)

    order = []
    for rel_id in refs:
        _append_once(order, rel_mapping, rel_id)
    return order or None


def _order_from_filenames(archive, names):
    """第三优先级：按 media 文件名中的数字排序。"""
    media = sorted(n for n in names if n.startswith(MEDIA_PREFIX))

    def sort_key(part_name):
        digits = re.findall(r"\d+", os.path.basename(part_name))
        return int(digits[0]) if digits else 0

    media.sort(key=sort_key)
    return media


def _read_document_order(archive, names, rel_mapping):
    """返回 (order, order_source)，三级回退，并如实上报来源。"""
    order = _order_from_xml(archive, names, rel_mapping)
    if order:
        return order, "xml"

    order = _order_from_regex(archive, rel_mapping)
    if order:
        return order, "regex"

    order = _order_from_filenames(archive, names)
    if order:
        return order, "filename"

    return [], "guess"
```

- [ ] **Step 7: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_docx_order.py -v`
Expected: 8 passed

- [ ] **Step 8: 提交**

```bash
git add backend/services/docx_extract.py backend/tests/test_docx_order.py backend/tests/fixtures/ .gitignore
git commit -m "feat: 实现 docx 媒体清单与三级保序解析，含可复现样例"
```

---

### Task 7: 实现图片提取写盘（不再清空输出目录）

**Files:**
- Modify: `backend/services/docx_extract.py`（追加）
- Test: `backend/tests/test_extract_write.py`

**Interfaces:**
- Consumes: `inspect_document`（Task 6）
- Produces: `extract_images(path: str, out_dir: str, pattern: str = "{index:03d}", start_index: int = 1, on_conflict: str = "autoSuffix", on_progress=None, should_cancel=None) -> dict`，返回 `{"written": [filename...], "skipped": [{"partName","reason"}...], "renamed": [{"from","to"}...], "orderSource": str, "cancelled": bool}`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_extract_write.py`：

```python
import os

import pytest

from backend.errors import AppError
from backend.services.docx_extract import extract_images

FIXTURE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


def _fixture(name):
    return os.path.join(FIXTURE_DIR, name)


def test_extract_writes_numbered_files_in_document_order(tmp_path):
    result = extract_images(_fixture("reordered.docx"), str(tmp_path))
    assert result["written"] == ["001.png", "002.png", "003.png"]
    assert os.path.isfile(tmp_path / "001.png")


def test_extract_preserves_existing_unrelated_files(tmp_path):
    """回归：旧实现会清空输出目录，这里必须保留用户自己的文件。"""
    keep = tmp_path / "my_notes.txt"
    keep.write_text("keep me", encoding="utf-8")

    extract_images(_fixture("plain.docx"), str(tmp_path))

    assert keep.exists()
    assert keep.read_text(encoding="utf-8") == "keep me"


def test_existing_target_file_is_not_overwritten_by_default(tmp_path):
    existing = tmp_path / "001.png"
    existing.write_bytes(b"sentinel")

    result = extract_images(_fixture("plain.docx"), str(tmp_path))

    assert existing.read_bytes() == b"sentinel"
    assert "001.png" not in result["written"]
    assert any(item["from"] == "001.png" for item in result["renamed"])


def test_custom_pattern_and_start_index(tmp_path):
    result = extract_images(
        _fixture("plain.docx"), str(tmp_path), pattern="pic-{index}", start_index=5
    )
    assert result["written"] == ["pic-5.png", "pic-6.png", "pic-7.png"]


def test_extract_creates_missing_output_directory(tmp_path):
    target = tmp_path / "nested" / "deep"
    result = extract_images(_fixture("plain.docx"), str(target))
    assert os.path.isdir(target)
    assert len(result["written"]) == 3


def test_extract_reports_order_source(tmp_path):
    result = extract_images(_fixture("reordered.docx"), str(tmp_path))
    assert result["orderSource"] == "xml"


def test_progress_callback_is_invoked_for_each_file(tmp_path):
    seen = []
    extract_images(
        _fixture("plain.docx"),
        str(tmp_path),
        on_progress=lambda done, total: seen.append((done, total)),
    )
    assert seen[-1] == (3, 3)
    assert len(seen) == 3


def test_cancellation_stops_after_current_file(tmp_path):
    """收到取消后必须写完当前文件再退出，不留半成品。"""
    calls = {"n": 0}

    def should_cancel():
        calls["n"] += 1
        return calls["n"] > 1

    result = extract_images(
        _fixture("plain.docx"), str(tmp_path), should_cancel=should_cancel
    )
    assert result["cancelled"] is True
    assert len(result["written"]) == 1
    assert os.path.isfile(tmp_path / result["written"][0])


def test_unwritable_output_directory_raises_structured_error():
    with pytest.raises(AppError) as excinfo:
        extract_images(_fixture("plain.docx"), "Z:\\definitely\\not\\writable")
    assert excinfo.value.code in ("E_OUTDIR_NOT_WRITABLE", "E_INTERNAL")
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_extract_write.py -v`
Expected: FAIL — `ImportError: cannot import name 'extract_images'`

- [ ] **Step 3: 实现提取写盘**

在 `backend/services/docx_extract.py` 顶部补 `import shutil`，并把 errors 导入改为：

```python
from backend.errors import (
    E_DOCX_NOT_FOUND,
    E_DOCX_NOT_ZIP,
    E_DOCX_NO_MEDIA,
    E_NAME_EXISTS,
    E_OUTDIR_NOT_WRITABLE,
    AppError,
)
```

在文件末尾追加：

```python
def _ensure_dir_writable(out_dir):
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as exc:
        raise AppError(
            E_OUTDIR_NOT_WRITABLE,
            "输出目录不可写（可能被占用或无权限）",
            detail=str(exc),
            hint="更换一个输出目录",
            context={"outDir": str(out_dir)},
        ) from exc

    probe = os.path.join(out_dir, ".write_probe")
    try:
        with open(probe, "w", encoding="utf-8") as handle:
            handle.write("")
        os.remove(probe)
    except OSError as exc:
        raise AppError(
            E_OUTDIR_NOT_WRITABLE,
            "输出目录不可写（可能被占用或无权限）",
            detail=str(exc),
            hint="更换一个输出目录",
            context={"outDir": str(out_dir)},
        ) from exc


def _unique_name(out_dir, base, extension):
    """在目录内寻找未被占用的文件名，返回 (filename, 是否发生改名)。"""
    candidate = base + extension
    if not os.path.exists(os.path.join(out_dir, candidate)):
        return candidate, False
    counter = 1
    while True:
        candidate = "%s_%d%s" % (base, counter, extension)
        if not os.path.exists(os.path.join(out_dir, candidate)):
            return candidate, True
        counter += 1


def extract_images(
    path,
    out_dir,
    pattern="{index:03d}",
    start_index=1,
    on_conflict="autoSuffix",
    on_progress=None,
    should_cancel=None,
):
    """按文档顺序提取图片到输出目录。

    与旧实现的两处关键差异：
      1. 绝不删除输出目录中已存在的文件；
      2. 遇到同名文件默认加序号，而不是覆盖。
    """
    if on_conflict not in ("autoSuffix", "overwrite", "reject"):
        raise AppError(
            E_NAME_EXISTS,
            "未知的冲突策略：%s" % (on_conflict,),
            hint="取值范围为 autoSuffix / overwrite / reject",
        )

    _ensure_dir_writable(out_dir)
    inspected = inspect_document(path)
    order = inspected["order"]
    by_part = {entry["partName"]: entry for entry in inspected["media"]}

    written = []
    skipped = []
    renamed = []
    cancelled = False
    total = len(order)

    with zipfile.ZipFile(path) as archive:
        for offset, part_name in enumerate(order):
            index = start_index + offset
            entry = by_part.get(part_name)
            if entry is None:
                skipped.append({"partName": part_name, "reason": "unsupported-format"})
                _report(on_progress, offset + 1, total)
                continue

            extension = EXTENSION_BY_FORMAT[entry["format"]]
            base = pattern.format(index=index)

            if on_conflict == "overwrite":
                filename = base + extension
                changed = os.path.exists(os.path.join(out_dir, filename))
            elif on_conflict == "reject":
                filename = base + extension
                if os.path.exists(os.path.join(out_dir, filename)):
                    skipped.append({"partName": part_name, "reason": "name-exists"})
                    _report(on_progress, offset + 1, total)
                    continue
                changed = False
            else:
                filename, changed = _unique_name(out_dir, base, extension)

            with archive.open(part_name) as source:
                data = source.read()

            target = os.path.join(out_dir, filename)
            with open(target, "wb") as handle:
                handle.write(data)

            written.append(filename)
            if changed:
                renamed.append({"from": base + extension, "to": filename})

            _report(on_progress, offset + 1, total)

            if should_cancel is not None and should_cancel():
                cancelled = True
                break

    return {
        "written": written,
        "skipped": skipped,
        "renamed": renamed,
        "orderSource": inspected["orderSource"],
        "cancelled": cancelled,
    }


def _report(on_progress, done, total):
    if on_progress is not None:
        on_progress(done, total)
```

同时把 `from backend.services.image_ops import sniff_format` 改为：

```python
from backend.services.image_ops import EXTENSION_BY_FORMAT, sniff_format
```

- [ ] **Step 4: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_extract_write.py -v`
Expected: 9 passed

- [ ] **Step 5: 提交**

```bash
git add backend/services/docx_extract.py backend/tests/test_extract_write.py
git commit -m "fix: 提取图片不再清空输出目录，冲突改为安全命名"
```

---

### Task 8: 实现目录扫描与排序

**Files:**
- Create: `backend/services/library.py`
- Test: `backend/tests/test_library_list.py`
- Test fixtures: `backend/tests/fixtures/library/`（运行时由测试创建，不入库）

**Interfaces:**
- Consumes: `sniff_format`（Task 3）、`AppError`（Task 1）
- Produces: `list_images(folder: str) -> list[dict]`，每项 `{"name","size","mtime","width","height"}`；`natural_sort_key(name: str) -> tuple`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_library_list.py`：

```python
import io
import os

import pytest
from PIL import Image

from backend.errors import AppError
from backend.services.library import list_images, natural_sort_key


def _png(tmp_path, name, size=(4, 4)):
    path = tmp_path / name
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.new("RGB", size, (5, 10, 15)).save(path)
    return path


def test_natural_sort_key_handles_pure_text_names():
    """回归：旧实现用 int(''.join(digits)) 排序，纯文字名会退化为 0。"""
    assert natural_sort_key("cover.png") == natural_sort_key("banner.png")
    assert natural_sort_key("cover.png")[1] == "cover.png"


def test_natural_sort_key_orders_numerically_not_lexically():
    keys = [natural_sort_key(n) for n in ("2.png", "10.png", "1.png")]
    assert [k for k in sorted(keys)] == [
        natural_sort_key("1.png"),
        natural_sort_key("2.png"),
        natural_sort_key("10.png"),
    ]


def test_list_images_returns_entry_fields(tmp_path):
    _png(tmp_path, "001.png", size=(7, 9))
    entries = list_images(str(tmp_path))
    assert len(entries) == 1
    entry = entries[0]
    assert entry["name"] == "001.png"
    assert entry["width"] == 7
    assert entry["height"] == 9
    assert entry["size"] > 0
    assert entry["mtime"] > 0


def test_list_images_sorts_by_leading_number(tmp_path):
    for name in ("010.png", "002.png", "001.png"):
        _png(tmp_path, name)
    names = [entry["name"] for entry in list_images(str(tmp_path))]
    assert names == ["001.png", "002.png", "010.png"]


def test_list_images_ignores_non_image_files(tmp_path):
    _png(tmp_path, "001.png")
    (tmp_path / "notes.txt").write_text("hello", encoding="utf-8")
    (tmp_path / "data.json").write_text("{}", encoding="utf-8")
    names = [entry["name"] for entry in list_images(str(tmp_path))]
    assert names == ["001.png"]


def test_list_images_ignores_subdirectories(tmp_path):
    _png(tmp_path, "001.png")
    (tmp_path / "sub").mkdir()
    _png(tmp_path, "sub" / "002.png")
    assert [e["name"] for e in list_images(str(tmp_path))] == ["001.png"]


def test_list_images_skips_corrupt_image_without_failing(tmp_path):
    _png(tmp_path, "001.png")
    (tmp_path / "002.png").write_bytes(b"corrupt bytes here")
    names = [entry["name"] for entry in list_images(str(tmp_path))]
    assert names == ["001.png"]


def test_list_images_on_missing_folder_raises(tmp_path):
    with pytest.raises(AppError) as excinfo:
        list_images(str(tmp_path / "nope"))
    assert excinfo.value.code == "E_DIR_NOT_FOUND"


def test_list_images_on_empty_folder_returns_empty_list(tmp_path):
    assert list_images(str(tmp_path)) == []


def test_list_images_mixed_numeric_and_text_names(tmp_path):
    _png(tmp_path, "cover.png")
    _png(tmp_path, "001.png")
    _png(tmp_path, "intro.png")
    names = [entry["name"] for entry in list_images(str(tmp_path))]
    # 数字名在前按数值排，文字名在后按字典序排
    assert names == ["001.png", "cover.png", "intro.png"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_library_list.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend.services.library'`

- [ ] **Step 3: 实现扫描与排序**

创建 `backend/services/library.py`：

```python
"""目录扫描、排序与文件重命名。

替代原 WordImageExtractorApp 的 load_image_files / save_rename / open_image_folder。
"""

import os
import re

from backend.errors import AppError
from backend.services.image_ops import EXTENSION_BY_FORMAT, sniff_format, open_image

E_DIR_NOT_FOUND = "E_DIR_NOT_FOUND"

# 仅读取文件头即可判定格式
SNIFF_BYTES = 32
_DIGITS = re.compile(r"(\d+)")


def natural_sort_key(name):
    """排序键：数字名按数值排在前，纯文字名按字典序排在后。

    返回 (分组, 数值或字符串, 完整名)。分组 0 表示含前导数字的文件名。
    第三项保证完全同键时排序稳定。
    """
    match = _DIGITS.search(name)
    if match and name.index(match.group(1)) <= 2:
        return (0, int(match.group(1)), name)
    return (1, name.lower(), name)


def list_images(folder):
    """列出目录中的图片。忽略非图片、子目录与损坏文件，不抛异常。"""
    if not os.path.isdir(folder):
        raise AppError(
            E_DIR_NOT_FOUND,
            "图片文件夹不存在",
            detail="path=%s" % (folder,),
            hint="重新选择文件夹",
            context={"dir": str(folder)},
        )

    entries = []
    with os.scandir(folder) as iterator:
        for item in iterator:
            if not item.is_file():
                continue
            try:
                head = _read_head(item.path)
            except OSError:
                continue
            if sniff_format(head) is None:
                continue
            try:
                image = open_image(item.path)
            except AppError:
                # 头信息像图片但无法解码，跳过而不是让整批失败
                continue
            stat = item.stat()
            entries.append(
                {
                    "name": item.name,
                    "size": stat.st_size,
                    "mtime": stat.st_mtime,
                    "width": image.width,
                    "height": image.height,
                }
            )

    entries.sort(key=lambda entry: natural_sort_key(entry["name"]))
    return entries


def _read_head(path):
    with open(path, "rb") as handle:
        return handle.read(SNIFF_BYTES)
```

- [ ] **Step 4: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_library_list.py -v`
Expected: 10 passed

- [ ] **Step 5: 提交**

```bash
git add backend/services/library.py backend/tests/test_library_list.py
git commit -m "feat: 实现图片目录扫描与自然排序，修正纯文字文件名排序"
```

---

### Task 9: 实现重命名与批量重命名两阶段提交

**Files:**
- Modify: `backend/services/library.py`（追加）
- Test: `backend/tests/test_library_rename.py`

**Interfaces:**
- Consumes: `list_images`（Task 8）、`AppError`、`ILLEGAL_NAME_CHARS`（Task 1）
- Produces: `validate_name(name: str) -> str`；`rename(folder, from_name, to_name, on_conflict="reject") -> dict`；`batch_rename(folder, entries: list[dict], on_conflict="reject") -> dict`，`entries` 形如 `[{"from": "001.png", "to": "封面"}]`，返回 `{"renamed": [...], "conflicts": [{"from","to","reason"}]}`

- [ ] **Step 1: 写失败的测试**

创建 `backend/tests/test_library_rename.py`：

```python
import pytest
from PIL import Image

from backend.errors import AppError
from backend.services.library import (
    batch_rename,
    list_images,
    rename,
    validate_name,
)


def _png(tmp_path, name, size=(4, 4)):
    path = tmp_path / name
    Image.new("RGB", size, (5, 10, 15)).save(path)
    return path


@pytest.mark.parametrize("bad", ['a\\b', "a/b", "a:b", "a*b", "a?b", 'a"b', "a<b", "a>b", "a|b"])
def test_validate_name_rejects_illegal_characters(bad):
    with pytest.raises(AppError) as excinfo:
        validate_name(bad)
    assert excinfo.value.code == "E_NAME_INVALID"


@pytest.mark.parametrize("bad", ["", "   ", "name.", "name "])
def test_validate_name_rejects_empty_and_trailing_forms(bad):
    with pytest.raises(AppError) as excinfo:
        validate_name(bad)
    assert excinfo.value.code == "E_NAME_INVALID"


def test_validate_name_accepts_chinese_and_spaces():
    assert validate_name("  封面 图 1  ") == "封面 图 1"


def test_rename_preserves_extension_and_content(tmp_path):
    _png(tmp_path, "001.png")
    before = (tmp_path / "001.png").read_bytes()

    result = rename(str(tmp_path), "001.png", "封面")

    assert result["path"].endswith("封面.png")
    assert (tmp_path / "封面.png").read_bytes() == before
    assert not (tmp_path / "001.png").exists()


def test_rename_accepts_name_with_extension(tmp_path):
    _png(tmp_path, "001.png")
    result = rename(str(tmp_path), "001.png", "封面.png")
    assert result["path"].endswith("封面.png")


def test_rename_default_rejects_existing_target(tmp_path):
    _png(tmp_path, "001.png")
    _png(tmp_path, "002.png")
    with pytest.raises(AppError) as excinfo:
        rename(str(tmp_path), "001.png", "002")
    assert excinfo.value.code == "E_NAME_EXISTS"
    assert (tmp_path / "002.png").exists()


def test_rename_with_auto_suffix_finds_a_free_name(tmp_path):
    _png(tmp_path, "001.png")
    _png(tmp_path, "002.png")
    result = rename(str(tmp_path), "001.png", "002", on_conflict="autoSuffix")
    assert result["path"].endswith("002_1.png")


def test_rename_to_same_name_is_a_noop(tmp_path):
    _png(tmp_path, "001.png")
    result = rename(str(tmp_path), "001.png", "001")
    assert result["path"].endswith("001.png")


def test_rename_missing_source_raises(tmp_path):
    with pytest.raises(AppError) as excinfo:
        rename(str(tmp_path), "ghost.png", "new")
    assert excinfo.value.code == "E_FILE_NOT_FOUND"


def test_batch_rename_applies_all_entries(tmp_path):
    for name in ("001.png", "002.png", "003.png"):
        _png(tmp_path, name)
    result = batch_rename(
        str(tmp_path),
        [
            {"from": "001.png", "to": "封面"},
            {"from": "002.png", "to": "目录"},
        ],
    )
    assert len(result["renamed"]) == 2
    assert result["conflicts"] == []
    assert {e["name"] for e in list_images(str(tmp_path))} == {"封面.png", "目录.png", "003.png"}


def test_batch_rename_handles_swap_without_losing_files(tmp_path):
    """001 与 002 互换名字：两阶段提交必须能安全完成。"""
    _png(tmp_path, "001.png", size=(7, 7))
    _png(tmp_path, "002.png", size=(9, 9))
    result = batch_rename(
        str(tmp_path),
        [
            {"from": "001.png", "to": "002.png"},
            {"from": "002.png", "to": "001.png"},
        ],
        on_conflict="overwrite",
    )
    assert result["conflicts"] == []
    from PIL import Image as PILImage

    assert PILImage.open(tmp_path / "001.png").size == (9, 9)
    assert PILImage.open(tmp_path / "002.png").size == (7, 7)


def test_batch_rename_reports_conflicts_and_changes_nothing(tmp_path):
    """任一条目失败时，整批不做任何改动。"""
    _png(tmp_path, "001.png")
    _png(tmp_path, "002.png")
    result = batch_rename(
        str(tmp_path),
        [
            {"from": "001.png", "to": "renamed"},
            {"from": "002.png", "to": "003.png"},
        ],
        on_conflict="reject",
    )
    assert result["renamed"] == []
    assert len(result["conflicts"]) == 1
    assert result["conflicts"][0]["to"] == "003.png"
    # 001 未被改名，因为整批回滚
    assert {e["name"] for e in list_images(str(tmp_path))} == {"001.png", "002.png"}
```

注意：`test_batch_rename_reports_conflicts_and_changes_nothing` 里 `002.png -> 003.png` 并不冲突（003 不存在），因此本测试要断言冲突来源应为**目标名非法或源文件不存在**。请在实现时按下述语义确定该用例：把第二个条目改为 `{"from": "ghost.png", "to": "003.png"}`，使冲突原因明确为 `source-missing`。同步修改断言中的 `conflicts[0]["reason"] == "source-missing"`。

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_library_rename.py -v`
Expected: FAIL — `ImportError: cannot import name 'batch_rename'`

- [ ] **Step 3: 实现校验与单张重命名**

在 `backend/services/library.py` 顶部把导入改为：

```python
import os
import re
import tempfile

from backend.errors import (
    E_NAME_EXISTS,
    E_NAME_INVALID,
    ILLEGAL_NAME_CHARS,
    AppError,
)
from backend.services.image_ops import EXTENSION_BY_FORMAT, open_image, sniff_format
```

并在文件末尾追加：

```python
E_FILE_NOT_FOUND = "E_FILE_NOT_FOUND"
E_TARGET_MISSING = "E_TARGET_MISSING"


def validate_name(name):
    """校验并归一化文件基本名（不含扩展名）。"""
    cleaned = (name or "").strip()
    if not cleaned:
        raise AppError(E_NAME_INVALID, "文件名不能为空", hint="输入一个非空的名称")

    offending = [ch for ch in ILLEGAL_NAME_CHARS if ch in cleaned]
    if offending:
        raise AppError(
            E_NAME_INVALID,
            '文件名含非法字符 %s' % " ".join(offending),
            detail="illegal=%r" % (offending,),
            hint="移除这些字符后重试",
            context={"name": cleaned},
        )

    if cleaned.endswith("."):
        cleaned = cleaned.rstrip(".")
    if not cleaned:
        raise AppError(E_NAME_INVALID, "文件名不能为空", hint="输入一个非空的名称")

    if cleaned != cleaned.rstrip(" "):
        cleaned = cleaned.rstrip(" ")

    reserved = {"CON", "PRN", "AUX", "NUL"}
    reserved |= {"COM%d" % i for i in range(1, 10)}
    reserved |= {"LPT%d" % i for i in range(1, 10)}
    if cleaned.upper().split(".")[0] in reserved:
        raise AppError(
            E_NAME_INVALID,
            "该名称是系统保留名，不能用作文件名",
            detail="name=%s" % cleaned,
            hint="换一个名称",
        )

    return cleaned


def _split_target(to_name, source_name):
    """返回 (基本名, 扩展名)。未指定扩展名时沿用源文件扩展名。"""
    _, source_ext = os.path.splitext(source_name)
    base, ext = os.path.splitext(to_name or "")
    if not ext:
        ext = source_ext
    return base, ext


def _unique_target(folder, base, ext):
    candidate = base + ext
    if not os.path.exists(os.path.join(folder, candidate)):
        return candidate
    counter = 1
    while True:
        candidate = "%s_%d%s" % (base, counter, ext)
        if not os.path.exists(os.path.join(folder, candidate)):
            return candidate
        counter += 1


def rename(folder, from_name, to_name, on_conflict="reject"):
    """重命名单个文件。默认拒绝覆盖已存在的目标。"""
    source_path = os.path.join(folder, from_name)
    if not os.path.isfile(source_path):
        raise AppError(
            E_FILE_NOT_FOUND,
            "源文件不存在",
            detail="path=%s" % source_path,
            hint="刷新列表后重试",
            context={"from": from_name},
        )

    base, ext = _split_target(to_name, from_name)
    base = validate_name(base)

    target_name = base + ext
    if target_name == from_name:
        return {"path": source_path, "name": from_name}

    target_path = os.path.join(folder, target_name)
    if os.path.exists(target_path):
        if on_conflict == "autoSuffix":
            target_name = _unique_target(folder, base, ext)
            target_path = os.path.join(folder, target_name)
        elif on_conflict == "overwrite":
            pass
        else:
            raise AppError(
                E_NAME_EXISTS,
                "目标文件名已存在",
                detail="target=%s" % target_name,
                hint="换一个名称，或选择自动加序号 / 覆盖",
                context={"from": from_name, "to": target_name},
            )

    os.replace(source_path, target_path)
    return {"path": target_path, "name": target_name}


def batch_rename(folder, entries, on_conflict="reject"):
    """两阶段提交的批量重命名。

    第一阶段在内存中规划全部目标名并检测冲突；只有全部条目都可行时，
    第二阶段才真正落盘。任一条目不可行时整批不做任何改动。
    """
    planned = []
    conflicts = []
    claimed = set()

    for entry in entries or []:
        from_name = entry.get("from")
        to_name = entry.get("to")
        source_path = os.path.join(folder, from_name or "")

        if not from_name or not os.path.isfile(source_path):
            conflicts.append({"from": from_name, "to": to_name, "reason": "source-missing"})
            continue

        try:
            base, ext = _split_target(to_name, from_name)
            base = validate_name(base)
        except AppError as exc:
            conflicts.append({"from": from_name, "to": to_name, "reason": exc.code})
            continue

        target_name = base + ext

        if target_name in claimed:
            conflicts.append({"from": from_name, "to": target_name, "reason": "duplicate-target"})
            continue

        if os.path.exists(os.path.join(folder, target_name)):
            other_sources = {
                item.get("from") for item in entries if item.get("from") != from_name
            }
            if target_name in other_sources:
                # 互换或链式改名：目标文件在本次批次中会被移走，不算冲突
                pass
            elif on_conflict == "autoSuffix":
                target_name = _unique_target(folder, base, ext)
            elif on_conflict == "overwrite":
                pass
            else:
                conflicts.append({"from": from_name, "to": target_name, "reason": "name-exists"})
                continue

        if target_name == from_name:
            planned.append({"from": from_name, "to": target_name, "skip": True})
            claimed.add(target_name)
            continue

        claimed.add(target_name)
        planned.append({"from": from_name, "to": target_name, "skip": False})

    if conflicts:
        return {"renamed": [], "conflicts": conflicts}

    return _apply_planned(folder, planned)


def _apply_planned(folder, planned):
    """第二阶段：先把所有源文件移到临时名，再移到最终名，避免互换时互相覆盖。"""
    renamed = []
    staged = []
    temp_dir = None

    try:
        temp_dir = tempfile.mkdtemp(prefix=".rename_stage_", dir=folder)
    except OSError as exc:
        raise AppError(
            E_TARGET_MISSING,
            "无法在目标目录创建临时暂存区",
            detail=str(exc),
            hint="确认目录可写且未被占用",
            context={"dir": str(folder)},
        ) from exc

    try:
        for index, item in enumerate(planned):
            if item["skip"]:
                continue
            staged_path = os.path.join(temp_dir, "%d" % index)
            os.replace(os.path.join(folder, item["from"]), staged_path)
            staged.append((staged_path, item))

        for staged_path, item in staged:
            final_path = os.path.join(folder, item["to"])
            os.replace(staged_path, final_path)
            renamed.append({"from": item["from"], "to": item["to"]})
    finally:
        if temp_dir and os.path.isdir(temp_dir):
            for leftover in os.listdir(temp_dir):
                try:
                    os.remove(os.path.join(temp_dir, leftover))
                except OSError:
                    pass
            try:
                os.rmdir(temp_dir)
            except OSError:
                pass

    return {"renamed": renamed, "conflicts": []}
```

- [ ] **Step 4: 修正冲突测试用例**

按 Step 1 的说明，把 `test_batch_rename_reports_conflicts_and_changes_nothing` 的第二个条目改为：

```python
            {"from": "ghost.png", "to": "003.png"},
```

并把断言改为：

```python
    assert result["conflicts"][0]["reason"] == "source-missing"
```

- [ ] **Step 5: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_library_rename.py -v`
Expected: 全部通过（含 12 个参数化用例展开后的总数）

- [ ] **Step 6: 运行全部后端测试**

Run: `venv\Scripts\python.exe -m pytest backend/tests -v`
Expected: 全部通过，无失败

- [ ] **Step 7: 提交**

```bash
git add backend/services/library.py backend/tests/test_library_rename.py
git commit -m "feat: 实现重命名与两阶段批量重命名，冲突不再留半成品"
```

---

### Task 10: 实现 RPC 分发器与自检入口

**Files:**
- Create: `backend/cache.py`
- Create: `backend/rpc.py`
- Modify: `backend/selftest.py`（替换占位）
- Test: `backend/tests/test_rpc.py`
- Test: `backend/tests/test_cache.py`

**Interfaces:**
- Consumes: 全部服务函数（Task 3–9）
- Produces:
  - `RenderCache(root_dir: str, max_bytes: int = 268435456)`，方法 `.put(source_path, ops, image) -> str`（返回 cacheKey）、`.get_path(cache_key: str) -> str | None`、`.is_valid_key(cache_key: str) -> bool`
  - `dispatch(method: str, params: dict, ctx: dict) -> dict`
  - `METHODS: dict[str, callable]`
  - `run_selftest() -> int`

- [ ] **Step 1: 写渲染缓存的失败测试**

创建 `backend/tests/test_cache.py`：

```python
import os

from PIL import Image

from backend.cache import RenderCache
from backend.services.image_ops import apply_ops


def _make(path, size=(20, 10), color=(200, 100, 50)):
    Image.new("RGB", size, color).save(path)
    return str(path)


def test_put_returns_hex_key_and_writes_a_file(tmp_path):
    source = _make(tmp_path / "a.png")
    cache = RenderCache(str(tmp_path / "cache"))
    key = cache.put(source, [], apply_ops(source, []))
    assert len(key) == 64
    assert os.path.isfile(cache.get_path(key))


def test_same_ops_hit_the_same_key(tmp_path):
    source = _make(tmp_path / "a.png")
    cache = RenderCache(str(tmp_path / "cache"))
    ops = [{"type": "rotate", "angle": 90}]
    first = cache.put(source, ops, apply_ops(source, ops))
    second = cache.put(source, ops, apply_ops(source, ops))
    assert first == second


def test_different_max_edge_produces_different_key(tmp_path):
    source = _make(tmp_path / "a.png")
    cache = RenderCache(str(tmp_path / "cache"))
    image = apply_ops(source, [])
    assert cache.put(source, [], image, max_edge=1000) != cache.put(source, [], image, max_edge=500)


def test_is_valid_key_rejects_traversal_attempt(tmp_path):
    cache = RenderCache(str(tmp_path / "cache"))
    assert cache.is_valid_key("..\\..\\secret") is False
    assert cache.is_valid_key("not-a-hex-key") is False
    assert cache.is_valid_key("a" * 64) is True


def test_get_path_returns_none_for_missing_key(tmp_path):
    cache = RenderCache(str(tmp_path / "cache"))
    assert cache.get_path("b" * 64) is None


def test_lru_eviction_keeps_total_under_limit(tmp_path):
    cache = RenderCache(str(tmp_path / "cache"), max_bytes=1)
    first = cache.put(_make(tmp_path / "a.png"), [], Image.new("RGB", (8, 8)))
    cache.put(_make(tmp_path / "b.png"), [], Image.new("RGB", (8, 8)))
    assert cache.get_path(first) is None
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_cache.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend.cache'`

- [ ] **Step 3: 实现渲染缓存**

创建 `backend/cache.py`：

```python
"""预览渲染缓存：LRU 淘汰 + 磁盘落盘。

key 由「源文件指纹 + 操作栈 + maxEdge」决定；几何部分与 image_ops.ops_signature 同源。
"""

import hashlib
import os
import re
import shutil
import tempfile

from backend.services.image_ops import ops_signature

DEFAULT_MAX_BYTES = 256 * 1024 * 1024
_KEY_RE = re.compile(r"^[0-9a-f]{64}$")


class RenderCache:
    def __init__(self, root_dir, max_bytes=DEFAULT_MAX_BYTES):
        self.root_dir = root_dir
        self.max_bytes = max_bytes
        os.makedirs(self.root_dir, exist_ok=True)

    def is_valid_key(self, cache_key):
        """key 必须是 64 位十六进制，杜绝路径穿越。"""
        return bool(isinstance(cache_key, str) and _KEY_RE.match(cache_key))

    def get_path(self, cache_key):
        if not self.is_valid_key(cache_key):
            return None
        path = os.path.join(self.root_dir, cache_key + ".png")
        return path if os.path.isfile(path) else None

    def put(self, source_path, ops, image, max_edge=None):
        """maxEdge 必须参与键计算：缩略图（128）与预览（2400）是两份不同产物。"""
        cache_key = self.key_for(source_path, ops, max_edge)
        target = os.path.join(self.root_dir, cache_key + ".png")
        if not os.path.isfile(target):
            self._write_atomic(image, target)
        self._evict_if_needed(protect=cache_key)
        return cache_key

    @staticmethod
    def key_for(source_path, ops, max_edge=None):
        """对外公开的键计算，使测试与调用方都能独立推出同一个键。"""
        signature = ops_signature(source_path, ops)
        if max_edge is None:
            return signature
        payload = "%s|edge=%s" % (signature, max_edge)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def _write_atomic(self, image, target):
        handle, temp_path = tempfile.mkstemp(suffix=".png", dir=self.root_dir)
        os.close(handle)
        try:
            image.convert("RGB").save(temp_path, format="PNG")
            os.replace(temp_path, target)
        except OSError:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise

    def _evict_if_needed(self, protect=None):
        entries = []
        total = 0
        for name in os.listdir(self.root_dir):
            path = os.path.join(self.root_dir, name)
            if not os.path.isfile(path):
                continue
            stat = os.stat(path)
            total += stat.st_size
            entries.append((stat.st_atime, path, stat.st_size, name))

        if total <= self.max_bytes:
            return

        entries.sort()
        for _, path, size, name in entries:
            if total <= self.max_bytes:
                break
            if protect and name == protect + ".png":
                continue
            try:
                os.remove(path)
                total -= size
            except OSError:
                continue

    def clear(self):
        shutil.rmtree(self.root_dir, ignore_errors=True)
        os.makedirs(self.root_dir, exist_ok=True)
```

- [ ] **Step 4: 运行缓存测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_cache.py -v`
Expected: 6 passed

- [ ] **Step 5: 写 RPC 分发的失败测试**

创建 `backend/tests/test_rpc.py`：

```python
import os

import pytest
from PIL import Image

from backend.errors import AppError
from backend.rpc import METHODS, dispatch
from backend.security import PathGuard

FIXTURE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


def _ctx(tmp_path):
    return {
        "paths": PathGuard([tmp_path, FIXTURE_DIR]),
        "cache": None,
        "root": str(tmp_path),
    }


def _png(tmp_path, name, size=(8, 8)):
    path = tmp_path / name
    Image.new("RGB", size, (5, 10, 15)).save(path)
    return path


def test_ping_reports_versions(tmp_path):
    result = dispatch("app.ping", {}, _ctx(tmp_path))
    assert "version" in result
    assert result["pillow"].count(".") >= 1
    assert result["python"]


def test_unknown_method_is_rejected(tmp_path):
    with pytest.raises(AppError) as excinfo:
        dispatch("nope.nothing", {}, _ctx(tmp_path))
    assert excinfo.value.code == "E_METHOD_NOT_FOUND"


def test_methods_table_is_static_and_non_empty():
    assert isinstance(METHODS, dict)
    assert "app.ping" in METHODS
    for name in METHODS:
        assert isinstance(name, str)
        assert "." in name


def test_docx_inspect_requires_allowed_path(tmp_path):
    outside = tmp_path.parent / "outside.docx"
    with pytest.raises(AppError) as excinfo:
        dispatch("docx.inspect", {"path": str(outside)}, _ctx(tmp_path))
    assert excinfo.value.code == "E_PATH_DENIED"


def test_docx_inspect_works_for_allowed_path(tmp_path):
    result = dispatch(
        "docx.inspect", {"path": os.path.join(FIXTURE_DIR, "plain.docx")}, _ctx(tmp_path)
    )
    assert len(result["order"]) == 3


def test_library_list_returns_entries(tmp_path):
    _png(tmp_path, "001.png")
    result = dispatch("library.list", {"dir": str(tmp_path)}, _ctx(tmp_path))
    assert [entry["name"] for entry in result["images"]] == ["001.png"]


def test_library_rename_validates_and_renames(tmp_path):
    _png(tmp_path, "001.png")
    result = dispatch(
        "library.rename",
        {"dir": str(tmp_path), "from": "001.png", "to": "封面"},
        _ctx(tmp_path),
    )
    assert result["name"] == "封面.png"
    assert os.path.isfile(tmp_path / "封面.png")


def test_image_signature_matches_direct_call(tmp_path):
    path = _png(tmp_path, "001.png")
    from backend.services.image_ops import ops_signature

    result = dispatch("image.signature", {"path": str(path), "ops": []}, _ctx(tmp_path))
    assert result["signature"] == ops_signature(str(path), [])


def test_image_save_writes_atomically_and_reports_bytes(tmp_path):
    path = _png(tmp_path, "001.png", size=(20, 10))
    target = tmp_path / "out.png"
    result = dispatch(
        "image.save",
        {
            "path": str(path),
            "targetPath": str(target),
            "ops": [{"type": "rotate", "angle": 90}],
            "format": "png",
        },
        _ctx(tmp_path),
    )
    assert result["bytes"] > 0
    assert Image.open(target).size == (10, 20)


def test_image_save_converts_transparency_for_jpeg(tmp_path):
    path = tmp_path / "rgba.png"
    Image.new("RGBA", (8, 8), (255, 0, 0, 0)).save(path)
    target = tmp_path / "out.jpg"
    dispatch(
        "image.save",
        {"path": str(path), "targetPath": str(target), "ops": [], "format": "jpeg"},
        _ctx(tmp_path),
    )
    assert Image.open(target).mode == "RGB"


def test_image_batch_save_reports_results(tmp_path):
    _png(tmp_path, "001.png")
    _png(tmp_path, "002.png")
    result = dispatch(
        "image.batchSave",
        {
            "items": [
                {"path": str(tmp_path / "001.png"), "ops": []},
                {"path": str(tmp_path / "002.png"), "ops": [{"type": "rotate", "angle": 90}]},
            ],
            "format": "png",
        },
        _ctx(tmp_path),
    )
    assert result["savedCount"] == 2
    assert result["failed"] == []
```

- [ ] **Step 6: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_rpc.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'backend.rpc'`

- [ ] **Step 7: 实现 RPC 分发器**

创建 `backend/rpc.py`：

```python
"""静态方法表分发。不做任何动态导入或 eval，方法名必须精确匹配。"""

import os
import platform
import sys
import tempfile

from PIL import __version__ as PILLOW_VERSION

from backend.cache import RenderCache
from backend.errors import E_INTERNAL, AppError, E_INVALID_OP
from backend.services import docx_extract, image_ops, library

E_METHOD_NOT_FOUND = "E_METHOD_NOT_FOUND"
APP_VERSION = "2.0.0"


def _require(params, key):
    value = params.get(key)
    if value is None or value == "":
        raise AppError(
            E_INVALID_OP,
            "缺少必需参数：%s" % key,
            hint="检查界面调用是否完整",
            context={"params": list(params.keys())},
        )
    return value


def _ping(params, ctx):
    return {
        "version": APP_VERSION,
        "python": platform.python_version(),
        "pillow": PILLOW_VERSION,
        "platform": sys.platform,
    }


def _docx_inspect(params, ctx):
    path = ctx["paths"].allow(_require(params, "path"))
    return docx_extract.inspect_document(str(path))


def _docx_extract(params, ctx):
    docx_path = ctx["paths"].allow(_require(params, "path"))
    out_dir = ctx["paths"].allow_root(_require(params, "outDir"))
    return docx_extract.extract_images(
        str(docx_path),
        str(out_dir),
        pattern=params.get("pattern") or "{index:03d}",
        start_index=int(params.get("startIndex") or 1),
        on_conflict=params.get("onConflict") or "autoSuffix",
    )


def _library_list(params, ctx):
    folder = ctx["paths"].allow(_require(params, "dir"))
    return {"images": library.list_images(str(folder))}


def _library_rename(params, ctx):
    folder = ctx["paths"].allow(_require(params, "dir"))
    return library.rename(
        str(folder),
        _require(params, "from"),
        _require(params, "to"),
        on_conflict=params.get("onConflict") or "reject",
    )


def _library_batch_rename(params, ctx):
    folder = ctx["paths"].allow(_require(params, "dir"))
    return library.batch_rename(
        str(folder),
        params.get("entries") or [],
        on_conflict=params.get("onConflict") or "reject",
    )


def _image_signature(params, ctx):
    path = ctx["paths"].allow(_require(params, "path"))
    return {"signature": image_ops.ops_signature(str(path), params.get("ops") or [])}


def _image_preview(params, ctx):
    path = ctx["paths"].allow(_require(params, "path"))
    ops = params.get("ops") or []
    max_edge = params.get("maxEdge")
    image = image_ops.apply_ops(str(path), ops)
    if max_edge:
        image = image.copy()
        image.thumbnail((int(max_edge), int(max_edge)))
    cache = ctx.get("cache") or _default_cache(ctx)
    cache_key = cache.put(str(path), ops, image, max_edge=max_edge)
    return {"cacheKey": cache_key}


def _image_save(params, ctx):
    path = ctx["paths"].allow(_require(params, "path"))
    target = ctx["paths"].allow_root(_require(params, "targetPath"))
    return image_ops.save(
        str(path),
        str(target),
        params.get("ops") or [],
        fmt=params.get("format") or "png",
        quality=params.get("quality"),
    )


def _image_batch_save(params, ctx):
    items = params.get("items") or []
    fmt = params.get("format") or "png"
    quality = params.get("quality")
    out_dir = params.get("outDir")
    if out_dir:
        out_dir = str(ctx["paths"].allow_root(out_dir))

    saved = []
    failed = []
    for item in items:
        try:
            source = str(ctx["paths"].allow(_require(item, "path")))
            if out_dir:
                base = os.path.splitext(os.path.basename(source))[0]
                target = os.path.join(out_dir, "%s.%s" % (base, _ext_for(fmt)))
                ctx["paths"].allow_root(target)
            else:
                target = source
            saved.append(
                image_ops.save(
                    source,
                    target,
                    item.get("ops") or [],
                    fmt=item.get("format") or fmt,
                    quality=item.get("quality", quality),
                )
            )
        except AppError as exc:
            failed.append({"path": item.get("path"), "error": exc.to_dict()})

    return {
        "saved": saved,
        "savedCount": len(saved),
        "failed": failed,
        "failedCount": len(failed),
    }


def _system_reveal(params, ctx):
    path = ctx["paths"].allow(_require(params, "path"))
    target = str(path if path.is_dir() else path.parent)
    if os.name != "nt":
        raise AppError(E_INTERNAL, "仅支持在 Windows 中打开文件夹")
    os.startfile(target)
    return {"ok": True}


def _ext_for(fmt):
    return image_ops.EXTENSION_BY_FORMAT.get(fmt, "." + fmt)


def _default_cache(ctx):
    if ctx.get("cache") is None:
        ctx["cache"] = RenderCache(os.path.join(tempfile.gettempdir(), "picture-tools-cache"))
    return ctx["cache"]


METHODS = {
    "app.ping": _ping,
    "docx.inspect": _docx_inspect,
    "docx.extract": _docx_extract,
    "library.list": _library_list,
    "library.rename": _library_rename,
    "library.batchRename": _library_batch_rename,
    "image.signature": _image_signature,
    "image.preview": _image_preview,
    "image.save": _image_save,
    "image.batchSave": _image_batch_save,
    "system.reveal": _system_reveal,
}


def dispatch(method, params, ctx):
    """按方法名分发。未知方法抛 E_METHOD_NOT_FOUND。"""
    handler = METHODS.get(method)
    if handler is None:
        raise AppError(
            E_METHOD_NOT_FOUND,
            "未知的后端方法：%s" % (method,),
            hint="检查前后端方法名是否一致",
            context={"method": method},
        )
    return handler(params or {}, ctx)
```

- [ ] **Step 8: 实现原子保存**

保存必须在 `backend/services/image_ops.py` 中实现，因为 `rpc.py` 依赖它。在 `backend/services/image_ops.py` 顶部补 `import tempfile`，并在文件末尾追加：

```python
def _final_encoding(ops, fmt, quality):
    """compress 操作可以覆盖传入的 format / quality。"""
    resolved_format = fmt
    resolved_quality = quality
    for op in ops or []:
        if op.get("type") == "compress":
            if op.get("format"):
                resolved_format = op["format"]
            if op.get("quality") is not None:
                resolved_quality = op["quality"]
    return resolved_format.lower(), resolved_quality


def _prepare_for_format(image, fmt):
    """JPEG 与 BMP 不支持透明通道，先合成白底。"""
    if fmt in OPAQUE_FORMATS and image.mode in ("RGBA", "LA", "P"):
        converted = image.convert("RGBA")
        background = Image.new("RGB", converted.size, (255, 255, 255))
        background.paste(converted, mask=converted.split()[-1])
        return background
    if fmt in OPAQUE_FORMATS:
        return image.convert("RGB")
    return image


def save(source_path, target_path, ops, fmt="png", quality=None):
    """把操作栈应用到源图并原子写入目标路径。

    先写临时文件再替换，任一步失败都不会留下写坏一半的图片。
    """
    fmt, quality = _final_encoding(ops, fmt, quality)
    image = apply_ops(source_path, ops)
    image = _prepare_for_format(image, fmt)

    target_dir = os.path.dirname(os.path.abspath(target_path)) or "."
    os.makedirs(target_dir, exist_ok=True)

    handle, temp_path = tempfile.mkstemp(suffix=".tmp", dir=target_dir)
    os.close(handle)
    try:
        save_kwargs = {}
        if fmt in ("jpeg", "jpg"):
            save_kwargs = {"quality": int(quality) if quality is not None else 85, "optimize": True}
            pil_format = "JPEG"
        elif fmt == "png":
            save_kwargs = {"optimize": True}
            pil_format = "PNG"
        else:
            pil_format = fmt.upper()
        image.save(temp_path, format=pil_format, **save_kwargs)
        os.replace(temp_path, target_path)
    except OSError as exc:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        if getattr(exc, "errno", None) in (28, 112):
            raise AppError(E_DISK_FULL, "磁盘空间不足", detail=str(exc)) from exc
        raise
    except Exception:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise

    return {"path": str(target_path), "bytes": os.path.getsize(target_path)}
```

同时把 `E_DISK_FULL` 加入 `image_ops.py` 的 errors 导入列表。

- [ ] **Step 9: 运行 RPC 测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_rpc.py -v`
Expected: 12 passed

- [ ] **Step 10: 实现自检入口**

覆盖写入 `backend/selftest.py`：

```python
"""--selftest：不依赖 pytest 的最小自检，验证核心逻辑可用。退出码 0 表示通过。"""

import io
import os
import sys
import tempfile
import zipfile

from PIL import Image

from backend.services.docx_extract import extract_images, inspect_document
from backend.services.image_ops import apply_ops, ops_signature, save, sniff_format
from backend.services.library import batch_rename, list_images
from backend.rpc import dispatch
from backend.security import PathGuard

FIXTURE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tests", "fixtures")


def _check(condition, label):
    if condition:
        print("  ok   %s" % label)
        return 0
    print("  FAIL %s" % label, file=sys.stderr)
    return 1


def run_selftest():
    failures = 0
    print("picture-tools backend selftest")

    buf = io.BytesIO()
    Image.new("RGB", (4, 4), (1, 2, 3)).save(buf, format="PNG")
    failures += _check(sniff_format(buf.getvalue()) == "png", "sniff png magic")

    with tempfile.TemporaryDirectory() as work:
        source = os.path.join(work, "a.png")
        Image.new("RGB", (40, 20), (9, 9, 9)).save(source)

        failures += _check(apply_ops(source, []).size == (40, 20), "empty ops keeps size")
        failures += _check(
            apply_ops(source, [{"type": "rotate", "angle": 90}]).size == (20, 40),
            "rotate 90 swaps dimensions",
        )
        failures += _check(
            ops_signature(source, [{"type": "flip", "axis": "horizontal"}] * 2)
            == ops_signature(source, []),
            "double flip returns to clean signature",
        )

        out = os.path.join(work, "out.png")
        result = save(source, out, [{"type": "crop", "rect": {"left": 0, "top": 0, "right": 20, "bottom": 10}}])
        failures += _check(Image.open(out).size == (20, 10), "crop writes expected size")
        failures += _check(result["bytes"] > 0, "save reports byte count")

        images = list_images(work)
        failures += _check([e["name"] for e in images] == ["a.png", "out.png"], "library lists images")

        rename_result = batch_rename(work, [{"from": "out.png", "to": "renamed.png"}])
        failures += _check(len(rename_result["renamed"]) == 1, "batch rename applies")

        sample = os.path.join(FIXTURE_DIR, "plain.docx")
        if os.path.isfile(sample):
            inspected = inspect_document(sample)
            failures += _check(len(inspected["order"]) == 3, "docx inspect finds three images")
            extracted = extract_images(sample, os.path.join(work, "extracted"))
            failures += _check(len(extracted["written"]) == 3, "docx extract writes three files")
        else:
            print("  skip docx checks (fixtures missing)")

        ctx = {"paths": PathGuard([work]), "cache": None}
        ping = dispatch("app.ping", {}, ctx)
        failures += _check(bool(ping["python"]), "rpc ping responds")

    if failures:
        print("selftest FAILED (%d checks)" % failures, file=sys.stderr)
        return 1
    print("selftest passed")
    return 0
```

- [ ] **Step 11: 运行自检确认通过**

Run: `venv\Scripts\python.exe -m backend --selftest`
Expected: 全部 `ok`，最后一行 `selftest passed`，退出码 0

- [ ] **Step 12: 运行全部后端测试**

Run: `venv\Scripts\python.exe -m pytest backend/tests -v`
Expected: 全部通过

- [ ] **Step 13: 提交**

```bash
git add backend/cache.py backend/rpc.py backend/selftest.py backend/services/image_ops.py backend/tests/test_rpc.py backend/tests/test_cache.py
git commit -m "feat: 新增 RPC 分发器、渲染缓存与后端自检入口"
```

---

## 里程碑 M3：Electron 外壳与提取流程

### Task 11: 实现 Python 子进程管理与 stdio 传输

**Files:**
- Create: `backend/transport_stdio.py`（替换占位）
- Create: `app/main/backend.js`
- Test: `backend/tests/test_stdio_transport.py`

**Interfaces:**
- Consumes: `dispatch`、`METHODS`（Task 10）
- Produces:
  - Python 侧：`serve_stdio() -> int`，从 stdin 逐行读取 `{"id","method","params"}`，向 stdout 输出 `{"id","result"}` 或 `{"id","error":{...}}`
  - Node 侧：`createBackend({pythonPath, backendDir, onEvent, onExit}) -> {request(method, params), restart(), stop(), status()}`

- [ ] **Step 1: 写 stdio 传输的失败测试**

创建 `backend/tests/test_stdio_transport.py`：

```python
import io
import json

from backend.transport_stdio import handle_line


def _ctx(tmp_path):
    from backend.security import PathGuard

    return {"paths": PathGuard([tmp_path]), "cache": None, "root": str(tmp_path)}


def test_valid_request_returns_result_envelope(tmp_path):
    request = json.dumps({"id": 1, "method": "app.ping", "params": {}})
    response = handle_line(request, _ctx(tmp_path))
    payload = json.loads(response)
    assert payload["id"] == 1
    assert "result" in payload
    assert "error" not in payload


def test_unknown_method_returns_error_envelope(tmp_path):
    request = json.dumps({"id": 7, "method": "nope", "params": {}})
    payload = json.loads(handle_line(request, _ctx(tmp_path)))
    assert payload["id"] == 7
    assert payload["error"]["code"] == "E_METHOD_NOT_FOUND"
    assert "result" not in payload


def test_malformed_json_returns_error_with_null_id(tmp_path):
    payload = json.loads(handle_line("{not json", _ctx(tmp_path)))
    assert payload["id"] is None
    assert payload["error"]["code"] == "E_INVALID_REQUEST"


def test_error_envelope_contains_user_facing_message(tmp_path):
    from backend.errors import AppError

    class Boom(Exception):
        pass

    request = json.dumps({"id": 2, "method": "library.list", "params": {"dir": str(tmp_path / "nope")}})
    payload = json.loads(handle_line(request, _ctx(tmp_path)))
    assert payload["error"]["code"] == "E_DIR_NOT_FOUND"
    assert payload["error"]["message"]


def test_internal_exception_is_wrapped_not_leaked(tmp_path, monkeypatch):
    import backend.transport_stdio as module

    def explode(params, ctx):
        raise RuntimeError("secret internal detail")

    monkeypatch.setitem(module.METHODS, "boom.test", explode)
    request = json.dumps({"id": 3, "method": "boom.test", "params": {}})
    payload = json.loads(handle_line(request, _ctx(tmp_path)))
    assert payload["error"]["code"] == "E_INTERNAL"
    assert "secret internal detail" in (payload["error"]["detail"] or "")


def test_responses_are_single_line(tmp_path):
    request = json.dumps({"id": 4, "method": "app.ping", "params": {}})
    assert "\n" not in handle_line(request, _ctx(tmp_path)).strip()
```

- [ ] **Step 2: 运行测试确认失败**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_stdio_transport.py -v`
Expected: FAIL — `ImportError: cannot import name 'handle_line'`

- [ ] **Step 3: 实现 stdio 传输**

覆盖写入 `backend/transport_stdio.py`：

```python
"""行分隔 JSON-RPC over stdin/stdout。

每行一个请求，每行一个响应；响应必须单行，不能含内嵌换行。
"""

import json
import os
import sys
import tempfile
import traceback

from backend.cache import RenderCache
from backend.errors import E_INTERNAL, AppError
from backend.rpc import METHODS, dispatch
from backend.security import PathGuard

E_INVALID_REQUEST = "E_INVALID_REQUEST"


def make_context():
    """建立请求上下文：路径白名单与渲染缓存。"""
    cache_dir = os.path.join(tempfile.gettempdir(), "picture-tools-cache")
    return {
        "paths": PathGuard(),
        "cache": RenderCache(cache_dir),
        "root": os.getcwd(),
    }


def handle_line(line, ctx):
    """处理一行请求，返回单行 JSON 响应字符串。"""
    try:
        request = json.loads(line)
    except (ValueError, TypeError) as exc:
        return _error_response(None, E_INVALID_REQUEST, "请求不是合法 JSON", str(exc))

    if not isinstance(request, dict):
        return _error_response(None, E_INVALID_REQUEST, "请求必须是 JSON 对象", None)

    request_id = request.get("id")
    method = request.get("method")
    params = request.get("params") or {}

    if not isinstance(method, str):
        return _error_response(request_id, E_INVALID_REQUEST, "缺少 method 字段", None)

    try:
        result = dispatch(method, params, ctx)
    except AppError as exc:
        return _error_response(request_id, exc.code, exc.message, exc.detail, exc.hint, exc.context)
    except Exception as exc:  # noqa: BLE001 - 兜底，避免后端被单次请求击穿
        return _error_response(
            request_id,
            E_INTERNAL,
            "内部错误",
            "%s: %s\n%s" % (type(exc).__name__, exc, traceback.format_exc()),
            "复制诊断信息后可重启后端",
        )

    return _dump({"id": request_id, "result": result})


def _error_response(request_id, code, message, detail=None, hint=None, context=None):
    return _dump(
        {
            "id": request_id,
            "error": {
                "code": code,
                "message": message,
                "detail": detail,
                "hint": hint,
                "context": context or {},
            },
        }
    )


def _dump(payload):
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def serve_stdio():
    """主循环：逐行读取请求并写回响应，直到 stdin 关闭。"""
    ctx = make_context()
    # 强制 UTF-8，避免 Windows 默认代码页破坏中文路径
    stdin = open(sys.stdin.fileno(), "r", encoding="utf-8", newline="\n")
    stdout = open(sys.stdout.fileno(), "w", encoding="utf-8", newline="\n")

    for line in stdin:
        line = line.strip()
        if not line:
            continue
        try:
            response = handle_line(line, ctx)
        except Exception as exc:  # noqa: BLE001
            response = _error_response(None, E_INTERNAL, "内部错误", str(exc))
        stdout.write(response + "\n")
        stdout.flush()

    return 0
```

- [ ] **Step 4: 运行测试确认通过**

Run: `venv\Scripts\python.exe -m pytest backend/tests/test_stdio_transport.py -v`
Expected: 6 passed

- [ ] **Step 5: 手动验证真实 stdio 往返**

Run:
```powershell
'{"id":1,"method":"app.ping","params":{}}' | venv\Scripts\python.exe -m backend --stdio
```
Expected: 输出单行 JSON，含 `"result"` 与 `"python"` 字段

- [ ] **Step 6: 实现 Node 侧子进程管理**

创建 `app/main/backend.js`：

```js
import { spawn } from 'node:child_process'
import { EventEmitter } from 'node:events'
import path from 'node:path'

const REQUEST_TIMEOUT_MS = 120000

/**
 * 启动并监管 Python 后端子进程，转发 JSON-RPC 请求。
 */
export function createBackend({ pythonPath, backendDir, onStatusChange } = {}) {
  const emitter = new EventEmitter()
  let child = null
  let buffer = ''
  let nextId = 1
  const pending = new Map()
  let stopped = false

  function status() {
    if (stopped) return 'stopped'
    return child && !child.killed ? 'running' : 'down'
  }

  function emitStatus() {
    if (onStatusChange) onStatusChange(status())
  }

  function handleLine(line) {
    let payload
    try {
      payload = JSON.parse(line)
    } catch {
      return
    }
    const entry = pending.get(payload.id)
    if (!entry) return
    pending.delete(payload.id)
    clearTimeout(entry.timer)
    if (payload.error) {
      const error = new Error(payload.error.message)
      error.code = payload.error.code
      error.detail = payload.error.detail
      error.hint = payload.error.hint
      error.context = payload.error.context
      entry.reject(error)
    } else {
      entry.resolve(payload.result)
    }
  }

  function start() {
    if (child && !child.killed) return
    const args = ['-m', 'backend', '--stdio']
    child = spawn(pythonPath, args, {
      cwd: backendDir,
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
    })

    child.stdout.setEncoding('utf-8')
    child.stderr.setEncoding('utf-8')

    child.stdout.on('data', (chunk) => {
      buffer += chunk
      let index = buffer.indexOf('\n')
      while (index >= 0) {
        const line = buffer.slice(0, index).trim()
        buffer = buffer.slice(index + 1)
        if (line) handleLine(line)
        index = buffer.indexOf('\n')
      }
    })

    child.stderr.on('data', (chunk) => {
      emitter.emit('stderr', chunk)
    })

    child.on('exit', (code, signal) => {
      const wasStopped = stopped
      child = null
      buffer = ''
      for (const [, entry] of pending) {
        clearTimeout(entry.timer)
        const error = new Error('后端进程已退出')
        error.code = 'E_BACKEND_DOWN'
        error.detail = `exit code=${code} signal=${signal}`
        entry.reject(error)
      }
      pending.clear()
      emitStatus()
      if (!wasStopped) emitter.emit('exit', { code, signal })
    })

    emitStatus()
  }

  function request(method, params = {}) {
    if (!child) {
      const error = new Error('后端进程已退出')
      error.code = 'E_BACKEND_DOWN'
      error.hint = '点击重启后端'
      return Promise.reject(error)
    }
    const id = nextId++
    const line = JSON.stringify({ id, method, params })
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(id)
        const error = new Error('后端请求超时')
        error.code = 'E_TIMEOUT'
        reject(error)
      }, REQUEST_TIMEOUT_MS)
      pending.set(id, { resolve, reject, timer })
      child.stdin.write(line + '\n', 'utf-8')
    })
  }

  function restart() {
    return new Promise((resolve) => {
      const previous = child
      if (!previous) {
        start()
        resolve({ ok: true })
        return
      }
      previous.once('exit', () => {
        start()
        resolve({ ok: true })
      })
      previous.kill()
    })
  }

  function stop() {
    stopped = true
    if (child) {
      child.stdin.end()
      child.kill()
      child = null
    }
    emitStatus()
  }

  start()

  return {
    request,
    restart,
    stop,
    status,
    on: emitter.on.bind(emitter),
    resolvePythonPath: () => path.resolve(pythonPath),
  }
}
```

- [ ] **Step 7: 用 Node 手动验证一次往返**

Run: `node -e "import('./app/main/backend.js').then(async (m) => { const b = m.createBackend({ pythonPath: 'venv\\\\Scripts\\\\python.exe', backendDir: '.' }); console.log(await b.request('app.ping')); b.stop(); })"`
Expected: 打印含 `python` 与 `pillow` 字段的对象

注意：若 `package.json` 的 `"type": "module"` 使该命令在 Windows 下引号转义困难，改为创建一个临时文件 `scripts/smoke-backend.mjs` 并运行 `node scripts/smoke-backend.mjs`，验证后删除该临时文件（不要提交）。

- [ ] **Step 8: 提交**

```bash
git add backend/transport_stdio.py backend/tests/test_stdio_transport.py app/main/backend.js
git commit -m "feat: 实现 stdio JSON-RPC 传输与 Python 子进程管理"
```

---

### Task 12: 注册 appimg 协议与 IPC 桥

**Files:**
- Create: `app/main/protocol.js`
- Create: `app/main/ipc.js`
- Create: `app/main/main.js`
- Create: `app/preload/preload.js`
- Create: `app/main/paths.js`

**Interfaces:**
- Consumes: `createBackend`（Task 11）
- Produces:
  - `app/main/paths.js`：`resolveBackendLaunch({isDev, resourcesPath, appRoot}) -> {pythonPath, backendDir}`
  - `app/main/protocol.js`：`registerAppImageScheme() `、`registerAppImageHandler({getCacheDir, isAllowed})`
  - `app/main/ipc.js`：`registerIpc({backend, getMainWindow})`
  - preload 暴露 `window.api = { request, onEvent, platform, appVersion }`

- [ ] **Step 1: 实现路径解析**

创建 `app/main/paths.js`：

```js
import path from 'node:path'

/**
 * 开发模式用系统/venv Python 运行 backend 目录；
 * 打包后运行随应用分发的 backend.exe。
 */
export function resolveBackendLaunch({ isDev, resourcesPath, appRoot }) {
  if (isDev) {
    const pythonPath =
      process.env.PICTURE_TOOLS_PYTHON || path.join(appRoot, 'venv', 'Scripts', 'python.exe')
    return { pythonPath, backendDir: appRoot, useStdioModule: true }
  }
  const dir = path.join(resourcesPath, 'backend')
  return {
    pythonPath: path.join(dir, 'backend.exe'),
    backendDir: dir,
    useStdioModule: false,
  }
}
```

- [ ] **Step 2: 实现 appimg 协议**

创建 `app/main/protocol.js`：

```js
import { protocol, net } from 'electron'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

export const APP_IMAGE_SCHEME = 'appimg'

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.webp'])

/**
 * 必须在 app ready 之前调用：把 appimg 注册为特权 scheme，
 * 使其具备 stream 能力且不受 CSP 默认限制。
 */
export function registerAppImageScheme() {
  protocol.registerSchemesAsPrivileged([
    {
      scheme: APP_IMAGE_SCHEME,
      privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true },
    },
  ])
}

/**
 * 处理 appimg://cache/<cacheKey>.png 形式的请求。
 * 只允许读取渲染缓存目录内的文件，且扩展名必须是图片。
 */
export function registerAppImageHandler({ cacheDir }) {
  protocol.handle(APP_IMAGE_SCHEME, (request) => {
    let parsed
    try {
      parsed = new URL(request.url)
    } catch {
      return new Response('bad request', { status: 400 })
    }

    if (parsed.hostname !== 'cache') {
      return new Response('forbidden', { status: 403 })
    }

    const relative = decodeURIComponent(parsed.pathname).replace(/^\/+/, '')
    const extension = path.extname(relative).toLowerCase()

    if (!IMAGE_EXTENSIONS.has(extension) || !/^[0-9a-f]{64}\.png$/.test(relative)) {
      return new Response('forbidden', { status: 403 })
    }

    const resolved = path.resolve(cacheDir, relative)
    const root = path.resolve(cacheDir)
    if (resolved !== root && !resolved.startsWith(root + path.sep)) {
      return new Response('forbidden', { status: 403 })
    }

    return net.fetch(pathToFileURL(resolved).toString())
  })
}
```

- [ ] **Step 3: 实现 IPC 桥**

创建 `app/main/ipc.js`：

```js
import { ipcMain, dialog, shell, app } from 'electron'

/**
 * 渲染进程的唯一后端入口。方法名前缀决定路由：
 *  - dialog.* / shell.* / backend.* 由主进程直接处理
 *  - 其余转发给 Python 后端
 */
export function registerIpc({ backend, getMainWindow }) {
  ipcMain.handle('api:request', async (_event, { method, params }) => {
    try {
      if (method === 'dialog.pickDocx') {
        const result = await dialog.showOpenDialog(getMainWindow(), {
          title: '选择 Word 文件',
          properties: ['openFile'],
          filters: [{ name: 'Word 文件', extensions: ['docx'] }],
        })
        return { ok: true, result: result.canceled ? { cancelled: true } : { path: result.filePaths[0] } }
      }

      if (method === 'dialog.pickFolder') {
        const result = await dialog.showOpenDialog(getMainWindow(), {
          title: (params && params.title) || '选择文件夹',
          properties: ['openDirectory', 'createDirectory'],
        })
        return { ok: true, result: result.canceled ? { cancelled: true } : { path: result.filePaths[0] } }
      }

      if (method === 'shell.openFolder') {
        if (!params || !params.path) throw new Error('缺少 path 参数')
        await shell.openPath(params.path)
        return { ok: true, result: { ok: true } }
      }

      if (method === 'backend.restart') {
        await backend.restart()
        return { ok: true, result: { ok: true } }
      }

      if (method === 'app.info') {
        return { ok: true, result: { version: app.getVersion(), platform: process.platform } }
      }

      const result = await backend.request(method, params || {})
      return { ok: true, result }
    } catch (error) {
      return {
        ok: false,
        error: {
          code: error.code || 'E_INTERNAL',
          message: error.message || '未知错误',
          detail: error.detail || null,
          hint: error.hint || null,
          context: error.context || {},
        },
      }
    }
  })

  // 后端主动通知（进度、任务完成）转发给渲染进程
  backend.on('event', (payload) => {
    const window = getMainWindow()
    if (window && !window.isDestroyed()) {
      window.webContents.send('api:event', payload)
    }
  })

  backend.on('exit', ({ code }) => {
    const window = getMainWindow()
    if (window && !window.isDestroyed()) {
      window.webContents.send('api:backend-exit', { code })
    }
  })
}
```

- [ ] **Step 4: 实现主进程入口**

创建 `app/main/main.js`：

```js
import { app, BrowserWindow, session } from 'electron'
import path from 'node:path'
import os from 'node:os'
import { fileURLToPath } from 'node:url'

import { createBackend } from './backend.js'
import { registerIpc } from './ipc.js'
import { registerAppImageHandler, registerAppImageScheme } from './protocol.js'
import { resolveBackendLaunch } from './paths.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const appRoot = path.resolve(__dirname, '..', '..')
const isDev = process.env.ELECTRON_DEV === '1'
const devServerUrl = 'http://127.0.0.1:5173'

const cacheDir = path.join(os.tmpdir(), 'picture-tools-cache')

registerAppImageScheme()

let mainWindow = null
let backend = null

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 960,
    minHeight: 640,
    backgroundColor: '#f5f6f8',
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(appRoot, 'app', 'preload', 'preload.js'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      webSecurity: true,
    },
  })

  mainWindow.once('ready-to-show', () => mainWindow.show())

  // 禁止任何导航到外部地址与新建窗口
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  mainWindow.webContents.on('will-navigate', (event, url) => {
    const allowed = isDev ? url.startsWith(devServerUrl) : url.startsWith('file://')
    if (!allowed) event.preventDefault()
  })

  if (isDev) {
    mainWindow.loadURL(devServerUrl)
    mainWindow.webContents.openDevTools({ mode: 'detach' })
  } else {
    mainWindow.loadFile(path.join(appRoot, 'app', 'renderer', 'dist', 'index.html'))
  }
}

function installCsp() {
  session.defaultSession.webRequest.onHeadersReceived((details, callback) => {
    const policy = isDev
      ? "default-src 'self' 'unsafe-inline' data: appimg: http://127.0.0.1:5173 ws://127.0.0.1:5173"
      : "default-src 'self' 'unsafe-inline' data: appimg:"
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [policy],
      },
    })
  })
}

app.whenReady().then(() => {
  const launch = resolveBackendLaunch({
    isDev,
    resourcesPath: process.resourcesPath,
    appRoot,
  })

  backend = createBackend({
    pythonPath: launch.pythonPath,
    backendDir: launch.backendDir,
    onStatusChange: (status) => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('api:backend-status', status)
      }
    },
  })

  registerAppImageHandler({ cacheDir })
  registerIpc({ backend, getMainWindow: () => mainWindow })
  installCsp()
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  app.quit()
})

app.on('before-quit', () => {
  if (backend) backend.stop()
})
```

- [ ] **Step 5: 实现 preload**

创建 `app/preload/preload.js`：

```js
const { contextBridge, ipcRenderer } = require('electron')

const ALLOWED_METHODS = new Set([
  'app.info',
  'app.ping',
  'backend.restart',
  'dialog.pickDocx',
  'dialog.pickFolder',
  'docx.inspect',
  'docx.extract',
  'library.list',
  'library.rename',
  'library.batchRename',
  'image.signature',
  'image.preview',
  'image.save',
  'image.batchSave',
  'shell.openFolder',
  'system.reveal',
])

contextBridge.exposeInMainWorld('api', {
  async request(method, params) {
    if (!ALLOWED_METHODS.has(method)) {
      return {
        ok: false,
        error: {
          code: 'E_METHOD_NOT_ALLOWED',
          message: `不允许调用的方法：${method}`,
          detail: null,
          hint: null,
          context: {},
        },
      }
    }
    return ipcRenderer.invoke('api:request', { method, params: params || {} })
  },

  onEvent(handler) {
    const listener = (_event, payload) => handler(payload)
    ipcRenderer.on('api:event', listener)
    return () => ipcRenderer.removeListener('api:event', listener)
  },

  onBackendStatus(handler) {
    const listener = (_event, status) => handler(status)
    ipcRenderer.on('api:backend-status', listener)
    return () => ipcRenderer.removeListener('api:backend-status', listener)
  },

  onBackendExit(handler) {
    const listener = (_event, payload) => handler(payload)
    ipcRenderer.on('api:backend-exit', listener)
    return () => ipcRenderer.removeListener('api:backend-exit', listener)
  },
})
```

- [ ] **Step 6: 创建 Vite 与 Electron 的目录占位，验证应用可启动**

创建 `app/renderer/index.html`：

```html
<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Word 图片提取与重命名工具</title>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/src/main.js"></script>
  </body>
</html>
```

创建 `vite.config.js`：

```js
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import path from 'node:path'

export default defineConfig({
  root: path.join(__dirname, 'app', 'renderer'),
  base: './',
  plugins: [vue()],
  server: { port: 5173, strictPort: true, host: '127.0.0.1' },
  build: {
    outDir: path.join(__dirname, 'app', 'renderer', 'dist'),
    emptyOutDir: true,
  },
  resolve: {
    alias: { '@': path.join(__dirname, 'app', 'renderer', 'src') },
  },
})
```

创建 `app/renderer/src/main.js` 的最小版本（Task 13 会扩充）：

```js
import { createApp } from 'vue'
import { createPinia } from 'pinia'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import zhCn from 'element-plus/dist/locale/zh-cn.mjs'

import App from './App.vue'
import './styles/tokens.css'

const app = createApp(App)
app.use(createPinia())
app.use(ElementPlus, { locale: zhCn })
app.mount('#app')
```

创建 `app/renderer/src/App.vue` 的最小版本：

```vue
<template>
  <div class="boot-check">
    <h1>Word 图片提取与重命名工具</h1>
    <p>后端状态：{{ status }}</p>
    <pre v-if="ping">{{ ping }}</pre>
  </div>
</template>

<script setup>
import { onMounted, ref } from 'vue'

const status = ref('连接中')
const ping = ref(null)

onMounted(async () => {
  const response = await window.api.request('app.ping', {})
  if (response.ok) {
    status.value = '已连接'
    ping.value = JSON.stringify(response.result, null, 2)
  } else {
    status.value = `连接失败：${response.error.message}`
  }
})
</script>
```

创建 `app/renderer/src/styles/tokens.css`：

```css
:root {
  --pt-space-1: 4px;
  --pt-space-2: 8px;
  --pt-space-3: 12px;
  --pt-space-4: 16px;
  --pt-space-5: 24px;
  --pt-radius-sm: 4px;
  --pt-radius-md: 8px;
  --pt-font-ui: "Microsoft YaHei UI", "Segoe UI", system-ui, sans-serif;
  --pt-font-mono: "Cascadia Mono", Consolas, monospace;
  --pt-pane-list-width: 288px;
  --pt-pane-inspector-width: 320px;
  --pt-color-bg: #f5f6f8;
  --pt-color-surface: #ffffff;
  --pt-color-border: #dcdfe6;
  --pt-color-text: #303133;
  --pt-color-text-muted: #909399;
  --pt-color-accent: #3b6ef5;
}

:root[data-theme='dark'] {
  --pt-color-bg: #1e1f22;
  --pt-color-surface: #26282c;
  --pt-color-border: #3a3d42;
  --pt-color-text: #e5e7eb;
  --pt-color-text-muted: #9aa0a6;
  --pt-color-accent: #5b8cff;
}

body {
  margin: 0;
  font-family: var(--pt-font-ui);
  background: var(--pt-color-bg);
  color: var(--pt-color-text);
}
```

- [ ] **Step 7: 安装依赖**

Run: `npm install`
Expected: 安装成功，生成 `node_modules/` 与 `package-lock.json`

- [ ] **Step 8: 启动开发模式验证端到端连通**

Run: `npm run dev`
Expected: Electron 窗口出现，页面显示「后端状态：已连接」，下方打印包含 `python` 与 `pillow` 字段的 JSON

验证后按 `Ctrl+C` 结束。

- [ ] **Step 9: 提交**

```bash
git add app/ vite.config.js package-lock.json
git commit -m "feat: 打通 Electron 外壳、appimg 协议与 IPC 桥，前后端首次连通"
```

---

### Task 13: 实现渲染进程传输层与三栏骨架

**Files:**
- Create: `app/renderer/src/api/transport.js`
- Create: `app/renderer/src/components/Toolbar.vue`
- Create: `app/renderer/src/components/StatusBar.vue`
- Create: `app/renderer/src/components/Workspace.vue`
- Create: `app/renderer/src/components/pane/ListPane.vue`
- Create: `app/renderer/src/components/pane/PreviewPane.vue`
- Create: `app/renderer/src/components/pane/InspectorPane.vue`
- Modify: `app/renderer/src/App.vue`
- Test: `app/renderer/src/api/__tests__/transport.spec.js`

**Interfaces:**
- Consumes: `window.api`（Task 12）
- Produces: `api.request(method, params) -> Promise<result>`（失败时抛带 `code` 的 Error）；`api.onEvent(handler) -> unsubscribe`；`api.imageUrl(cacheKey) -> string`；`api.isBackendHealthy() -> boolean`

- [ ] **Step 1: 写传输层的失败测试**

创建 `app/renderer/src/api/__tests__/transport.spec.js`：

```js
import { beforeEach, describe, expect, it, vi } from 'vitest'

import { createTransport } from '../transport.js'

function makeBridge(response) {
  return {
    request: vi.fn(async () => response),
    onEvent: vi.fn(() => () => {}),
    onBackendStatus: vi.fn(() => () => {}),
    onBackendExit: vi.fn(() => () => {}),
  }
}

describe('transport', () => {
  let bridge

  beforeEach(() => {
    bridge = makeBridge({ ok: true, result: { value: 42 } })
  })

  it('resolves with result on success', async () => {
    const transport = createTransport(bridge)
    await expect(transport.request('app.ping')).resolves.toEqual({ value: 42 })
    expect(bridge.request).toHaveBeenCalledWith('app.ping', {})
  })

  it('throws an Error carrying the backend error code', async () => {
    const failing = makeBridge({
      ok: false,
      error: { code: 'E_NAME_EXISTS', message: '目标文件名已存在', hint: '换一个名称' },
    })
    const transport = createTransport(failing)
    await expect(transport.request('library.rename')).rejects.toMatchObject({
      code: 'E_NAME_EXISTS',
      message: '目标文件名已存在',
      hint: '换一个名称',
    })
  })

  it('throws E_NO_BRIDGE when window.api is absent', async () => {
    const transport = createTransport(null)
    await expect(transport.request('app.ping')).rejects.toMatchObject({ code: 'E_NO_BRIDGE' })
  })

  it('builds appimg urls from a cache key', () => {
    const transport = createTransport(bridge)
    expect(transport.imageUrl('a'.repeat(64))).toBe(`appimg://cache/${'a'.repeat(64)}.png`)
  })

  it('rejects invalid cache keys instead of building a url', () => {
    const transport = createTransport(bridge)
    expect(transport.imageUrl('../escape')).toBe('')
    expect(transport.imageUrl('')).toBe('')
  })

  it('forwards backend status changes', () => {
    const transport = createTransport(bridge)
    transport.onBackendStatus?.(() => {})
    expect(bridge.onBackendStatus).toHaveBeenCalled()
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/api/__tests__/transport.spec.js`
Expected: FAIL — 无法解析 `../transport.js`

- [ ] **Step 3: 实现传输层**

创建 `app/renderer/src/api/transport.js`：

```js
const CACHE_KEY_RE = /^[0-9a-f]{64}$/

/**
 * 把 IPC 桥包装成会抛错的 Promise 接口。
 * stdio 与 HTTP 两种模式共用同一组调用方，差异只在这里。
 */
export function createTransport(bridge) {
  const api = bridge || (typeof window !== 'undefined' ? window.api : null)

  async function request(method, params = {}) {
    if (!api) {
      const error = new Error('后端桥未就绪')
      error.code = 'E_NO_BRIDGE'
      error.hint = '重启应用后重试'
      throw error
    }

    const response = await api.request(method, params)
    if (!response) {
      const error = new Error('后端无响应')
      error.code = 'E_NO_BRIDGE'
      throw error
    }
    if (!response.ok) {
      const payload = response.error || {}
      const error = new Error(payload.message || '请求失败')
      error.code = payload.code || 'E_INTERNAL'
      error.detail = payload.detail || null
      error.hint = payload.hint || null
      error.context = payload.context || {}
      throw error
    }
    return response.result
  }

  function imageUrl(cacheKey) {
    if (!CACHE_KEY_RE.test(cacheKey || '')) return ''
    return `appimg://cache/${cacheKey}.png`
  }

  function onEvent(handler) {
    return api && api.onEvent ? api.onEvent(handler) : () => {}
  }

  function onBackendStatus(handler) {
    return api && api.onBackendStatus ? api.onBackendStatus(handler) : () => {}
  }

  function onBackendExit(handler) {
    return api && api.onBackendExit ? api.onBackendExit(handler) : () => {}
  }

  return { request, imageUrl, onEvent, onBackendStatus, onBackendExit }
}

export const api = createTransport(null)
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/api/__tests__/transport.spec.js`
Expected: 6 passed

- [ ] **Step 5: 实现三栏骨架组件**

创建 `app/renderer/src/components/Workspace.vue`：

```vue
<template>
  <div class="workspace">
    <aside class="pane pane--list" :style="{ width: listWidth + 'px' }">
      <ListPane />
    </aside>

    <div
      class="splitter"
      role="separator"
      aria-label="调整图片列表宽度"
      @mousedown="startDrag('list', $event)"
    />

    <main class="pane pane--preview">
      <PreviewPane />
    </main>

    <div
      class="splitter"
      role="separator"
      aria-label="调整属性面板宽度"
      @mousedown="startDrag('inspector', $event)"
    />

    <aside class="pane pane--inspector" :style="{ width: inspectorWidth + 'px' }">
      <InspectorPane />
    </aside>
  </div>
</template>

<script setup>
import { onBeforeUnmount, ref } from 'vue'

import ListPane from './pane/ListPane.vue'
import PreviewPane from './pane/PreviewPane.vue'
import InspectorPane from './pane/InspectorPane.vue'

const MIN_LIST = 220
const MAX_LIST = 420
const MIN_INSPECTOR = 260
const MAX_INSPECTOR = 460

const listWidth = ref(288)
const inspectorWidth = ref(320)

let dragging = null

function startDrag(kind, event) {
  dragging = { kind, startX: event.clientX, startWidth: kind === 'list' ? listWidth.value : inspectorWidth.value }
  window.addEventListener('mousemove', onMove)
  window.addEventListener('mouseup', stopDrag)
  event.preventDefault()
}

function onMove(event) {
  if (!dragging) return
  const delta = event.clientX - dragging.startX
  if (dragging.kind === 'list') {
    listWidth.value = Math.min(MAX_LIST, Math.max(MIN_LIST, dragging.startWidth + delta))
  } else {
    inspectorWidth.value = Math.min(MAX_INSPECTOR, Math.max(MIN_INSPECTOR, dragging.startWidth - delta))
  }
}

function stopDrag() {
  dragging = null
  window.removeEventListener('mousemove', onMove)
  window.removeEventListener('mouseup', stopDrag)
}

onBeforeUnmount(stopDrag)
</script>

<style scoped>
.workspace {
  display: flex;
  flex: 1;
  min-height: 0;
}

.pane {
  display: flex;
  flex-direction: column;
  min-width: 0;
  overflow: hidden;
}

.pane--preview {
  flex: 1;
  background: var(--pt-color-surface);
}

.splitter {
  width: 4px;
  cursor: col-resize;
  background: var(--pt-color-border);
  flex: none;
}

.splitter:hover {
  background: var(--pt-color-accent);
}
</style>
```

创建 `app/renderer/src/components/Toolbar.vue`：

```vue
<template>
  <header class="toolbar">
    <div class="toolbar__group">
      <el-button type="primary" @click="$emit('import-word')">导入 Word</el-button>
      <el-button @click="$emit('open-folder')">打开文件夹</el-button>
    </div>

    <div class="toolbar__group toolbar__group--center">
      <el-input
        v-model="keyword"
        placeholder="搜索文件名"
        clearable
        style="width: 220px"
        @update:model-value="$emit('search', $event)"
      />
      <el-select
        :model-value="sortBy"
        style="width: 150px"
        @update:model-value="$emit('sort', $event)"
      >
        <el-option label="原始顺序" value="original" />
        <el-option label="文件名" value="name" />
        <el-option label="修改时间" value="mtime" />
        <el-option label="尺寸" value="size" />
      </el-select>
    </div>

    <div class="toolbar__group toolbar__group--end">
      <el-button :disabled="!dirtyCount" @click="$emit('save-current')">保存</el-button>
      <el-button :disabled="!dirtyCount" @click="$emit('save-all')">全部保存</el-button>
      <el-button text @click="$emit('toggle-theme')">{{ theme === 'dark' ? '浅色' : '深色' }}</el-button>
    </div>
  </header>
</template>

<script setup>
import { ref } from 'vue'

defineProps({
  sortBy: { type: String, default: 'original' },
  dirtyCount: { type: Number, default: 0 },
  theme: { type: String, default: 'light' },
})

defineEmits(['import-word', 'open-folder', 'search', 'sort', 'save-current', 'save-all', 'toggle-theme'])

const keyword = ref('')
</script>

<style scoped>
.toolbar {
  display: flex;
  align-items: center;
  gap: var(--pt-space-3);
  padding: var(--pt-space-2) var(--pt-space-3);
  background: var(--pt-color-surface);
  border-bottom: 1px solid var(--pt-color-border);
  flex: none;
}

.toolbar__group {
  display: flex;
  align-items: center;
  gap: var(--pt-space-2);
}

.toolbar__group--center {
  flex: 1;
  justify-content: center;
}

.toolbar__group--end {
  margin-left: auto;
}
</style>
```

创建 `app/renderer/src/components/StatusBar.vue`：

```vue
<template>
  <footer class="status-bar">
    <span>共 {{ total }} 张</span>
    <el-divider direction="vertical" />
    <span>已选 {{ selected }}</span>
    <el-divider direction="vertical" />
    <span :class="{ 'is-dirty': dirtyCount > 0 }">未保存 {{ dirtyCount }}</span>
    <span class="status-bar__spacer" />
    <el-tag :type="backendTagType" size="small">{{ backendLabel }}</el-tag>
  </footer>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  total: { type: Number, default: 0 },
  selected: { type: Number, default: 0 },
  dirtyCount: { type: Number, default: 0 },
  backendStatus: { type: String, default: 'down' },
})

const backendLabel = computed(() => {
  if (props.backendStatus === 'running') return '后端已连接'
  if (props.backendStatus === 'stopped') return '后端已停止'
  return '后端未连接'
})

const backendTagType = computed(() => (props.backendStatus === 'running' ? 'success' : 'danger'))
</script>

<style scoped>
.status-bar {
  display: flex;
  align-items: center;
  gap: var(--pt-space-2);
  padding: var(--pt-space-1) var(--pt-space-3);
  font-size: 12px;
  color: var(--pt-color-text-muted);
  background: var(--pt-color-surface);
  border-top: 1px solid var(--pt-color-border);
  flex: none;
}

.status-bar__spacer {
  flex: 1;
}

.is-dirty {
  color: var(--pt-color-accent);
  font-weight: 600;
}
</style>
```

创建 `app/renderer/src/components/pane/ListPane.vue`（本任务仅静态占位，Task 15 补数据）：

```vue
<template>
  <div class="list-pane">
    <div class="list-pane__header">图片列表</div>
    <div class="list-pane__body">
      <el-empty v-if="images.length === 0" description="尚未打开图片文件夹" :image-size="72" />
      <ul v-else class="list-pane__items">
        <li v-for="item in images" :key="item.name" class="list-pane__item">
          <span class="list-pane__name">{{ item.name }}</span>
          <span class="list-pane__meta">{{ item.width }}×{{ item.height }}</span>
        </li>
      </ul>
    </div>
    <div class="list-pane__footer">已选 {{ selected }}</div>
  </div>
</template>

<script setup>
defineProps({
  images: { type: Array, default: () => [] },
  selected: { type: Number, default: 0 },
})
</script>

<style scoped>
.list-pane {
  display: flex;
  flex-direction: column;
  height: 100%;
  background: var(--pt-color-surface);
  border-right: 1px solid var(--pt-color-border);
}

.list-pane__header,
.list-pane__footer {
  padding: var(--pt-space-2) var(--pt-space-3);
  font-size: 12px;
  font-weight: 600;
  color: var(--pt-color-text-muted);
  flex: none;
}

.list-pane__footer {
  border-top: 1px solid var(--pt-color-border);
}

.list-pane__body {
  flex: 1;
  overflow: auto;
  min-height: 0;
}

.list-pane__items {
  list-style: none;
  margin: 0;
  padding: 0;
}

.list-pane__item {
  display: flex;
  justify-content: space-between;
  gap: var(--pt-space-2);
  padding: var(--pt-space-1) var(--pt-space-3);
  font-size: 13px;
}

.list-pane__meta {
  color: var(--pt-color-text-muted);
  font-size: 12px;
}
</style>
```

创建 `app/renderer/src/components/pane/PreviewPane.vue`（本任务仅静态占位，Task 18–20 补画布）：

```vue
<template>
  <div class="preview-pane">
    <el-empty description="尚未选择图片" :image-size="96" />
  </div>
</template>

<style scoped>
.preview-pane {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
}
</style>
```

创建 `app/renderer/src/components/pane/InspectorPane.vue`（本任务仅静态占位，Task 21 补内容）：

```vue
<template>
  <div class="inspector-pane">
    <el-collapse :model-value="['info', 'rename', 'edit']">
      <el-collapse-item title="文件信息" name="info">
        <el-empty description="尚未选择图片" :image-size="56" />
      </el-collapse-item>
      <el-collapse-item title="重命名" name="rename">
        <el-empty description="尚未选择图片" :image-size="56" />
      </el-collapse-item>
      <el-collapse-item title="编辑参数" name="edit">
        <el-empty description="尚未选择图片" :image-size="56" />
      </el-collapse-item>
    </el-collapse>
  </div>
</template>

<style scoped>
.inspector-pane {
  height: 100%;
  overflow: auto;
  padding: 0 var(--pt-space-2);
  background: var(--pt-color-surface);
  border-left: 1px solid var(--pt-color-border);
}
</style>
```

- [ ] **Step 6: 组装 App.vue**

覆盖写入 `app/renderer/src/App.vue`：

```vue
<template>
  <div class="app-shell">
    <Toolbar
      :sort-by="library.sortBy"
      :dirty-count="editor.dirtyCount"
      :theme="theme"
      @import-word="onImportWord"
      @open-folder="onOpenFolder"
      @sort="library.setSortBy"
      @search="library.setKeyword"
      @save-current="onSaveCurrent"
      @save-all="onSaveAll"
      @toggle-theme="toggleTheme"
    />

    <Workspace />

    <StatusBar
      :total="library.images.length"
      :selected="library.selectedCount"
      :dirty-count="editor.dirtyCount"
      :backend-status="backendStatus"
    />
  </div>
</template>

<script setup>
import { onMounted, onBeforeUnmount, ref } from 'vue'
import { ElMessage } from 'element-plus'

import Toolbar from './components/Toolbar.vue'
import Workspace from './components/Workspace.vue'
import StatusBar from './components/StatusBar.vue'
import { api } from './api/transport.js'
import { useLibraryStore } from './stores/library.js'
import { useEditorStore } from './stores/editor.js'

const library = useLibraryStore()
const editor = useEditorStore()

const backendStatus = ref('down')
const theme = ref('light')

let unsubscribeStatus = null
let unsubscribeExit = null

function toggleTheme() {
  theme.value = theme.value === 'dark' ? 'light' : 'dark'
  document.documentElement.dataset.theme = theme.value
}

async function onImportWord() {
  ElMessage.info('导入流程将在后续任务中接入')
}

async function onOpenFolder() {
  ElMessage.info('打开文件夹流程将在后续任务中接入')
}

async function onSaveCurrent() {
  ElMessage.info('保存流程将在后续任务中接入')
}

async function onSaveAll() {
  ElMessage.info('批量保存流程将在后续任务中接入')
}

onMounted(async () => {
  unsubscribeStatus = api.onBackendStatus((status) => {
    backendStatus.value = status
  })
  unsubscribeExit = api.onBackendExit(() => {
    backendStatus.value = 'down'
    ElMessage.error('后端进程已退出，请点击重启')
  })

  try {
    await api.request('app.ping', {})
    backendStatus.value = 'running'
  } catch (error) {
    backendStatus.value = 'down'
    ElMessage.error(error.message)
  }
})

onBeforeUnmount(() => {
  if (unsubscribeStatus) unsubscribeStatus()
  if (unsubscribeExit) unsubscribeExit()
})
</script>

<style scoped>
.app-shell {
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}
</style>
```

- [ ] **Step 7: 创建 store 骨架（Task 15、19 会扩充）**

创建 `app/renderer/src/stores/library.js`：

```js
import { defineStore } from 'pinia'

export const useLibraryStore = defineStore('library', {
  state: () => ({
    folder: '',
    images: [],
    selectedNames: [],
    keyword: '',
    sortBy: 'original',
    loading: false,
  }),
  getters: {
    selectedCount: (state) => state.selectedNames.length,
    visibleImages(state) {
      const keyword = state.keyword.trim().toLowerCase()
      const filtered = keyword
        ? state.images.filter((item) => item.name.toLowerCase().includes(keyword))
        : [...state.images]
      if (state.sortBy === 'name') {
        filtered.sort((a, b) => a.name.localeCompare(b.name, 'zh-Hans-CN'))
      } else if (state.sortBy === 'mtime') {
        filtered.sort((a, b) => a.mtime - b.mtime)
      } else if (state.sortBy === 'size') {
        filtered.sort((a, b) => b.width * b.height - a.width * a.height)
      }
      return filtered
    },
  },
  actions: {
    setKeyword(value) {
      this.keyword = value || ''
    },
    setSortBy(value) {
      this.sortBy = value || 'original'
    },
  },
})
```

创建 `app/renderer/src/stores/editor.js`（Task 17 会用完整版本替换）：

```js
import { defineStore } from 'pinia'

/**
 * 最小骨架：只承载「当前选中哪张图」。
 * 操作栈、撤销历史与脏标记在 Task 17 引入。
 */
export const useEditorStore = defineStore('editor', {
  state: () => ({
    folder: '',
    currentName: '',
    currentPath: '',
  }),
  actions: {
    setCurrent(name, folder) {
      this.currentName = name || ''
      if (folder) this.folder = folder
      const trimmed = (this.folder || '').replace(/[\\/]+$/, '')
      this.currentPath = name ? (trimmed ? `${trimmed}\\${name}` : name) : ''
    },
  },
})
```

- [ ] **Step 8: 运行单元测试与开发模式**

Run: `npx vitest run`
Expected: transport 测试通过

Run: `npm run dev`
Expected: 窗口显示顶部工具栏、三栏布局、底部状态栏；状态栏右下显示「后端已连接」

- [ ] **Step 9: 提交**

```bash
git add app/renderer/
git commit -m "feat: 实现传输层与三栏工作台骨架"
```

---

### Task 14: 实现 Word 导入与提取流程

**Files:**
- Create: `app/renderer/src/composables/useExtract.js`
- Create: `app/renderer/src/stores/tasks.js`
- Create: `app/renderer/src/components/dialog/ExtractDialog.vue`
- Modify: `app/renderer/src/App.vue`
- Test: `app/renderer/src/composables/__tests__/extract.spec.js`

**Interfaces:**
- Consumes: `api`（Task 13）、`library` store（Task 13）
- Produces: `summarizeExtractResult(result) -> { headline, skippedNote, orderWarning }`；`useExtract()` 返回 `{ inspecting, inspectingBusy, confirmExtract, cancelInspect }`

- [ ] **Step 1: 写摘要逻辑的失败测试**

创建 `app/renderer/src/composables/__tests__/extract.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { summarizeExtractResult } from '../useExtract.js'

const base = { written: [], skipped: [], renamed: [], orderSource: 'xml', cancelled: false }

describe('summarizeExtractResult', () => {
  it('reports a plain success count', () => {
    const summary = summarizeExtractResult({ ...base, written: ['001.png', '002.png'] })
    expect(summary.headline).toContain('2')
    expect(summary.orderWarning).toBeNull()
  })

  it('warns when order came from filenames', () => {
    const summary = summarizeExtractResult({ ...base, written: ['001.png'], orderSource: 'filename' })
    expect(summary.orderWarning).toContain('文件名')
  })

  it('warns when order was guessed', () => {
    const summary = summarizeExtractResult({ ...base, written: ['001.png'], orderSource: 'guess' })
    expect(summary.orderWarning).toBeTruthy()
  })

  it('does not warn for xml or regex order', () => {
    expect(summarizeExtractResult({ ...base, written: ['a.png'], orderSource: 'xml' }).orderWarning).toBeNull()
    expect(summarizeExtractResult({ ...base, written: ['a.png'], orderSource: 'regex' }).orderWarning).toBeNull()
  })

  it('mentions skipped files and rename count', () => {
    const summary = summarizeExtractResult({
      ...base,
      written: ['001.png'],
      skipped: [{ partName: 'word/media/x.emf', reason: 'unsupported-format' }],
      renamed: [{ from: '001.png', to: '001_1.png' }],
    })
    expect(summary.skippedNote).toContain('1')
    expect(summary.headline).toContain('1')
  })

  it('flags cancellation', () => {
    const summary = summarizeExtractResult({ ...base, written: ['001.png'], cancelled: true })
    expect(summary.headline).toContain('取消')
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/extract.spec.js`
Expected: FAIL — 无法解析 `../useExtract.js`

- [ ] **Step 3: 实现提取流程组合式函数**

创建 `app/renderer/src/composables/useExtract.js`：

```js
import { ref } from 'vue'

import { api } from '../api/transport.js'
import { useLibraryStore } from '../stores/library.js'
import { useTasksStore } from '../stores/tasks.js'

const ORDER_WARNINGS = {
  filename: '未能读取图片插入顺序，已按媒体文件名排序，请核对结果',
  guess: '未能确定图片插入顺序，结果可能不准确，请在列表中手动核对顺序',
}

/**
 * 把后端提取结果转成面向用户的摘要。排序来源不可靠时必须给出警告。
 */
export function summarizeExtractResult(result) {
  const writtenCount = (result.written || []).length
  const skippedCount = (result.skipped || []).length

  let headline = `已提取 ${writtenCount} 张图片`
  if (result.cancelled) {
    headline = `已取消，已提取 ${writtenCount} 张图片`
  }
  if ((result.renamed || []).length > 0) {
    headline += `，其中 ${result.renamed.length} 张因重名自动加了序号`
  }

  const skippedNote =
    skippedCount > 0 ? `另有 ${skippedCount} 张被跳过（格式不支持或重名）` : null

  return {
    headline,
    skippedNote,
    orderWarning: ORDER_WARNINGS[result.orderSource] || null,
  }
}

export function useExtract() {
  const library = useLibraryStore()
  const tasks = useTasksStore()

  const inspecting = ref(null)
  const inspectingBusy = ref(false)
  const extracting = ref(false)
  const progress = ref({ done: 0, total: 0 })

  async function chooseWordFile() {
    const response = await api.request('dialog.pickDocx', {})
    if (response.cancelled) return null
    return response.path
  }

  async function inspect(path) {
    inspectingBusy.value = true
    try {
      const result = await api.request('docx.inspect', { path })
      inspecting.value = {
        path,
        mediaCount: result.media.length,
        orderSource: result.orderSource,
        order: result.order,
      }
      return inspecting.value
    } finally {
      inspectingBusy.value = false
    }
  }

  async function confirmExtract({ path, outDir, pattern, startIndex, onConflict }) {
    extracting.value = true
    progress.value = { done: 0, total: inspecting.value ? inspecting.value.mediaCount : 0 }
    const taskId = tasks.start({ type: 'extract', label: '提取图片' })
    try {
      const result = await api.request('docx.extract', {
        path,
        outDir,
        pattern: pattern || '{index:03d}',
        startIndex: startIndex || 1,
        onConflict: onConflict || 'autoSuffix',
      })
      tasks.finish(taskId, summarizeExtractResult(result))
      library.folder = outDir
      await library.loadFolder(outDir)
      return { result, summary: summarizeExtractResult(result) }
    } catch (error) {
      tasks.fail(taskId, error)
      throw error
    } finally {
      extracting.value = false
    }
  }

  function cancelInspect() {
    inspecting.value = null
  }

  return {
    inspecting,
    inspectingBusy,
    extracting,
    progress,
    chooseWordFile,
    inspect,
    confirmExtract,
    cancelInspect,
  }
}
```

- [ ] **Step 4: 实现任务 store**

创建 `app/renderer/src/stores/tasks.js`：

```js
import { defineStore } from 'pinia'

let seq = 0

export const useTasksStore = defineStore('tasks', {
  state: () => ({
    items: [],
  }),
  getters: {
    active: (state) => state.items.filter((item) => item.status === 'running'),
    hasActive: (state) => state.items.some((item) => item.status === 'running'),
  },
  actions: {
    start({ type, label }) {
      seq += 1
      const id = `task-${seq}`
      this.items.push({
        id,
        type,
        label,
        status: 'running',
        done: 0,
        total: 0,
        summary: null,
        error: null,
      })
      return id
    },
    updateProgress(id, done, total) {
      const task = this.items.find((item) => item.id === id)
      if (task) {
        task.done = done
        task.total = total
      }
    },
    finish(id, summary) {
      const task = this.items.find((item) => item.id === id)
      if (task) {
        task.status = 'done'
        task.summary = summary
      }
    },
    fail(id, error) {
      const task = this.items.find((item) => item.id === id)
      if (task) {
        task.status = 'failed'
        task.error = { code: error.code, message: error.message, hint: error.hint }
      }
    },
    cancel(id) {
      const task = this.items.find((item) => item.id === id)
      if (task) task.status = 'cancelled'
    },
    clearFinished() {
      this.items = this.items.filter((item) => item.status === 'running')
    },
  },
})
```

- [ ] **Step 5: 在 library store 中补 loadFolder**

在 `app/renderer/src/stores/library.js` 的 `actions` 中加入：

```js
    async loadFolder(folder) {
      this.loading = true
      try {
        const result = await api.request('library.list', { dir: folder })
        this.folder = folder
        this.images = result.images
        this.selectedNames = []
      } finally {
        this.loading = false
      }
    },
    selectOnly(name) {
      this.selectedNames = name ? [name] : []
    },
    toggleSelection(name) {
      if (this.selectedNames.includes(name)) {
        this.selectedNames = this.selectedNames.filter((item) => item !== name)
      } else {
        this.selectedNames = [...this.selectedNames, name]
      }
    },
```

并在文件顶部补导入：

```js
import { api } from '../api/transport.js'
```

- [ ] **Step 6: 实现提取对话框**

创建 `app/renderer/src/components/dialog/ExtractDialog.vue`：

```vue
<template>
  <el-dialog
    :model-value="visible"
    title="导入 Word 并提取图片"
    width="560px"
    @update:model-value="$emit('update:visible', $event)"
  >
    <el-form label-width="110px">
      <el-form-item label="Word 文件">
        <el-input :model-value="wordPath" readonly placeholder="尚未选择">
          <template #append>
            <el-button @click="$emit('pick-word')">浏览</el-button>
          </template>
        </el-input>
      </el-form-item>

      <el-form-item label="输出文件夹">
        <el-input :model-value="outDir" readonly placeholder="尚未选择">
          <template #append>
            <el-button @click="$emit('pick-outdir')">浏览</el-button>
          </template>
        </el-input>
      </el-form-item>

      <el-form-item label="命名格式">
        <el-input v-model="localPattern" placeholder="{index:03d}" />
        <div class="hint">使用 {index} 占位起始序号，例如 pic-{index} 得到 pic-1、pic-2</div>
      </el-form-item>

      <el-form-item label="起始序号">
        <el-input-number v-model="localStart" :min="1" :max="9999" />
      </el-form-item>

      <el-form-item label="重名处理">
        <el-select v-model="localConflict" style="width: 100%">
          <el-option label="自动加序号（推荐）" value="autoSuffix" />
          <el-option label="跳过已存在" value="reject" />
          <el-option label="覆盖已存在" value="overwrite" />
        </el-select>
      </el-form-item>

      <el-alert
        v-if="orderWarning"
        type="warning"
        :closable="false"
        show-icon
        :title="orderWarning"
      />
    </el-form>

    <template #footer>
      <el-button @click="$emit('update:visible', false)">取消</el-button>
      <el-button
        type="primary"
        :loading="busy"
        :disabled="!wordPath || !outDir"
        @click="$emit('confirm')"
      >
        开始提取
      </el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { computed, ref, watch } from 'vue'

const props = defineProps({
  visible: Boolean,
  wordPath: { type: String, default: '' },
  outDir: { type: String, default: '' },
  orderSource: { type: String, default: '' },
  busy: Boolean,
})

defineEmits(['update:visible', 'pick-word', 'pick-outdir', 'confirm'])

const localPattern = ref('{index:03d}')
const localStart = ref(1)
const localConflict = ref('autoSuffix')

const orderWarning = computed(() => {
  if (props.orderSource === 'filename') return '未能读取图片插入顺序，将按媒体文件名排序，请核对结果'
  if (props.orderSource === 'guess') return '未能确定图片插入顺序，结果可能不准确'
  return null
})

watch(
  () => props.visible,
  (open) => {
    if (open) {
      localPattern.value = '{index:03d}'
      localStart.value = 1
      localConflict.value = 'autoSuffix'
    }
  },
)

defineExpose({ localPattern, localStart, localConflict })
</script>

<style scoped>
.hint {
  font-size: 12px;
  color: var(--pt-color-text-muted);
  line-height: 1.6;
}
</style>
```

- [ ] **Step 7: 在 App.vue 中接入导入流程**

把 `app/renderer/src/App.vue` 的 `<script setup>` 中 `onImportWord` 与 `onOpenFolder` 替换为：

```js
import ExtractDialog from './components/dialog/ExtractDialog.vue'
import { useExtract } from './composables/useExtract.js'

const extract = useExtract()
const extractVisible = ref(false)
const wordPath = ref('')
const outDir = ref('')
const extractBusy = ref(false)

async function onImportWord() {
  const path = await extract.chooseWordFile()
  if (!path) return
  wordPath.value = path
  outDir.value = `${path.replace(/\.docx$/i, '')}_images`
  extractVisible.value = true
  try {
    await extract.inspect(path)
  } catch (error) {
    ElMessage.error(error.message)
    extractVisible.value = false
  }
}

async function onPickOutDir() {
  const response = await api.request('dialog.pickFolder', { title: '选择输出文件夹' })
  if (!response.cancelled) outDir.value = response.path
}

async function onConfirmExtract() {
  extractBusy.value = true
  try {
    const pattern = extractDialogRef.value?.localPattern || '{index:03d}'
    const startIndex = extractDialogRef.value?.localStart || 1
    const onConflict = extractDialogRef.value?.localConflict || 'autoSuffix'
    const { summary } = await extract.confirmExtract({
      path: wordPath.value,
      outDir: outDir.value,
      pattern,
      startIndex,
      onConflict,
    })
    ElMessage.success(summary.headline)
    if (summary.orderWarning) ElMessage.warning(summary.orderWarning)
    extractVisible.value = false
  } catch (error) {
    ElMessage.error(`${error.message}${error.hint ? `（${error.hint}）` : ''}`)
  } finally {
    extractBusy.value = false
  }
}

async function onOpenFolder() {
  const response = await api.request('dialog.pickFolder', { title: '选择图片文件夹' })
  if (response.cancelled) return
  try {
    await library.loadFolder(response.path)
    ElMessage.success(`已加载 ${library.images.length} 张图片`)
  } catch (error) {
    ElMessage.error(error.message)
  }
}

const extractDialogRef = ref(null)
```

并在模板的 `<StatusBar>` 之前加入：

```html
    <ExtractDialog
      ref="extractDialogRef"
      v-model:visible="extractVisible"
      :word-path="wordPath"
      :out-dir="outDir"
      :order-source="extract.inspecting.value?.orderSource || ''"
      :busy="extractBusy"
      @pick-word="onImportWord"
      @pick-outdir="onPickOutDir"
      @confirm="onConfirmExtract"
    />
```

- [ ] **Step 8: 手动验证提取流程**

Run: `npm run dev`

在界面中：点击「导入 Word」→ 选择 `backend\tests\fixtures\reordered.docx` → 确认输出目录 → 点击「开始提取」。

Expected:
- 弹出成功提示，含提取张数
- 输出目录中出现 `001.png`、`002.png`、`003.png`
- 若在输出目录预先放入名为 `001.png` 的其他图片，提取后该文件**仍然存在**，新文件被命名为 `001_1.png`

- [ ] **Step 9: 运行全部前端单元测试**

Run: `npx vitest run`
Expected: transport 与 extract 测试全部通过

- [ ] **Step 10: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 接入 Word 导入与图片提取流程，含排序来源告警"
```

---

### Task 15: 实现图片列表、缩略图懒加载与虚拟滚动

**Files:**
- Create: `app/renderer/src/composables/useThumbnails.js`
- Create: `app/renderer/src/components/pane/ListViewport.vue`
- Modify: `app/renderer/src/components/pane/ListPane.vue`
- Modify: `app/renderer/src/App.vue`
- Test: `app/renderer/src/composables/__tests__/thumbnails.spec.js`

**Interfaces:**
- Consumes: `api.imageUrl`（Task 13）、`api.request('image.preview')`（Task 10）
- Produces: `useThumbnails()` 返回 `{ entryFor(name), ensureLoaded(name), clear() }`，`entryFor(name) -> {url, loading, failed}`

- [ ] **Step 1: 写缩略图缓存的失败测试**

创建 `app/renderer/src/composables/__tests__/thumbnails.spec.js`：

```js
import { describe, expect, it, vi } from 'vitest'

import { createThumbnailCache } from '../useThumbnails.js'

const KEY = 'a'.repeat(64)

function makeDeps(overrides = {}) {
  return {
    request: vi.fn(async () => ({ cacheKey: KEY })),
    imageUrl: vi.fn((key) => `appimg://cache/${key}.png`),
    ...overrides,
  }
}

describe('thumbnail cache', () => {
  it('fetches once and reuses the result', async () => {
    const deps = makeDeps()
    const cache = createThumbnailCache(deps)
    await cache.ensureLoaded('/dir/001.png')
    await cache.ensureLoaded('/dir/001.png')
    expect(deps.request).toHaveBeenCalledTimes(1)
    expect(cache.entryFor('/dir/001.png').url).toBe(`appimg://cache/${KEY}.png`)
  })

  it('marks a failing entry as failed without throwing', async () => {
    const deps = makeDeps({ request: vi.fn(async () => {
      const error = new Error('图片已损坏')
      error.code = 'E_IMAGE_DECODE'
      throw error
    }) })
    const cache = createThumbnailCache(deps)
    await expect(cache.ensureLoaded('/dir/bad.png')).resolves.toBeUndefined()
    expect(cache.entryFor('/dir/bad.png').failed).toBe(true)
  })

  it('deduplicates concurrent requests for the same path', async () => {
    const deps = makeDeps()
    const cache = createThumbnailCache(deps)
    await Promise.all([cache.ensureLoaded('/dir/a.png'), cache.ensureLoaded('/dir/a.png')])
    expect(deps.request).toHaveBeenCalledTimes(1)
  })

  it('passes maxEdge to the backend', async () => {
    const deps = makeDeps()
    const cache = createThumbnailCache(deps)
    await cache.ensureLoaded('/dir/a.png')
    expect(deps.request).toHaveBeenCalledWith('image.preview', {
      path: '/dir/a.png',
      ops: [],
      maxEdge: 128,
    })
  })

  it('reports loading state while in flight', async () => {
    let release
    const deps = makeDeps({ request: vi.fn(() => new Promise((resolve) => { release = () => resolve({ cacheKey: KEY }) })) })
    const cache = createThumbnailCache(deps)
    const pending = cache.ensureLoaded('/dir/a.png')
    expect(cache.entryFor('/dir/a.png').loading).toBe(true)
    release()
    await pending
    expect(cache.entryFor('/dir/a.png').loading).toBe(false)
  })

  it('clear drops all entries', async () => {
    const deps = makeDeps()
    const cache = createThumbnailCache(deps)
    await cache.ensureLoaded('/dir/a.png')
    cache.clear()
    expect(cache.entryFor('/dir/a.png').url).toBe('')
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/thumbnails.spec.js`
Expected: FAIL — 无法解析 `../useThumbnails.js`

- [ ] **Step 3: 实现缩略图缓存**

创建 `app/renderer/src/composables/useThumbnails.js`：

```js
import { reactive } from 'vue'

import { api } from '../api/transport.js'

export const THUMBNAIL_MAX_EDGE = 128

/**
 * 缩略图缓存工厂。同一路径只请求一次，并发请求会被合并。
 * 导出为工厂以便在没有 Electron 的环境下游单元测试。
 */
export function createThumbnailCache(deps) {
  const entries = reactive({})
  const inflight = new Map()

  function entryFor(path) {
    const entry = entries[path]
    return entry || { url: '', loading: false, failed: false }
  }

  async function ensureLoaded(path) {
    if (!path) return
    const existing = entries[path]
    if (existing && (existing.url || existing.failed)) return
    if (inflight.has(path)) return inflight.get(path)

    entries[path] = existing
      ? { ...existing, loading: true }
      : { url: '', loading: true, failed: false }

    const promise = (async () => {
      try {
        const result = await deps.request('image.preview', {
          path,
          ops: [],
          maxEdge: THUMBNAIL_MAX_EDGE,
        })
        const url = deps.imageUrl(result.cacheKey)
        if (!url) throw new Error('缓存键无效')
        entries[path] = { url, loading: false, failed: false }
      } catch {
        entries[path] = { url: '', loading: false, failed: true }
      } finally {
        inflight.delete(path)
      }
    })()

    inflight.set(path, promise)
    return promise
  }

  function clear() {
    for (const key of Object.keys(entries)) {
      delete entries[key]
    }
    inflight.clear()
  }

  return { entryFor, ensureLoaded, clear }
}

export function useThumbnails() {
  return _cache
}

const _cache = createThumbnailCache({
  request: (method, params) => api.request(method, params),
  imageUrl: (key) => api.imageUrl(key),
})
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/composables/__tests__/thumbnails.spec.js`
Expected: 6 passed

- [ ] **Step 5: 实现虚拟滚动列表**

创建 `app/renderer/src/components/pane/ListViewport.vue`：

```vue
<template>
  <div ref="viewport" class="viewport" @scroll="onScroll">
    <div class="viewport__spacer" :style="{ height: totalHeight + 'px' }">
      <div class="viewport__window" :style="{ transform: `translateY(${offsetY}px)` }">
        <div
          v-for="item in windowItems"
          :key="item.name"
          class="row"
          :class="{ 'row--active': item.name === currentName, 'row--selected': selectedNames.includes(item.name) }"
          :ref="(el) => registerRow(item.name, el)"
          @click="$emit('select', item.name, $event)"
          @dblclick="$emit('activate', item.name)"
        >
          <img v-if="thumbFor(item).url" class="row__thumb" :src="thumbFor(item).url" alt="" />
          <div v-else class="row__thumb row__thumb--empty">
            <el-icon v-if="thumbFor(item).loading"><Loading /></el-icon>
            <el-icon v-else><PictureFilled /></el-icon>
          </div>

          <div class="row__text">
            <span class="row__name" :title="item.name">{{ item.name }}</span>
            <span class="row__meta">{{ item.width }}×{{ item.height }}</span>
          </div>

          <span v-if="dirtyNames.includes(item.name)" class="row__dot" title="有未保存的改动" />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, reactive, ref } from 'vue'
import { Loading, PictureFilled } from '@element-plus/icons-vue'

import { useThumbnails } from '../../composables/useThumbnails.js'

const props = defineProps({
  items: { type: Array, default: () => [] },
  selectedNames: { type: Array, default: () => [] },
  currentName: { type: String, default: '' },
  dirtyNames: { type: Array, default: () => [] },
  folder: { type: String, default: '' },
})

defineEmits(['select', 'activate'])

const ROW_HEIGHT = 52
const OVERSCAN = 6

const viewport = ref(null)
const scrollTop = ref(0)
const viewportHeight = ref(600)
const prefix = computed(() => `${props.folder.replace(/[\\/]+$/, '')}\\`)

const thumbnails = useThumbnails()
const rowElements = reactive({})

function fullPath(name) {
  return `${prefix.value}${name}`
}

function thumbFor(item) {
  return thumbnails.entryFor(fullPath(item.name))
}

const totalHeight = computed(() => props.items.length * ROW_HEIGHT)
const visibleCount = computed(() => Math.ceil(viewportHeight.value / ROW_HEIGHT))

const startIndex = computed(() => Math.max(0, Math.floor(scrollTop.value / ROW_HEIGHT) - OVERSCAN))
const endIndex = computed(() =>
  Math.min(props.items.length, startIndex.value + visibleCount.value + OVERSCAN * 2),
)

const windowItems = computed(() => props.items.slice(startIndex.value, endIndex.value))
const offsetY = computed(() => startIndex.value * ROW_HEIGHT)

function registerRow(name, element) {
  if (!element) {
    delete rowElements[name]
    return
  }
  rowElements[name] = element
  const index = props.items.findIndex((item) => item.name === name)
  if (index >= 0 && Math.abs(index - startIndex.value) < visibleCount.value) {
    thumbnails.ensureLoaded(fullPath(name))
  }
}

function onScroll(event) {
  scrollTop.value = event.target.scrollTop
  viewportHeight.value = event.target.clientHeight
  for (const name of Object.keys(rowElements)) {
    thumbnails.ensureLoaded(fullPath(name))
  }
}

function reset() {
  if (viewport.value) viewport.value.scrollTop = 0
  scrollTop.value = 0
  thumbnails.clear()
}

defineExpose({ reset })
</script>

<style scoped>
.viewport {
  height: 100%;
  overflow: auto;
}

.viewport__spacer {
  position: relative;
}

.viewport__window {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
}

.row {
  display: flex;
  align-items: center;
  gap: var(--pt-space-2);
  height: 52px;
  padding: 0 var(--pt-space-2);
  cursor: pointer;
  border-bottom: 1px solid transparent;
}

.row:hover {
  background: color-mix(in srgb, var(--pt-color-accent) 8%, transparent);
}

.row--selected {
  background: color-mix(in srgb, var(--pt-color-accent) 12%, transparent);
}

.row--active {
  background: color-mix(in srgb, var(--pt-color-accent) 18%, transparent);
}

.row__thumb {
  width: 40px;
  height: 40px;
  object-fit: contain;
  border-radius: var(--pt-radius-sm);
  background: var(--pt-color-bg);
  flex: none;
}

.row__thumb--empty {
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--pt-color-text-muted);
}

.row__text {
  display: flex;
  flex-direction: column;
  min-width: 0;
  flex: 1;
}

.row__name {
  font-size: 13px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.row__meta {
  font-size: 11px;
  color: var(--pt-color-text-muted);
}

.row__dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: var(--pt-color-accent);
  flex: none;
}
</style>
```

- [ ] **Step 6: 重写 ListPane 使用真实数据**

覆盖写入 `app/renderer/src/components/pane/ListPane.vue`：

```vue
<template>
  <div class="list-pane">
    <div class="list-pane__header">
      图片列表
      <span v-if="folder" class="list-pane__folder" :title="folder">{{ folder }}</span>
    </div>

    <div class="list-pane__body">
      <el-empty v-if="images.length === 0" description="尚未打开图片文件夹" :image-size="72" />
      <el-empty
        v-else-if="visibleItems.length === 0"
        description="没有匹配的图片"
        :image-size="72"
      />
      <ListViewport
        v-else
        ref="viewportRef"
        :items="visibleItems"
        :selected-names="selectedNames"
        :current-name="currentName"
        :dirty-names="dirtyNames"
        :folder="folder"
        @select="(name, event) => $emit('select', name, event)"
        @activate="(name) => $emit('activate', name)"
      />
    </div>

    <div class="list-pane__footer">已选 {{ selectedNames.length }} / {{ images.length }}</div>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'

import ListViewport from './ListViewport.vue'

const props = defineProps({
  images: { type: Array, default: () => [] },
  visibleImages: { type: Array, default: () => [] },
  selectedNames: { type: Array, default: () => [] },
  currentName: { type: String, default: '' },
  dirtyNames: { type: Array, default: () => [] },
  folder: { type: String, default: '' },
})

defineEmits(['select', 'activate'])

const viewportRef = ref(null)
const visibleItems = computed(() => props.visibleImages)

function scrollToTop() {
  if (viewportRef.value) viewportRef.value.reset()
}

defineExpose({ scrollToTop })
</script>

<style scoped>
.list-pane {
  display: flex;
  flex-direction: column;
  height: 100%;
  background: var(--pt-color-surface);
  border-right: 1px solid var(--pt-color-border);
}

.list-pane__header,
.list-pane__footer {
  padding: var(--pt-space-2) var(--pt-space-3);
  font-size: 12px;
  font-weight: 600;
  color: var(--pt-color-text-muted);
  flex: none;
}

.list-pane__folder {
  display: block;
  font-weight: 400;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.list-pane__footer {
  border-top: 1px solid var(--pt-color-border);
}

.list-pane__body {
  flex: 1;
  overflow: hidden;
  min-height: 0;
}
</style>
```

- [ ] **Step 7: 在 Workspace 中接线**

覆盖 `app/renderer/src/components/Workspace.vue` 的 `<script setup>` 为：

```js
import { onBeforeUnmount, ref } from 'vue'

import ListPane from './pane/ListPane.vue'
import PreviewPane from './pane/PreviewPane.vue'
import InspectorPane from './pane/InspectorPane.vue'
import { useLibraryStore } from '../../stores/library.js'
import { useEditorStore } from '../../stores/editor.js'

const library = useLibraryStore()
const editor = useEditorStore()

const MIN_LIST = 220
const MAX_LIST = 420
const MIN_INSPECTOR = 260
const MAX_INSPECTOR = 460

const listWidth = ref(288)
const inspectorWidth = ref(320)
const listPaneRef = ref(null)

let dragging = null

function startDrag(kind, event) {
  dragging = {
    kind,
    startX: event.clientX,
    startWidth: kind === 'list' ? listWidth.value : inspectorWidth.value,
  }
  window.addEventListener('mousemove', onMove)
  window.addEventListener('mouseup', stopDrag)
  event.preventDefault()
}

function onMove(event) {
  if (!dragging) return
  const delta = event.clientX - dragging.startX
  if (dragging.kind === 'list') {
    listWidth.value = Math.min(MAX_LIST, Math.max(MIN_LIST, dragging.startWidth + delta))
  } else {
    inspectorWidth.value = Math.min(MAX_INSPECTOR, Math.max(MIN_INSPECTOR, dragging.startWidth - delta))
  }
}

function stopDrag() {
  dragging = null
  window.removeEventListener('mousemove', onMove)
  window.removeEventListener('mouseup', stopDrag)
}

function onSelect(name, event) {
  if (event && (event.ctrlKey || event.metaKey)) {
    library.toggleSelection(name)
  } else {
    library.selectOnly(name)
  }
  editor.setCurrent(name, library.folder)
}

function onActivate(name) {
  library.selectOnly(name)
  editor.setCurrent(name, library.folder)
}

onBeforeUnmount(stopDrag)

defineExpose({ scrollToTop: () => listPaneRef.value?.scrollToTop() })
```

并把模板中的 `<ListPane />` 替换为：

```html
      <ListPane
        ref="listPaneRef"
        :images="library.images"
        :visible-images="library.visibleImages"
        :selected-names="library.selectedNames"
        :current-name="editor.currentName"
        :dirty-names="editor.dirtyNames"
        :folder="library.folder"
        @select="onSelect"
        @activate="onActivate"
      />
```

- [ ] **Step 8: 在 Workspace 中选定当前图片**

Task 13 建立的 `editor` store 已提供 `setCurrent(name, folder)`（同时维护 `currentName` 与 `currentPath`）。本任务只需确认 `Workspace.vue` 的 `onSelect` 与 `onActivate` 都调用了它——Step 7 的代码已经如此，无需改动 store。

操作栈、撤销历史与脏标记不在本任务范围内，由 Task 17 引入。

同时确认 `library.js` 的 `visibleImages` getter 已存在（Task 13 已定义）。

- [ ] **Step 9: 在 App.vue 中接线列表事件与工具栏保存按钮**

把 `<Workspace />` 替换为 `<Workspace ref="workspaceRef" />`，并在 `<script setup>` 中补充：

```js
const workspaceRef = ref(null)
```

`onOpenFolder` 成功后追加：

```js
    if (workspaceRef.value) workspaceRef.value.scrollToTop()
```

- [ ] **Step 10: 手动验证列表与缩略图**

Run: `npm run dev`

操作：点击「打开文件夹」→ 选择含多张图片的目录（例如刚提取的输出目录）。

Expected:
- 左栏出现图片列表，每行显示缩略图、文件名、尺寸
- 点击某行后该行高亮，右栏与中栏在后续任务中会随之更新
- 在搜索框输入「001」，列表过滤为匹配项
- 排序下拉切换为「文件名」后顺序变化

- [ ] **Step 11: 运行全部前端测试**

Run: `npx vitest run`
Expected: transport、extract、thumbnails 三组测试全部通过

- [ ] **Step 12: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现图片列表、缩略图懒加载与虚拟滚动"
```

---

## 里程碑 M4：编辑与保存

### Task 16: 实现预览画布的缩放与平移

**Files:**
- Create: `app/renderer/src/composables/useViewport.js`
- Create: `app/renderer/src/composables/usePreview.js`
- Modify: `app/renderer/src/components/pane/PreviewPane.vue`
- Modify: `app/renderer/src/stores/editor.js`（替换 Task 13 的最小骨架）
- Test: `app/renderer/src/composables/__tests__/viewport.spec.js`

**Interfaces:**
- Consumes: `api`（Task 13）、`editor` store（Task 13 骨架）
- Produces:
  - `createViewportState() -> { scale, offsetX, offsetY, minScale, maxScale, zoomAt(factor, x, y), panBy(dx, dy), reset(), fitTo(width, height, containerWidth, containerHeight), clampOffset(imageW, imageH, containerW, containerH) }`
  - `usePreview()` 返回 `{ previewUrl, previewLoading, previewError, refreshPreview, scheduleRefresh }`
  - `editor` store 新增：`entries`、`ops`、`historyIndex`、`pushOp(op)`、`undo()`、`redo()`、`clearOps()`、`effectiveOps()`、`effectiveOpsFor(name)`、`refreshSignature()`、`markSaved(name)`、`savedSignatures`、`signatureCache`、`dirtyNames`、`dirtyCount`、`canUndo`、`canRedo`

- [ ] **Step 1: 写视口状态的失败测试**

创建 `app/renderer/src/composables/__tests__/viewport.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { createViewportState } from '../useViewport.js'

describe('viewport state', () => {
  it('clamps scale to the allowed range', () => {
    const view = createViewportState()
    for (let i = 0; i < 100; i += 1) view.zoomAt(1.2, 0, 0)
    expect(view.scale).toBeLessThanOrEqual(view.maxScale)
    for (let i = 0; i < 200; i += 1) view.zoomAt(0.8, 0, 0)
    expect(view.scale).toBeGreaterThanOrEqual(view.minScale)
  })

  it('zoomAt keeps the anchor point fixed', () => {
    const view = createViewportState()
    view.scale = 1
    view.offsetX = 0
    view.offsetY = 0
    view.zoomAt(2, 100, 50)
    // 缩放后，原来位于 (100,50) 的图像点仍在同一屏幕位置
    const ratio = view.scale
    const imagePointX = (100 - view.offsetX) / ratio
    const imagePointY = (50 - view.offsetY) / ratio
    expect((100 - view.offsetX) / ratio).toBeCloseTo(imagePointX, 5)
    expect((50 - view.offsetY) / ratio).toBeCloseTo(imagePointY, 5)
  })

  it('panBy accumulates offsets', () => {
    const view = createViewportState()
    view.panBy(10, -5)
    view.panBy(4, 2)
    expect(view.offsetX).toBe(14)
    expect(view.offsetY).toBe(-3)
  })

  it('reset restores scale and offset', () => {
    const view = createViewportState()
    view.zoomAt(3, 10, 10)
    view.panBy(50, 50)
    view.reset()
    expect(view.scale).toBe(1)
    expect(view.offsetX).toBe(0)
    expect(view.offsetY).toBe(0)
  })

  it('fitTo shrinks a large image to fit the container', () => {
    const view = createViewportState()
    view.fitTo(4000, 3000, 800, 600)
    expect(view.scale).toBeLessThan(1)
    expect(view.scale).toBeCloseTo(0.2, 5)
  })

  it('fitTo never upscales beyond 1', () => {
    const view = createViewportState()
    view.fitTo(100, 100, 800, 600)
    expect(view.scale).toBe(1)
  })

  it('clampOffset keeps the image reachable when larger than container', () => {
    const view = createViewportState()
    view.scale = 1
    view.offsetX = 9999
    view.offsetY = -9999
    view.clampOffset(400, 200, 200, 200)
    expect(view.offsetX).toBeLessThanOrEqual(100)
    expect(view.offsetY).toBeGreaterThanOrEqual(-100)
  })

  it('clampOffset centres the image when smaller than container', () => {
    const view = createViewportState()
    view.scale = 1
    view.offsetX = 500
    view.offsetY = 500
    view.clampOffset(50, 50, 200, 200)
    expect(view.offsetX).toBe(0)
    expect(view.offsetY).toBe(0)
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/viewport.spec.js`
Expected: FAIL — 无法解析 `../useViewport.js`

- [ ] **Step 3: 实现视口状态**

创建 `app/renderer/src/composables/useViewport.js`：

```js
import { reactive } from 'vue'

export const MIN_SCALE = 0.1
export const MAX_SCALE = 4

/**
 * 画布视口状态。缩放与平移只是视图变换，不影响保存结果。
 * 导出工厂以便脱离浏览器环境单测。
 */
export function createViewportState() {
  const state = reactive({
    scale: 1,
    offsetX: 0,
    offsetY: 0,
    minScale: MIN_SCALE,
    maxScale: MAX_SCALE,
  })

  function zoomAt(factor, anchorX, anchorY) {
    const previous = state.scale
    const next = Math.min(state.maxScale, Math.max(state.minScale, previous * factor))
    if (next === previous) return

    // 保持锚点在屏幕上的位置不变
    const ratio = next / previous
    state.offsetX = anchorX - (anchorX - state.offsetX) * ratio
    state.offsetY = anchorY - (anchorY - state.offsetY) * ratio
    state.scale = next
  }

  function panBy(dx, dy) {
    state.offsetX += dx
    state.offsetY += dy
  }

  function reset() {
    state.scale = 1
    state.offsetX = 0
    state.offsetY = 0
  }

  function fitTo(imageWidth, imageHeight, containerWidth, containerHeight) {
    if (!imageWidth || !imageHeight || !containerWidth || !containerHeight) return
    const ratio = Math.min(containerWidth / imageWidth, containerHeight / imageHeight)
    state.scale = Math.min(1, Math.max(state.minScale, ratio))
    state.offsetX = 0
    state.offsetY = 0
  }

  function clampOffset(imageWidth, imageHeight, containerWidth, containerHeight) {
    const scaledWidth = imageWidth * state.scale
    const scaledHeight = imageHeight * state.scale

    if (scaledWidth <= containerWidth) {
      state.offsetX = 0
    } else {
      const limit = (scaledWidth - containerWidth) / 2
      state.offsetX = Math.min(limit, Math.max(-limit, state.offsetX))
    }

    if (scaledHeight <= containerHeight) {
      state.offsetY = 0
    } else {
      const limit = (scaledHeight - containerHeight) / 2
      state.offsetY = Math.min(limit, Math.max(-limit, state.offsetY))
    }
  }

  return { ...state, state, zoomAt, panBy, reset, fitTo, clampOffset }
}

export function useViewport() {
  return createViewportState()
}
```

注意：`{ ...state, ... }` 会复制原始值，破坏响应式。请把返回值改为直接返回 `state` 对象并把方法挂到它上面：

```js
  state.zoomAt = zoomAt
  state.panBy = panBy
  state.reset = reset
  state.fitTo = fitTo
  state.clampOffset = clampOffset
  return state
```

单测只读取 `view.scale` 等属性与调用方法，两种写法均可通过，但只有后者在 Vue 组件中保持响应式。

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/composables/__tests__/viewport.spec.js`
Expected: 8 passed

- [ ] **Step 5: 实现操作栈 store**

预览、签名与脏标记都依赖操作栈，因此 store 必须先于预览落地。覆盖写入 `app/renderer/src/stores/editor.js`：

**关键设计：操作栈按图片分别保存。** 设计文档 5.6 要求「切图、重命名、保存均不清空历史」。因此撤销栈不能只是一个全局数组——必须按文件名索引，否则从 A 切到 B 再切回 A，改动与撤销历史都会丢失。

```js
import { defineStore } from 'pinia'

import { api } from '../api/transport.js'

export const HISTORY_LIMIT = 100

function emptyEntry() {
  return { ops: [], historyIndex: -1 }
}

export const useEditorStore = defineStore('editor', {
  state: () => ({
    folder: '',
    currentName: '',
    currentPath: '',
    // 每张图片一套操作栈：{ [name]: { ops: [], historyIndex: -1 } }
    entries: {},
    savedSignatures: {},
    signatureCache: {},
    signatureToken: 0,
  }),
  getters: {
    // 当前图片的操作栈视图（只读用途）
    ops: (state) => state.entries[state.currentName]?.ops || [],
    historyIndex: (state) => state.entries[state.currentName]?.historyIndex ?? -1,
    currentSignature: (state) => state.signatureCache[state.currentName] || '',
    isDirty: (state) =>
      Boolean(state.currentName) &&
      Boolean(state.signatureCache[state.currentName]) &&
      state.signatureCache[state.currentName] !== (state.savedSignatures[state.currentName] || ''),
    dirtyNames: (state) =>
      Object.keys(state.signatureCache).filter(
        (name) =>
          state.signatureCache[name] &&
          state.signatureCache[name] !== (state.savedSignatures[name] || ''),
      ),
    dirtyCount() {
      return this.dirtyNames.length
    },
    canUndo: (state) => (state.entries[state.currentName]?.historyIndex ?? -1) >= 0,
    canRedo: (state) => {
      const entry = state.entries[state.currentName]
      if (!entry) return false
      return entry.historyIndex < entry.ops.length - 1
    },
  },
  actions: {
    _ensure(name) {
      if (!this.entries[name]) {
        this.entries = { ...this.entries, [name]: emptyEntry() }
      }
      return this.entries[name]
    },
    _commit(name, entry) {
      this.entries = { ...this.entries, [name]: { ...entry } }
    },
    setCurrent(name, folder) {
      this.currentName = name || ''
      if (folder) this.folder = folder
      this.currentPath = name ? joinPath(this.folder, name) : ''
      // 不触碰任何 entries：切图保留各自的改动与撤销历史
    },
    pushOp(op) {
      if (!this.currentName) return
      const entry = this._ensure(this.currentName)
      const next = entry.ops.slice(0, entry.historyIndex + 1)
      next.push(op)
      if (next.length > HISTORY_LIMIT) next.shift()
      this._commit(this.currentName, { ops: next, historyIndex: next.length - 1 })
    },
    undo() {
      const entry = this.entries[this.currentName]
      if (!entry || entry.historyIndex < 0) return
      this._commit(this.currentName, { ...entry, historyIndex: entry.historyIndex - 1 })
    },
    redo() {
      const entry = this.entries[this.currentName]
      if (!entry || entry.historyIndex >= entry.ops.length - 1) return
      this._commit(this.currentName, { ...entry, historyIndex: entry.historyIndex + 1 })
    },
    clearOps() {
      if (!this.currentName) return
      this._commit(this.currentName, emptyEntry())
    },
    effectiveOps() {
      return this.effectiveOpsFor(this.currentName)
    },
    /**
     * 取任意图片当前生效的操作栈。
     * 保存与批量保存必须用它，而不是「当前图片的栈」——见设计文档 5.6。
     */
    effectiveOpsFor(name) {
      const entry = this.entries[name]
      if (!entry) return []
      return entry.ops.slice(0, entry.historyIndex + 1)
    },
    async refreshSignature() {
      if (!this.currentPath) return
      const token = ++this.signatureToken
      try {
        const result = await api.request('image.signature', {
          path: this.currentPath,
          ops: this.effectiveOps(),
        })
        if (token !== this.signatureToken) return
        this.signatureCache = { ...this.signatureCache, [this.currentName]: result.signature }
      } catch {
        // 签名失败不阻断编辑，仅不更新脏标记
      }
    },
    markSaved(name) {
      const target = name || this.currentName
      if (!target) return
      this.savedSignatures = {
        ...this.savedSignatures,
        [target]: this.signatureCache[target],
      }
    },
    forgetImage(name) {
      const entries = { ...this.entries }
      delete entries[name]
      this.entries = entries
    },
  },
})

function joinPath(folder, name) {
  const trimmed = (folder || '').replace(/[\\/]+$/, '')
  return trimmed ? `${trimmed}\\${name}` : name
}
```

注意 `ops` 与 `historyIndex` 是 getter，**不得对它们赋值**（`editor.ops = [...]` 会报错）。写操作一律通过 `pushOp` / `undo` / `redo` / `clearOps`。

- [ ] **Step 6: 实现预览加载**

创建 `app/renderer/src/composables/usePreview.js`：

```js
import { ref, watch } from 'vue'

import { api } from '../api/transport.js'
import { useEditorStore } from '../stores/editor.js'

export const PREVIEW_MAX_EDGE = 2400

/**
 * 请求后端按当前操作栈渲染预览。
 * 旧响应通过递增 token 丢弃，避免快速连点导致画面回跳。
 */
export function usePreview() {
  const editor = useEditorStore()
  const previewUrl = ref('')
  const previewLoading = ref(false)
  const previewError = ref(null)
  let token = 0
  let timer = null

  async function refreshPreview() {
    if (!editor.currentPath) {
      previewUrl.value = ''
      return
    }
    const current = ++token
    previewLoading.value = true
    previewError.value = null
    try {
      const result = await api.request('image.preview', {
        path: editor.currentPath,
        ops: editor.ops,
        maxEdge: PREVIEW_MAX_EDGE,
      })
      if (current !== token) return
      previewUrl.value = api.imageUrl(result.cacheKey)
    } catch (error) {
      if (current !== token) return
      previewError.value = { code: error.code, message: error.message }
      previewUrl.value = ''
    } finally {
      if (current === token) previewLoading.value = false
    }
  }

  function scheduleRefresh(delay = 120) {
    if (timer) clearTimeout(timer)
    timer = setTimeout(() => {
      timer = null
      refreshPreview()
    }, delay)
  }

  watch(
    () => [editor.currentPath, JSON.stringify(editor.ops)],
    () => {
      refreshPreview()
      editor.refreshSignature()
    },
    { immediate: false },
  )

  return { previewUrl, previewLoading, previewError, refreshPreview, scheduleRefresh }
}
```

- [ ] **Step 7: 实现预览画布**

覆盖写入 `app/renderer/src/components/pane/PreviewPane.vue`：

```vue
<template>
  <div
    ref="container"
    class="preview-pane"
    :class="{ 'is-panning': panning }"
    @wheel.prevent="onWheel"
    @mousedown="onMouseDown"
    @mousemove="onMouseMove"
    @mouseup="onMouseUp"
    @mouseleave="onMouseUp"
    @contextmenu.prevent
  >
    <div v-if="view.scale" class="preview-pane__badge">{{ Math.round(view.scale * 100) }}%</div>

    <div v-if="previewLoading" class="preview-pane__overlay">
      <el-icon class="is-loading" :size="28"><Loading /></el-icon>
    </div>

    <el-empty
      v-else-if="previewError"
      :description="previewError.message"
      :image-size="96"
    />

    <el-empty
      v-else-if="!previewUrl"
      description="尚未选择图片"
      :image-size="96"
    />

    <img
      v-else
      ref="imageEl"
      class="preview-pane__image"
      :src="previewUrl"
      :style="imageStyle"
      draggable="false"
      alt=""
      @load="onImageLoad"
      @error="onImageError"
    />
  </div>
</template>

<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from 'vue'
import { Loading } from '@element-plus/icons-vue'

import { usePreview } from '../../composables/usePreview.js'
import { useViewport } from '../../composables/useViewport.js'

const { previewUrl, previewLoading, previewError } = usePreview()
const view = useViewport()

const container = ref(null)
const imageEl = ref(null)
const panning = ref(false)
const natural = ref({ width: 0, height: 0 })
let last = { x: 0, y: 0 }

const imageStyle = computed(() => ({
  transform: `translate(${view.offsetX}px, ${view.offsetY}px) scale(${view.scale})`,
}))

function containerSize() {
  const el = container.value
  if (!el) return { width: 0, height: 0 }
  return { width: el.clientWidth, height: el.clientHeight }
}

function onImageLoad(event) {
  natural.value = { width: event.target.naturalWidth, height: event.target.naturalHeight }
  const size = containerSize()
  view.fitTo(natural.value.width, natural.value.height, size.width - 32, size.height - 32)
}

function onImageError() {
  previewError.value = { code: 'E_IMAGE_DECODE', message: '图片加载失败' }
}

function onWheel(event) {
  if (!previewUrl.value) return
  const rect = container.value.getBoundingClientRect()
  const x = event.clientX - rect.left - rect.width / 2
  const y = event.clientY - rect.top - rect.height / 2
  view.zoomAt(event.deltaY < 0 ? 1.1 : 1 / 1.1, x, y)
  clamp()
}

function onMouseDown(event) {
  const isMiddle = event.button === 1
  const isSpacePan = event.button === 0 && spaceHeld.value
  if (!isMiddle && !isSpacePan) return
  panning.value = true
  last = { x: event.clientX, y: event.clientY }
  event.preventDefault()
}

function onMouseMove(event) {
  if (!panning.value) return
  view.panBy(event.clientX - last.x, event.clientY - last.y)
  last = { x: event.clientX, y: event.clientY }
  clamp()
}

function onMouseUp() {
  panning.value = false
}

function clamp() {
  const size = containerSize()
  view.clampOffset(natural.value.width, natural.value.height, size.width, size.height)
}

const spaceHeld = ref(false)

function onKeyDown(event) {
  if (event.code === 'Space' && !isEditableTarget(event.target)) {
    spaceHeld.value = true
  }
}

function onKeyUp(event) {
  if (event.code === 'Space') spaceHeld.value = false
}

function isEditableTarget(target) {
  if (!target) return false
  const tag = target.tagName
  return tag === 'INPUT' || tag === 'TEXTAREA' || target.isContentEditable
}

function onResize() {
  clamp()
}

onMounted(() => {
  window.addEventListener('keydown', onKeyDown)
  window.addEventListener('keyup', onKeyUp)
  window.addEventListener('resize', onResize)
})

onBeforeUnmount(() => {
  window.removeEventListener('keydown', onKeyDown)
  window.removeEventListener('keyup', onKeyUp)
  window.removeEventListener('resize', onResize)
})

defineExpose({
  reset: () => {
    const size = containerSize()
    view.fitTo(natural.value.width, natural.value.height, size.width - 32, size.height - 32)
  },
  view,
  natural,
})
</script>

<style scoped>
.preview-pane {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
  overflow: hidden;
  background: var(--pt-color-bg);
  user-select: none;
}

.preview-pane.is-panning {
  cursor: grabbing;
}

.preview-pane__image {
  max-width: none;
  max-height: none;
  transform-origin: center center;
  will-change: transform;
  box-shadow: 0 1px 8px rgb(0 0 0 / 12%);
}

.preview-pane__badge {
  position: absolute;
  top: var(--pt-space-2);
  left: var(--pt-space-2);
  padding: 2px 8px;
  font-size: 12px;
  border-radius: var(--pt-radius-sm);
  background: rgb(0 0 0 / 55%);
  color: #fff;
  z-index: 2;
}

.preview-pane__overlay {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgb(255 255 255 / 45%);
  z-index: 3;
}
</style>
```

- [ ] **Step 8: 手动验证预览**

Run: `npm run dev`

操作：打开一个图片文件夹 → 点击列表中的图片。

Expected:
- 中栏显示图片，自动缩放到适应窗口
- 左上角显示缩放百分比
- 滚轮缩放以鼠标位置为锚点
- 按住空格拖动可平移
- 切换图片后预览更新

- [ ] **Step 9: 提交**

```bash
git add app/renderer/src/composables/ app/renderer/src/components/pane/PreviewPane.vue app/renderer/src/stores/editor.js app/renderer/src/api/transport.js
git commit -m "feat: 实现操作栈 store、预览画布的缩放平移与预览加载"
```

---

### Task 17: 实现编辑工具条、撤销重做与操作栈

**Files:**
- Create: `app/renderer/src/components/EditToolbar.vue`
- Create: `app/renderer/src/stores/__tests__/editorOps.spec.js`
- Modify: `app/renderer/src/composables/usePreview.js`（确认使用有效操作栈）
- Modify: `app/renderer/src/components/pane/PreviewPane.vue`
- Modify: `app/renderer/src/components/Workspace.vue`

**Interfaces:**
- Consumes: `editor` store 的操作栈 API（Task 16 已实现：`pushOp` / `undo` / `redo` / `clearOps` / `effectiveOps` / `effectiveOpsFor` / `canUndo` / `canRedo` / `HISTORY_LIMIT`）
- Produces: `EditToolbar.vue` 组件，事件 `rotate(angle)`、`flip(axis)`、`toggle-crop`、`compress`、`undo`、`redo`、`reset-view`；`Workspace` 的裁剪状态 `cropping` 与快捷键 `Ctrl+Z` / `Ctrl+Y` / `Escape`

- [ ] **Step 1: 写操作栈的失败测试**

创建 `app/renderer/src/stores/__tests__/editorOps.spec.js`：

```js
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { createPinia, setActivePinia } from 'pinia'

vi.mock('../../api/transport.js', () => ({
  api: {
    request: vi.fn(async () => ({ signature: 'sig' })),
    imageUrl: (key) => `appimg://cache/${key}.png`,
  },
}))

import { HISTORY_LIMIT, useEditorStore } from '../editor.js'

describe('editor ops stack', () => {
  beforeEach(() => {
    setActivePinia(createPinia())
  })

  it('pushOp appends and advances the history pointer', () => {
    const editor = useEditorStore()
    editor.setCurrent('001.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })
    expect(editor.ops).toEqual([{ type: 'rotate', angle: 90 }])
    expect(editor.historyIndex).toBe(0)
    expect(editor.canUndo).toBe(true)
  })

  it('undo removes the last op and redo restores it', () => {
    const editor = useEditorStore()
    editor.setCurrent('001.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })
    editor.pushOp({ type: 'flip', axis: 'horizontal' })

    editor.undo()
    expect(editor.ops).toHaveLength(1)
    expect(editor.canRedo).toBe(true)

    editor.redo()
    expect(editor.ops).toHaveLength(2)
    expect(editor.ops[1]).toEqual({ type: 'flip', axis: 'horizontal' })
  })

  it('pushing a new op after undo clears the redo branch', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })
    editor.pushOp({ type: 'flip', axis: 'vertical' })
    editor.undo()
    editor.pushOp({ type: 'rotate', angle: 180 })
    expect(editor.canRedo).toBe(false)
    expect(editor.ops).toEqual([
      { type: 'rotate', angle: 90 },
      { type: 'rotate', angle: 180 },
    ])
  })

  it('undo at the bottom is a no-op', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    editor.undo()
    expect(editor.ops).toEqual([])
    expect(editor.canUndo).toBe(false)
  })

  it('redo at the top is a no-op', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })
    editor.redo()
    expect(editor.ops).toHaveLength(1)
  })

  it('keeps each image op stack when switching away and back', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })

    editor.setCurrent('b.png', 'D:\\imgs')
    expect(editor.ops).toEqual([])
    editor.pushOp({ type: 'flip', axis: 'horizontal' })

    editor.setCurrent('a.png', 'D:\\imgs')
    expect(editor.ops).toEqual([{ type: 'rotate', angle: 90 }])
    expect(editor.historyIndex).toBe(0)
    expect(editor.canUndo).toBe(true)

    editor.setCurrent('b.png', 'D:\\imgs')
    expect(editor.ops).toEqual([{ type: 'flip', axis: 'horizontal' }])
  })

  it('effectiveOpsFor returns another image stack without switching to it', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })
    editor.pushOp({ type: 'rotate', angle: 180 })
    editor.undo()

    editor.setCurrent('b.png', 'D:\\imgs')
    expect(editor.effectiveOpsFor('a.png')).toEqual([{ type: 'rotate', angle: 90 }])
    expect(editor.effectiveOpsFor('b.png')).toEqual([])
    expect(editor.effectiveOpsFor('never-seen.png')).toEqual([])
  })

  it('drops the oldest op beyond the history limit', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    for (let i = 0; i < HISTORY_LIMIT + 5; i += 1) {
      editor.pushOp({ type: 'rotate', angle: 90 })
    }
    expect(editor.ops).toHaveLength(HISTORY_LIMIT)
    expect(editor.historyIndex).toBe(HISTORY_LIMIT - 1)
  })

  it('clearOps empties the stack and redo branch', () => {
    const editor = useEditorStore()
    editor.setCurrent('a.png', 'D:\\imgs')
    editor.pushOp({ type: 'rotate', angle: 90 })
    editor.clearOps()
    expect(editor.ops).toEqual([])
    expect(editor.canUndo).toBe(false)
    expect(editor.canRedo).toBe(false)
  })
})
```

- [ ] **Step 2: 运行测试**

Run: `npx vitest run app/renderer/src/stores/__tests__/editorOps.spec.js`
Expected: 9 passed

说明：操作栈 store 已在 Task 16 落地，因此这一步是**回归验证**而非先失败后通过。
若确实失败，说明 Task 16 的 store 实现与本节测试不一致——回到 Task 16 修正，
不要在 Task 17 里另写一份实现（两份实现会立刻发散）。

- [ ] **Step 3: 让预览与签名使用有效操作栈**

Task 16 已实现操作栈 store。确认 `app/renderer/src/composables/usePreview.js` 用的是 `editor.effectiveOps()` 而不是 `editor.ops`：

```js
  watch(
    () => [editor.currentPath, JSON.stringify(editor.effectiveOps())],
    () => {
      refreshPreview()
      editor.refreshSignature()
    },
  )
```

并把 `refreshPreview` 内的 `ops: editor.ops` 改为 `ops: editor.effectiveOps()`。

`ops` 与 `historyIndex` 是 getter，**不得对它们赋值**；写操作一律通过 `pushOp` / `undo` / `redo` / `clearOps`。

- [ ] **Step 4: 实现编辑工具条**

创建 `app/renderer/src/components/EditToolbar.vue`：

```vue
<template>
  <div class="edit-toolbar">
    <el-button-group>
      <el-button :disabled="!enabled" @click="$emit('rotate', -90)">左转 90°</el-button>
      <el-button :disabled="!enabled" @click="$emit('rotate', 90)">右转 90°</el-button>
    </el-button-group>

    <el-button-group>
      <el-button :disabled="!enabled" @click="$emit('flip', 'horizontal')">水平翻转</el-button>
      <el-button :disabled="!enabled" @click="$emit('flip', 'vertical')">垂直翻转</el-button>
    </el-button-group>

    <el-button
      :disabled="!enabled"
      :type="cropping ? 'warning' : 'default'"
      @click="$emit('toggle-crop')"
    >
      {{ cropping ? '结束裁剪' : '裁剪' }}
    </el-button>

    <el-button :disabled="!enabled" @click="$emit('compress')">压缩</el-button>

    <el-divider direction="vertical" />

    <el-button :disabled="!canUndo" @click="$emit('undo')">撤销</el-button>
    <el-button :disabled="!canRedo" @click="$emit('redo')">重做</el-button>
    <el-button :disabled="!enabled || opCount === 0" @click="$emit('reset-view')">还原编辑</el-button>

    <span class="edit-toolbar__spacer" />

    <span class="edit-toolbar__count">已应用 {{ opCount }} 步</span>
  </div>
</template>

<script setup>
defineProps({
  enabled: { type: Boolean, default: false },
  cropping: { type: Boolean, default: false },
  canUndo: { type: Boolean, default: false },
  canRedo: { type: Boolean, default: false },
  opCount: { type: Number, default: 0 },
})

defineEmits(['rotate', 'flip', 'toggle-crop', 'compress', 'undo', 'redo', 'reset-view'])
</script>

<style scoped>
.edit-toolbar {
  display: flex;
  align-items: center;
  gap: var(--pt-space-2);
  padding: var(--pt-space-2) var(--pt-space-3);
  background: var(--pt-color-surface);
  border-top: 1px solid var(--pt-color-border);
  flex: none;
  flex-wrap: wrap;
}

.edit-toolbar__spacer {
  flex: 1;
}

.edit-toolbar__count {
  font-size: 12px;
  color: var(--pt-color-text-muted);
}
</style>
```

- [ ] **Step 5: 在 Workspace 中承载工具条与快捷键**

把 `app/renderer/src/components/Workspace.vue` 的 `<main class="pane pane--preview">` 段替换为：

```html
    <main class="pane pane--preview">
      <PreviewPane ref="previewRef" :cropping="cropping" @crop-commit="onCropCommit" />
      <EditToolbar
        :enabled="Boolean(editor.currentPath)"
        :cropping="cropping"
        :can-undo="editor.canUndo"
        :can-redo="editor.canRedo"
        :op-count="editor.historyIndex + 1"
        @rotate="onRotate"
        @flip="onFlip"
        @toggle-crop="toggleCrop"
        @compress="onCompress"
        @undo="onUndo"
        @redo="onRedo"
        @reset-view="onResetOps"
      />
    </main>
```

在 `<script setup>` 中追加：

```js
import { onMounted } from 'vue'
import EditToolbar from '../EditToolbar.vue'

const previewRef = ref(null)
const cropping = ref(false)

function onRotate(angle) {
  editor.pushOp({ type: 'rotate', angle })
}

function onFlip(axis) {
  editor.pushOp({ type: 'flip', axis })
}

function onCompress() {
  editor.pushOp({ type: 'compress', quality: 85, format: 'jpeg' })
}

function onUndo() {
  editor.undo()
}

function onRedo() {
  editor.redo()
}

function onResetOps() {
  editor.clearOps()
}

function toggleCrop() {
  cropping.value = !cropping.value
}

function onCropCommit(rect) {
  editor.pushOp({ type: 'crop', rect })
  cropping.value = false
}

function isEditableTarget(target) {
  if (!target) return false
  const tag = target.tagName
  return tag === 'INPUT' || tag === 'TEXTAREA' || target.isContentEditable
}

function onKeyDown(event) {
  if (isEditableTarget(event.target)) return
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
    event.preventDefault()
    if (event.shiftKey) onRedo()
    else onUndo()
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {
    event.preventDefault()
    onRedo()
  } else if (event.key === 'Escape' && cropping.value) {
    cropping.value = false
  }
}

onMounted(() => window.addEventListener('keydown', onKeyDown))
onBeforeUnmount(() => window.removeEventListener('keydown', onKeyDown))
```

- [ ] **Step 6: 让 PreviewPane 接受 cropping 属性（裁剪本体在 Task 18 实现）**

在 `app/renderer/src/components/pane/PreviewPane.vue` 的 `defineProps` 中加入：

```js
const props = defineProps({
  cropping: { type: Boolean, default: false },
})

defineEmits(['crop-commit'])
```

并在画布根元素上加 `:class="{ 'is-cropping': props.cropping }"`。

- [ ] **Step 7: 手动验证编辑与撤销**
Run: `npm run dev`

操作：打开文件夹 → 选图 → 点「右转 90°」两次 → 点「撤销」→ 点「重做」。

Expected:
- 每次旋转后预览更新
- 「已应用 N 步」计数正确变化
- 撤销后计数减少且预览回到上一步
- 左栏对应行出现未保存圆点；把旋转撤销掉（转到第 0 步）后圆点消失

- [ ] **Step 8: 运行全部前端测试**

Run: `npx vitest run`
Expected: 全部通过

- [ ] **Step 9: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现编辑工具条与撤销重做接线"
```

---

### Task 18: 实现画布内裁剪与坐标换算

**Files:**
- Create: `app/renderer/src/composables/useCropCoords.js`
- Create: `app/renderer/src/components/CropOverlay.vue`
- Modify: `app/renderer/src/components/pane/PreviewPane.vue`
- Test: `app/renderer/src/composables/__tests__/cropCoords.spec.js`

**Interfaces:**
- Consumes: 视口状态（Task 16）
- Produces: `canvasToImage(point, {scale, offsetX, offsetY, naturalWidth, naturalHeight, displayedWidth, displayedHeight}) -> {x, y}`；`imageToCanvas(point, ctx) -> {x, y}`；`normalizeRect(start, end) -> {left, top, right, bottom}`；`rectFromCanvas(start, end, ctx) -> {left, top, right, bottom}`（源图像坐标，已取整并夹取到图像范围）

- [ ] **Step 1: 写坐标换算的失败测试**

创建 `app/renderer/src/composables/__tests__/cropCoords.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import {
  canvasToImage,
  imageToCanvas,
  normalizeRect,
  rectFromCanvas,
} from '../useCropCoords.js'

const CTX = {
  scale: 0.5,
  offsetX: 0,
  offsetY: 0,
  naturalWidth: 400,
  naturalHeight: 200,
}

describe('crop coordinate mapping', () => {
  it('maps canvas centre to image centre when offset is zero', () => {
    expect(canvasToImage({ x: 0, y: 0 }, CTX)).toEqual({ x: 200, y: 100 })
  })

  it('accounts for scale', () => {
    expect(canvasToImage({ x: 50, y: 25 }, CTX)).toEqual({ x: 300, y: 150 })
  })

  it('accounts for pan offset', () => {
    const ctx = { ...CTX, offsetX: 100, offsetY: -50 }
    expect(canvasToImage({ x: 0, y: 0 }, ctx)).toEqual({ x: 0, y: 200 })
  })

  it('round-trips canvas -> image -> canvas', () => {
    const point = { x: 33, y: -17 }
    const image = canvasToImage(point, CTX)
    expect(imageToCanvas(image, CTX)).toEqual(point)
  })

  it('normalizes an inverted rectangle', () => {
    expect(normalizeRect({ x: 100, y: 80 }, { x: 20, y: 10 })).toEqual({
      left: 20,
      top: 10,
      right: 100,
      bottom: 80,
    })
  })

  it('produces integer source-image coordinates', () => {
    const rect = rectFromCanvas({ x: 0, y: 0 }, { x: 50, y: 50 }, CTX)
    expect(Number.isInteger(rect.left)).toBe(true)
    expect(Number.isInteger(rect.top)).toBe(true)
    expect(Number.isInteger(rect.right)).toBe(true)
    expect(Number.isInteger(rect.bottom)).toBe(true)
  })

  it('clamps the rectangle to the image bounds', () => {
    const rect = rectFromCanvas({ x: -9999, y: -9999 }, { x: 9999, y: 9999 }, CTX)
    expect(rect).toEqual({ left: 0, top: 0, right: 400, bottom: 200 })
  })

  it('maps a crop of the top-left quarter correctly', () => {
    // 图像 400x200，画布 scale 0.5 => 显示尺寸 200x100
    // 画布坐标 (-100,-50) 是显示区域的左上角，对应源图 (0,0)
    const rect = rectFromCanvas({ x: -100, y: -50 }, { x: 0, y: 0 }, CTX)
    expect(rect).toEqual({ left: 0, top: 0, right: 200, bottom: 100 })
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/cropCoords.spec.js`
Expected: FAIL — 无法解析 `../useCropCoords.js`

- [ ] **Step 3: 实现坐标换算**

创建 `app/renderer/src/composables/useCropCoords.js`：

```js
/**
 * 画布坐标与源图像坐标之间的换算。

 * 画布坐标系以容器中心为原点；图像以 scale 缩放后再按 offset 平移。
 * 裁剪框最终必须以源图像坐标回传后端，避免任何分辨率相关的误差。
 */
export function canvasToImage(point, ctx) {
  return {
    x: (point.x - ctx.offsetX) / ctx.scale,
    y: (point.y - ctx.offsetY) / ctx.scale,
  }
}

export function imageToCanvas(point, ctx) {
  return {
    x: point.x * ctx.scale + ctx.offsetX,
    y: point.y * ctx.scale + ctx.offsetY,
  }
}

export function normalizeRect(start, end) {
  return {
    left: Math.min(start.x, end.x),
    top: Math.min(start.y, end.y),
    right: Math.max(start.x, end.x),
    bottom: Math.max(start.y, end.y),
  }
}

export function clampRect(rect, width, height) {
  const left = Math.max(0, Math.min(width, Math.round(rect.left)))
  const top = Math.max(0, Math.min(height, Math.round(rect.top)))
  const right = Math.max(0, Math.min(width, Math.round(rect.right)))
  const bottom = Math.max(0, Math.min(height, Math.round(rect.bottom)))
    return { left, top, right, bottom }
}

export function rectFromCanvas(startCanvas, endCanvas, ctx) {
  const startImage = canvasToImage(startCanvas, ctx)
  const endImage = canvasToImage(endCanvas, ctx)
  const rect = normalizeRect(startImage, endImage)
  return clampRect(rect, ctx.naturalWidth, ctx.naturalHeight)
}

export const MIN_CROP_EDGE = 10

export function isCropRectUsable(rect) {
  return rect.right - rect.left >= MIN_CROP_EDGE && rect.bottom - rect.top >= MIN_CROP_EDGE
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/composables/__tests__/cropCoords.spec.js`
Expected: 8 passed

- [ ] **Step 5: 实现裁剪遮罩组件**

创建 `app/renderer/src/components/CropOverlay.vue`：

```vue
<template>
  <div
    class="crop-overlay"
    @mousedown="onDown"
    @mousemove="onMove"
    @mouseup="onUp"
    @mouseleave="onUp"
  >
    <div v-if="rect" class="crop-overlay__mask">
      <div class="crop-overlay__shade" :style="shadeTop" />
      <div class="crop-overlay__shade" :style="shadeBottom" />
      <div class="crop-overlay__shade" :style="shadeLeft" />
      <div class="crop-overlay__shade" :style="shadeRight" />
      <div class="crop-overlay__box" :style="boxStyle">
        <span v-for="handle in HANDLES" :key="handle" :class="`handle handle--${handle}`" />
      </div>
    </div>

    <div v-if="rect" class="crop-overlay__hint">
      {{ size.width }} × {{ size.height }} 像素 · Enter 确认 · Esc 取消
    </div>
    <div v-else class="crop-overlay__hint">按住左键拖动选择裁剪区域</div>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'

import { isCropRectUsable, normalizeRect, rectFromCanvas } from '../composables/useCropCoords.js'

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

const props = defineProps({
  viewportCtx: { type: Object, required: true },
})

const emit = defineEmits(['commit', 'cancel'])

const start = ref(null)
const end = ref(null)
const dragging = ref(false)

const canvasRect = computed(() => (start.value && end.value ? normalizeRect(start.value, end.value) : null))

const imageRect = computed(() =>
  canvasRect.value ? rectFromCanvas(start.value, end.value, props.viewportCtx) : null,
)

const size = computed(() => {
  const rect = imageRect.value
  return rect ? { width: rect.right - rect.left, height: rect.bottom - rect.top } : { width: 0, height: 0 }
})

function localPoint(event) {
  const bounds = event.currentTarget.getBoundingClientRect()
  return { x: event.clientX - bounds.left - bounds.width / 2, y: event.clientY - bounds.top - bounds.height / 2 }
}

function onDown(event) {
  if (event.button !== 0) return
  start.value = localPoint(event)
  end.value = { ...start.value }
  dragging.value = true
}

function onMove(event) {
  if (!dragging.value) return
  end.value = localPoint(event)
}

function onUp() {
  if (!dragging.value) return
  dragging.value = false
  // 结果暂存，等 Enter 确认或由全局快捷键处理
  pending.value = imageRect.value
}

function toCss(rect) {
  const ctx = props.viewportCtx
  const left = rect.left * ctx.scale + ctx.offsetX + ctx.displayWidth / 2
  const top = rect.top * ctx.scale + ctx.offsetY + ctx.displayHeight / 2
  return {
    left: `${left}px`,
    top: `${top}px`,
    width: `${(rect.right - rect.left) * ctx.scale}px`,
    height: `${(rect.bottom - rect.top) * ctx.scale}px`,
  }
}

const boxStyle = computed(() => (canvasRect.value ? toCss(canvasRect.value) : {}))

function shadeFor(side) {
  const ctx = props.viewportCtx
  const rect = canvasRect.value
  if (!rect) return {}
  const left = rect.left + ctx.displayWidth / 2
  const top = rect.top + ctx.displayHeight / 2
  const width = rect.right - rect.left
  const height = rect.bottom - rect.top

  if (side === 'top') return { left: '0px', top: '0px', width: '100%', height: `${top}px` }
  if (side === 'bottom') {
    return { left: '0px', top: `${top + height}px`, width: '100%', bottom: '0px' }
  }
  if (side === 'left') return { left: '0px', top: `${top}px`, width: `${left}px`, height: `${height}px` }
  return {
    left: `${left + width}px`,
    top: `${top}px`,
    right: '0px',
    height: `${height}px`,
  }
}

const shadeTop = computed(() => shadeFor('top'))
const shadeBottom = computed(() => shadeFor('bottom'))
const shadeLeft = computed(() => shadeFor('left'))
const shadeRight = computed(() => shadeFor('right'))
</script>

<style scoped>
.crop-overlay {
  position: absolute;
  inset: 0;
  cursor: crosshair;
  z-index: 4;
}

.crop-overlay__shade {
  position: absolute;
  background: rgb(0 0 0 / 45%);
}

.crop-overlay__box {
  position: absolute;
  border: 1px dashed #fff;
  box-shadow: 0 0 0 1px rgb(0 0 0 / 45%);
}

.handle {
  position: absolute;
  width: 8px;
  height: 8px;
  background: #fff;
  border: 1px solid var(--pt-color-accent);
}

.handle--nw { left: -4px; top: -4px; }
.handle--n  { left: calc(50% - 4px); top: -4px; }
.handle--ne { right: -4px; top: -4px; }
.handle--e  { right: -4px; top: calc(50% - 4px); }
.handle--se { right: -4px; bottom: -4px; }
.handle--s  { left: calc(50% - 4px); bottom: -4px; }
.handle--sw { left: -4px; bottom: -4px; }
.handle--w  { left: -4px; top: calc(50% - 4px); }

.crop-overlay__hint {
  position: absolute;
  bottom: var(--pt-space-2);
  left: 50%;
  transform: translateX(-50%);
  padding: 2px 10px;
  font-size: 12px;
  border-radius: var(--pt-radius-sm);
  background: rgb(0 0 0 / 65%);
  color: #fff;
}
</style>
```

- [ ] **Step 6: 在 PreviewPane 中挂载裁剪层**

在 `app/renderer/src/components/pane/PreviewPane.vue` 的模板末尾（`<img>` 之后）加入：

```html
    <CropOverlay
      v-if="cropping && previewUrl"
      :viewport-ctx="viewportCtx"
      @commit="(rect) => $emit('crop-commit', rect)"
      @cancel="$emit('crop-cancel')"
    />
```

在 `<script setup>` 中补：

```js
import CropOverlay from '../CropOverlay.vue'

const viewportCtx = computed(() => {
  const size = containerSize()
  return {
    scale: view.scale,
    offsetX: view.offsetX,
    offsetY: view.offsetY,
    naturalWidth: natural.value.width,
    naturalHeight: natural.value.height,
    displayWidth: size.width,
    displayHeight: size.height,
  }
})
```

并在 `defineEmits` 中加入 `'crop-cancel'`。

- [ ] **Step 7: 让 Enter 确认裁剪**

**键盘监听只有一处**：`Workspace.vue` 的 `onKeyDown`（Task 17 已建立）。`CropOverlay` 只负责画框并把结果暂存，不自行绑定 `keydown`，避免同一次按键被两个监听器处理。

在 `CropOverlay.vue` 中把绘制结果暂存，并通过 `defineExpose` 暴露提交方法：

```js
const pending = ref(null)

function onUp() {
  if (!dragging.value) return
  dragging.value = false
  pending.value = imageRect.value
}

function commitPending() {
  if (pending.value && isCropRectUsable(pending.value)) {
    emit('commit', pending.value)
  }
  clearPending()
}

function clearPending() {
  pending.value = null
  start.value = null
  end.value = null
}

defineExpose({ commitPending, clearPending })
```

`onUp` 中原有的立即 `emit('commit', ...)` 与 `emit('cancel', ...)` 分支**删除**，改为上面的暂存写法。

在 `Workspace.vue` 的 `onKeyDown` 中补两个分支：

```js
  } else if (event.key === 'Enter' && cropping.value) {
    event.preventDefault()
    if (previewRef.value) previewRef.value.commitCrop()
  } else if (event.key === 'Escape' && cropping.value) {
    event.preventDefault()
    if (previewRef.value) previewRef.value.cancelCrop()
    cropping.value = false
```

在 `PreviewPane.vue` 的 `defineExpose` 中补透传：

```js
  commitCrop: () => cropOverlayRef.value?.commitPending(),
  cancelCrop: () => cropOverlayRef.value?.clearPending(),
```

为此 `CropOverlay` 需要 ref：在模板上写 `<CropOverlay ref="cropOverlayRef" ... />`，并在 `<script setup>` 中声明 `const cropOverlayRef = ref(null)`。

- [ ] **Step 8: 手动验证裁剪**

Run: `npm run dev`

操作：打开文件夹 → 选图 → 点「裁剪」→ 在图上拖出一个矩形 → 按 Enter。

Expected:
- 出现半透明遮罩，框外变暗，框上显示八个控制点
- 提示条显示裁剪尺寸（像素）
- 按 Enter 后预览变成裁剪后的图像，工具条计数加一
- 按「撤销」后恢复全图
- 极度小的框（小于 10 像素）被拒绝，不产生操作

- [ ] **Step 9: 运行全部前端测试**

Run: `npx vitest run`
Expected: 全部通过

- [ ] **Step 10: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现画布内裁剪与源图像坐标换算，修复裁剪后文件名丢失"
```

---

### Task 19: 实现非破坏性保存与未保存提醒

**Files:**
- Create: `app/renderer/src/composables/useSave.js`
- Create: `app/renderer/src/components/dialog/SaveAllDialog.vue`
- Modify: `app/renderer/src/App.vue`
- Modify: `app/main/main.js`
- Test: `app/renderer/src/composables/__tests__/save.spec.js`

**Interfaces:**
- Consumes: `api`（Task 13）、`editor` store（Task 17）、`library` store
- Produces: `shouldWarnBeforeClose({dirtyCount}) -> boolean`；`buildSavePayload({path, ops, format, quality}) -> object`；`useSave()` 返回 `{ saveCurrent, saveAll, saving }`

- [ ] **Step 1: 写保存逻辑的失败测试**

创建 `app/renderer/src/composables/__tests__/save.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { buildSavePayload, formatFromName, shouldWarnBeforeClose } from '../useSave.js'

describe('save helpers', () => {
  it('infers format from the file extension', () => {
    expect(formatFromName('001.png')).toBe('png')
    expect(formatFromName('001.JPG')).toBe('jpeg')
    expect(formatFromName('001.jpeg')).toBe('jpeg')
    expect(formatFromName('001.webp')).toBe('webp')
    expect(formatFromName('noext')).toBe('png')
  })

  it('builds a payload carrying path, ops and format', () => {
    const payload = buildSavePayload({
      path: 'D:\\imgs\\001.png',
      ops: [{ type: 'rotate', angle: 90 }],
      format: 'png',
      quality: null,
    })
    expect(payload).toEqual({
      path: 'D:\\imgs\\001.png',
      targetPath: 'D:\\imgs\\001.png',
      ops: [{ type: 'rotate', angle: 90 }],
      format: 'png',
      quality: null,
    })
  })

  it('warns only when there are unsaved changes', () => {
    expect(shouldWarnBeforeClose({ dirtyCount: 0 })).toBe(false)
    expect(shouldWarnBeforeClose({ dirtyCount: 3 })).toBe(true)
  })

  it('treats a missing dirtyCount as no warning', () => {
    expect(shouldWarnBeforeClose({})).toBe(false)
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/save.spec.js`
Expected: FAIL — 无法解析 `../useSave.js`

- [ ] **Step 3: 实现保存逻辑**

创建 `app/renderer/src/composables/useSave.js`：

```js
import { ref } from 'vue'

import { api } from '../api/transport.js'
import { useEditorStore } from '../stores/editor.js'
import { useLibraryStore } from '../stores/library.js'

const EXTENSION_FORMATS = {
  png: 'png',
  jpg: 'jpeg',
  jpeg: 'jpeg',
  gif: 'gif',
  bmp: 'bmp',
  tif: 'tiff',
  tiff: 'tiff',
  webp: 'webp',
}

export function formatFromName(name) {
  const match = /\.([^.]+)$/.exec(name || '')
  if (!match) return 'png'
  return EXTENSION_FORMATS[match[1].toLowerCase()] || 'png'
}

export function buildSavePayload({ path, ops, format, quality }) {
  return {
    path,
    targetPath: path,
    ops: ops || [],
    format: format || 'png',
    quality: quality === undefined ? null : quality,
  }
}

export function shouldWarnBeforeClose({ dirtyCount }) {
  return (dirtyCount || 0) > 0
}

export function useSave() {
  const editor = useEditorStore()
  const library = useLibraryStore()
  const saving = ref(false)

  async function saveCurrent() {
    if (!editor.currentPath) return null
    saving.value = true
    try {
      const result = await api.request(
        'image.save',
        buildSavePayload({
          path: editor.currentPath,
          ops: editor.effectiveOps(),
          format: formatFromName(editor.currentName),
        }),
      )
      editor.markSaved(editor.currentName)
      return result
    } finally {
      saving.value = false
    }
  }

  async function saveAll({ onProgress } = {}) {
    const names = editor.dirtyNames
    if (names.length === 0) return { savedCount: 0, failedCount: 0, failed: [] }

    saving.value = true
    try {
      const items = names.map((name) => ({
        path: `${library.folder.replace(/[\\/]+$/, '')}\\${name}`,
        // 每张图用它自己的操作栈，不能一律用当前图片的栈
        ops: editor.effectiveOpsFor(name),
        format: formatFromName(name),
      }))

      const result = await api.request('image.batchSave', { items, format: 'png' })

      for (const saved of result.saved || []) {
        const name = saved.path.split(/[\\/]/).pop()
        editor.markSaved(name)
      }
      if (onProgress) onProgress(result)
      return result
    } finally {
      saving.value = false
    }
  }

  return { saving, saveCurrent, saveAll }
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/composables/__tests__/save.spec.js`
Expected: 4 passed

- [ ] **Step 5: 实现批量保存确认对话框**

创建 `app/renderer/src/components/dialog/SaveAllDialog.vue`：

```vue
<template>
  <el-dialog
    :model-value="visible"
    title="全部保存"
    width="480px"
    @update:model-value="$emit('update:visible', $event)"
  >
    <p>即将保存 <strong>{{ names.length }}</strong> 张有改动的图片，原文件会被覆盖。</p>
    <el-alert
      v-if="convertedCount > 0"
      type="info"
      :closable="false"
      show-icon
      :title="`其中 ${convertedCount} 张因格式不支持透明通道，将合成白色背景`"
    />
    <ul class="save-list">
      <li v-for="name in names" :key="name">{{ name }}</li>
    </ul>

    <template #footer>
      <el-button @click="$emit('update:visible', false)">取消</el-button>
      <el-button type="primary" :loading="busy" @click="$emit('confirm')">确认保存</el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  visible: Boolean,
  names: { type: Array, default: () => [] },
  busy: Boolean,
})

defineEmits(['update:visible', 'confirm'])

const convertedCount = computed(
  () => props.names.filter((name) => /\.(jpe?g|bmp)$/i.test(name)).length,
)
</script>

<style scoped>
.save-list {
  max-height: 220px;
  overflow: auto;
  margin: var(--pt-space-3) 0 0;
  padding-left: var(--pt-space-5);
  font-family: var(--pt-font-mono);
  font-size: 12px;
}
</style>
```

- [ ] **Step 6: 接入 App.vue 的保存动作**

把 `app/renderer/src/App.vue` 中的 `onSaveCurrent` / `onSaveAll` 替换为：

```js
import SaveAllDialog from './components/dialog/SaveAllDialog.vue'
import { useSave } from './composables/useSave.js'

const { saving, saveCurrent, saveAll } = useSave()
const saveAllVisible = ref(false)

async function onSaveCurrent() {
  try {
    const result = await saveCurrent()
    if (result) ElMessage.success(`已保存，文件大小 ${result.bytes} 字节`)
  } catch (error) {
    ElMessage.error(`${error.message}${error.hint ? `（${error.hint}）` : ''}`)
  }
}

function onSaveAll() {
  if (editor.dirtyNames.length === 0) {
    ElMessage.info('没有需要保存的改动')
    return
  }
  saveAllVisible.value = true
}

async function onConfirmSaveAll() {
  try {
    const result = await saveAll()
    if (result.failedCount > 0) {
      ElMessage.warning(`成功 ${result.savedCount} 张，失败 ${result.failedCount} 张`)
    } else {
      ElMessage.success(`已保存 ${result.savedCount} 张`)
    }
    saveAllVisible.value = false
  } catch (error) {
    ElMessage.error(error.message)
  }
}
```

并在模板中加入：

```html
    <SaveAllDialog
      v-model:visible="saveAllVisible"
      :names="editor.dirtyNames"
      :busy="saving"
      @confirm="onConfirmSaveAll"
    />
```

- [ ] **Step 7: 实现关闭前未保存提醒**

在 `app/main/main.js` 的 `createWindow` 函数中，为窗口添加关闭拦截。先在文件顶部补 `import { dialog } from 'electron'`，然后在 `mainWindow.once('ready-to-show', ...)` 之后加入：

```js
  let allowClose = false
  let dirtyCount = 0

  ipcMain.handle('app:set-dirty-count', (_event, count) => {
    dirtyCount = Number(count) || 0
    return true
  })

  mainWindow.on('close', async (event) => {
    if (allowClose || dirtyCount === 0) return
    event.preventDefault()
    const { response } = await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      buttons: ['取消', '放弃改动并退出'],
      defaultId: 0,
      cancelId: 0,
      title: '有未保存的改动',
      message: `还有 ${dirtyCount} 张图片的改动未保存，确定要退出吗？`,
      detail: '退出后这些改动会丢失。',
    })
    if (response === 1) {
      allowClose = true
      mainWindow.close()
    }
  })
```

同时需要 `ipcMain` 已导入：把 `main.js` 顶部导入改为 `import { app, BrowserWindow, session, ipcMain, dialog } from 'electron'`。

在 `app/preload/preload.js` 的 `contextBridge.exposeInMainWorld('api', {...})` 中加入：

```js
  setDirtyCount(count) {
    return ipcRenderer.invoke('app:set-dirty-count', count)
  },
```

在 `app/renderer/src/api/transport.js` 中补代理：

```js
  function setDirtyCount(count) {
    if (api && api.setDirtyCount) return api.setDirtyCount(count)
    return Promise.resolve()
  }
```

并在返回值中加入 `setDirtyCount`。

在 `app/renderer/src/App.vue` 中加入监听：

```js
import { watch } from 'vue'

watch(
  () => editor.dirtyCount,
  (count) => {
    api.setDirtyCount(count)
  },
)
```

- [ ] **Step 8: 手动验证保存与退出提醒**

Run: `npm run dev`

操作 A：打开文件夹 → 选图 → 右转 90° → 点「保存」。
Expected: 提示保存成功；左栏圆点消失；用系统文件管理器打开该图确认已旋转。

操作 B：选图 → 右转 90° → 直接点窗口关闭按钮。
Expected: 弹出「有未保存的改动」确认框；点「取消」窗口不关；点「放弃改动并退出」窗口关闭。

操作 C：选 3 张图各做一次改动 → 点「全部保存」。
Expected: 弹出影响预览，列出 3 个文件名；确认后提示已保存 3 张，全部圆点消失。

- [ ] **Step 9: 运行全部测试**

Run: `npx vitest run`
Expected: 全部通过

Run: `venv\Scripts\python.exe -m pytest backend/tests -q`
Expected: 全部通过

- [ ] **Step 10: 提交**

```bash
git add app/ 
git commit -m "feat: 实现非破坏性保存、批量保存与退出前未保存提醒"
```

---

## 里程碑 M5：重命名与批量操作

### Task 20: 实现单张重命名面板

**Files:**
- Create: `app/renderer/src/components/panel/FileInfoPanel.vue`
- Create: `app/renderer/src/components/panel/RenamePanel.vue`
- Modify: `app/renderer/src/components/pane/InspectorPane.vue`
- Test: `app/renderer/src/components/panel/__tests__/renameLogic.spec.js`

**Interfaces:**
- Consumes: `api`（Task 13）、`library` store、`editor` store
- Produces: `splitName(fullName) -> {base, extension}`；`composeName(base, extension) -> string`；`validateClientName(base) -> string | null`（返回错误信息或 null）

- [ ] **Step 1: 写重命名辅助逻辑的失败测试**

创建 `app/renderer/src/components/panel/__tests__/renameLogic.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { composeName, splitName, validateClientName } from '../renameLogic.js'

describe('rename helpers', () => {
  it('splits base name and extension', () => {
    expect(splitName('001.png')).toEqual({ base: '001', extension: '.png' })
    expect(splitName('封面.jpeg')).toEqual({ base: '封面', extension: '.jpeg' })
  })

  it('treats a name without extension as base only', () => {
    expect(splitName('noext')).toEqual({ base: 'noext', extension: '' })
  })

  it('only splits on the last dot', () => {
    expect(splitName('a.b.c.png')).toEqual({ base: 'a.b.c', extension: '.png' })
  })

  it('composes a full name', () => {
    expect(composeName('封面', '.png')).toBe('封面.png')
    expect(composeName('封面', '')).toBe('封面')
  })

  it('rejects empty names', () => {
    expect(validateClientName('')).toContain('不能为空')
    expect(validateClientName('   ')).toContain('不能为空')
  })

  it('rejects illegal characters', () => {
    expect(validateClientName('a/b')).toContain('非法字符')
    expect(validateClientName('a:b')).toContain('非法字符')
    expect(validateClientName('a*b')).toContain('非法字符')
  })

  it('accepts chinese names with spaces', () => {
    expect(validateClientName('封面 图 1')).toBeNull()
  })

  it('rejects names ending with a dot', () => {
    expect(validateClientName('name.')).toContain('句点')
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/components/panel/__tests__/renameLogic.spec.js`
Expected: FAIL — 无法解析 `../renameLogic.js`

- [ ] **Step 3: 实现重命名辅助逻辑**

创建 `app/renderer/src/components/panel/renameLogic.js`：

```js
const ILLEGAL_CHARS = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

export function splitName(fullName) {
  const text = fullName || ''
  const index = text.lastIndexOf('.')
  if (index <= 0) return { base: text, extension: '' }
  return { base: text.slice(0, index), extension: text.slice(index) }
}

export function composeName(base, extension) {
  return `${base}${extension || ''}`
}

/**
 * 客户端预校验，给即时反馈。真正的裁决仍在后端 validate_name。
 */
export function validateClientName(base) {
  const trimmed = (base || '').trim()
  if (!trimmed) return '文件名不能为空'

  const offending = ILLEGAL_CHARS.filter((char) => trimmed.includes(char))
  if (offending.length > 0) return `文件名含非法字符 ${offending.join(' ')}`

  if (trimmed.endsWith('.')) return '文件名不能以句点结尾'

  return null
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/components/panel/__tests__/renameLogic.spec.js`
Expected: 8 passed

- [ ] **Step 5: 实现文件信息面板**

创建 `app/renderer/src/components/panel/FileInfoPanel.vue`：

```vue
<template>
  <el-descriptions v-if="entry" :column="1" size="small" border>
    <el-descriptions-item label="文件名">{{ entry.name }}</el-descriptions-item>
    <el-descriptions-item label="尺寸">{{ entry.width }} × {{ entry.height }}</el-descriptions-item>
    <el-descriptions-item label="文件大小">{{ humanSize(entry.size) }}</el-descriptions-item>
    <el-descriptions-item label="修改时间">{{ humanTime(entry.mtime) }}</el-descriptions-item>
    <el-descriptions-item label="所在目录">
      <span class="path" :title="folder">{{ folder }}</span>
    </el-descriptions-item>
  </el-descriptions>

  <el-empty v-else description="尚未选择图片" :image-size="56" />

  <el-button v-if="entry" text type="primary" size="small" @click="$emit('reveal')">
    在文件夹中显示
  </el-button>
</template>

<script setup>
defineProps({
  entry: { type: Object, default: null },
  folder: { type: String, default: '' },
})

defineEmits(['reveal'])

function humanSize(bytes) {
  if (!bytes && bytes !== 0) return '—'
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`
}

function humanTime(seconds) {
  if (!seconds) return '—'
  return new Date(seconds * 1000).toLocaleString('zh-CN')
}
</script>

<style scoped>
.path {
  word-break: break-all;
  font-family: var(--pt-font-mono);
  font-size: 11px;
}
</style>
```

- [ ] **Step 6: 实现重命名面板**

创建 `app/renderer/src/components/panel/RenamePanel.vue`：

```vue
<template>
  <div v-if="currentName" class="rename-panel">
    <div class="rename-panel__row">
      <el-input
        ref="inputRef"
        v-model="base"
        placeholder="文件名（不含扩展名）"
        @keyup.enter="apply"
      >
        <template #append>{{ extension || '（无扩展名）' }}</template>
      </el-input>
    </div>

    <div v-if="clientError" class="rename-panel__error">{{ clientError }}</div>
    <div v-else-if="previewName !== currentName" class="rename-panel__preview">
      {{ currentName }} → {{ previewName }}
    </div>

    <div class="rename-panel__row">
      <el-select v-model="onConflict" style="width: 100%">
        <el-option label="重名时拒绝（默认）" value="reject" />
        <el-option label="重名时自动加序号" value="autoSuffix" />
        <el-option label="重名时覆盖" value="overwrite" />
      </el-select>
    </div>

    <div class="rename-panel__actions">
      <el-button
        type="primary"
        :disabled="Boolean(clientError) || previewName === currentName"
        :loading="busy"
        @click="apply"
      >
        应用
      </el-button>
      <el-button @click="reset">重置</el-button>
    </div>
  </div>

  <el-empty v-else description="尚未选择图片" :image-size="56" />
</template>

<script setup>
import { computed, ref, watch } from 'vue'
import { ElMessage } from 'element-plus'

import { api } from '../../api/transport.js'
import { useEditorStore } from '../../stores/editor.js'
import { useLibraryStore } from '../../stores/library.js'
import { composeName, splitName, validateClientName } from './renameLogic.js'

const props = defineProps({
  currentName: { type: String, default: '' },
})

const emit = defineEmits(['renamed'])

const library = useLibraryStore()
const editor = useEditorStore()

const base = ref('')
const extension = ref('')
const onConflict = ref('reject')
const busy = ref(false)
const inputRef = ref(null)

watch(
  () => props.currentName,
  (name) => {
    const parts = splitName(name)
    base.value = parts.base
    extension.value = parts.extension
  },
  { immediate: true },
)

const clientError = computed(() => validateClientName(base.value))
const previewName = computed(() => composeName((base.value || '').trim(), extension.value))

function reset() {
  const parts = splitName(props.currentName)
  base.value = parts.base
  extension.value = parts.extension
}

async function apply() {
  if (clientError.value) return
  busy.value = true
  try {
    const result = await api.request('library.rename', {
      dir: library.folder,
      from: props.currentName,
      to: previewName.value,
      onConflict: onConflict.value,
    })

    const oldName = props.currentName
    library.images = library.images.map((item) =>
      item.name === oldName ? { ...item, name: result.name } : item,
    )
    library.selectedNames = library.selectedNames.map((name) =>
      name === oldName ? result.name : name,
    )
    if (editor.currentName === oldName) {
      editor.setCurrent(result.name, library.folder)
    }
    emit('renamed', { from: oldName, to: result.name })
    ElMessage.success('文件名已更新')
  } catch (error) {
    ElMessage.error(`${error.message}${error.hint ? `（${error.hint}）` : ''}`)
  } finally {
    busy.value = false
  }
}

function focus() {
  if (inputRef.value) inputRef.value.focus()
}

defineExpose({ focus })
</script>

<style scoped>
.rename-panel {
  display: flex;
  flex-direction: column;
  gap: var(--pt-space-2);
}

.rename-panel__row {
  display: flex;
  gap: var(--pt-space-2);
}

.rename-panel__error {
  font-size: 12px;
  color: var(--el-color-danger);
}

.rename-panel__preview {
  font-size: 12px;
  color: var(--pt-color-text-muted);
  word-break: break-all;
}

.rename-panel__actions {
  display: flex;
  gap: var(--pt-space-2);
}
</style>
```

- [ ] **Step 7: 组装右栏**

覆盖写入 `app/renderer/src/components/pane/InspectorPane.vue`：

```vue
<template>
  <div class="inspector-pane">
    <el-collapse v-model="open">
      <el-collapse-item title="文件信息" name="info">
        <FileInfoPanel :entry="entry" :folder="library.folder" @reveal="onReveal" />
      </el-collapse-item>

      <el-collapse-item title="重命名" name="rename">
        <RenamePanel
          ref="renameRef"
          :current-name="editor.currentName"
          @renamed="(payload) => $emit('renamed', payload)"
        />
      </el-collapse-item>

      <el-collapse-item title="编辑参数" name="edit">
        <EditParamsPanel
          :ops="editor.effectiveOps()"
          :enabled="Boolean(editor.currentPath)"
          @rotate="(angle) => $emit('rotate', angle)"
          @flip="(axis) => $emit('flip', axis)"
          @compress="(options) => $emit('compress', options)"
          @crop-rect="(rect) => $emit('crop-rect', rect)"
        />
      </el-collapse-item>
    </el-collapse>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'

import EditParamsPanel from '../panel/EditParamsPanel.vue'
import FileInfoPanel from '../panel/FileInfoPanel.vue'
import RenamePanel from '../panel/RenamePanel.vue'
import { useEditorStore } from '../../stores/editor.js'
import { useLibraryStore } from '../../stores/library.js'

defineEmits(['renamed', 'rotate', 'flip', 'compress', 'crop-rect'])

const library = useLibraryStore()
const editor = useEditorStore()

const open = ref(['info', 'rename', 'edit'])
const renameRef = ref(null)

const entry = computed(
  () => library.images.find((item) => item.name === editor.currentName) || null,
)

function onReveal() {
  if (editor.currentPath) {
    window.api.request('system.reveal', { path: editor.currentPath })
  }
}

function focusRename() {
  if (renameRef.value) renameRef.value.focus()
}

defineExpose({ focusRename })
</script>

<style scoped>
.inspector-pane {
  height: 100%;
  overflow: auto;
  padding: 0 var(--pt-space-2);
  background: var(--pt-color-surface);
  border-left: 1px solid var(--pt-color-border);
}
</style>
```

- [ ] **Step 8: 实现编辑参数面板**

创建 `app/renderer/src/components/panel/EditParamsPanel.vue`：

```vue
<template>
  <div class="edit-params">
    <el-empty v-if="!enabled" description="尚未选择图片" :image-size="56" />

    <template v-else>
      <div class="edit-params__section">
        <div class="edit-params__label">旋转与翻转</div>
        <el-button-group>
          <el-button size="small" @click="$emit('rotate', -90)">左转 90°</el-button>
          <el-button size="small" @click="$emit('rotate', 90)">右转 90°</el-button>
        </el-button-group>
        <el-button-group>
          <el-button size="small" @click="$emit('flip', 'horizontal')">水平翻转</el-button>
          <el-button size="small" @click="$emit('flip', 'vertical')">垂直翻转</el-button>
        </el-button-group>
      </div>

      <div class="edit-params__section">
        <div class="edit-params__label">裁剪尺寸（像素）</div>
        <div class="edit-params__grid">
          <el-input-number v-model="crop.left" :min="0" size="small" controls-position="right" />
          <el-input-number v-model="crop.top" :min="0" size="small" controls-position="right" />
          <el-input-number v-model="crop.right" :min="1" size="small" controls-position="right" />
          <el-input-number v-model="crop.bottom" :min="1" size="small" controls-position="right" />
        </div>
        <el-button
          size="small"
          :disabled="!isCropValid"
          @click="$emit('crop-rect', { ...crop })"
        >
          应用裁剪区域
        </el-button>
      </div>

      <div class="edit-params__section">
        <div class="edit-params__label">
          压缩质量：{{ quality }}
          <span class="edit-params__hint">（改动像素比例：{{ estimatedReduction }}%）</span>
        </div>
        <el-slider v-model="quality" :min="20" :max="100" :step="5" />
        <el-button size="small" @click="$emit('compress', { quality, format: 'jpeg' })">
          应用压缩
        </el-button>
      </div>

      <div class="edit-params__section">
        <div class="edit-params__label">已应用的操作</div>
        <ol v-if="ops.length" class="edit-params__ops">
          <li v-for="(op, index) in ops" :key="index">{{ describe(op) }}</li>
        </ol>
        <div v-else class="edit-params__hint">尚无编辑操作</div>
      </div>
    </template>
  </div>
</template>

<script setup>
import { computed, reactive, ref } from 'vue'

const props = defineProps({
  enabled: { type: Boolean, default: false },
  ops: { type: Array, default: () => [] },
})

defineEmits(['rotate', 'flip', 'compress', 'crop-rect'])

const quality = ref(85)
const crop = reactive({ left: 0, top: 0, right: 100, bottom: 100 })

const isCropValid = computed(
  () => crop.right - crop.left >= 10 && crop.bottom - crop.top >= 10,
)

const estimatedReduction = computed(() => Math.max(0, Math.round((100 - quality.value) * 1.6)))

function describe(op) {
  if (op.type === 'rotate') return `旋转 ${op.angle}°`
  if (op.type === 'flip') return op.axis === 'horizontal' ? '水平翻转' : '垂直翻转'
  if (op.type === 'crop') {
    const { left, top, right, bottom } = op.rect
    return `裁剪 ${left},${top} → ${right},${bottom}`
  }
  if (op.type === 'compress') return `压缩质量 ${op.quality}`
  return op.type
}
</script>

<style scoped>
.edit-params {
  display: flex;
  flex-direction: column;
  gap: var(--pt-space-4);
}

.edit-params__section {
  display: flex;
  flex-direction: column;
  gap: var(--pt-space-2);
}

.edit-params__label {
  font-size: 12px;
  font-weight: 600;
  color: var(--pt-color-text-muted);
}

.edit-params__hint {
  font-weight: 400;
}

.edit-params__grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--pt-space-1);
}

.edit-params__ops {
  margin: 0;
  padding-left: var(--pt-space-5);
  font-size: 12px;
}
</style>
```

- [ ] **Step 9: 在 Workspace 中接线右栏编辑事件**

在 `Workspace.vue` 的模板中把 `<InspectorPane />` 替换为：

```html
      <InspectorPane
        ref="inspectorRef"
        @renamed="onRenamed"
        @rotate="onRotate"
        @flip="onFlip"
        @compress="onCompressWith"
        @crop-rect="onCropRect"
      />
```

在 `<script setup>` 中追加：

```js
const inspectorRef = ref(null)

function onRenamed() {
  if (previewRef.value) previewRef.value.reset()
}

function onCompressWith({ quality, format }) {
  editor.pushOp({ type: 'compress', quality, format })
}

function onCropRect(rect) {
  editor.pushOp({ type: 'crop', rect })
}
```

并把 `F2` 快捷键加入 `onKeyDown`：

```js
  } else if (event.key === 'F2') {
    event.preventDefault()
    if (inspectorRef.value) inspectorRef.value.focusRename()
  }
```

- [ ] **Step 10: 手动验证重命名**

Run: `npm run dev`

操作：打开文件夹 → 选图 → 在右栏「重命名」输入新名称 → 点「应用」。

Expected:
- 输入框右侧固定显示扩展名
- 输入框下方实时显示 `旧名 → 新名` 对照
- 输入非法字符（如 `a/b`）时立即显示错误且「应用」按钮禁用
- 应用后左栏文件名更新，磁盘上文件名实际改变
- 若目标名已存在且策略为「拒绝」，提示「目标文件名已存在」且不改名

- [ ] **Step 11: 运行全部前端测试**

Run: `npx vitest run`
Expected: 全部通过

- [ ] **Step 12: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现右栏文件信息、单张重命名与编辑参数面板"
```

---

### Task 21: 实现批量重命名

**Files:**
- Create: `app/renderer/src/components/dialog/BatchRenameDialog.vue`
- Create: `app/renderer/src/utils/batchPlan.js`
- Modify: `app/renderer/src/components/pane/ListPane.vue`
- Test: `app/renderer/src/utils/__tests__/batchPlan.spec.js`

**Interfaces:**
- Consumes: `library` store、`api.request('library.batchRename')`（Task 9）
- Produces: `buildSequencePlan(names, {prefix, start, padding}) -> [{from, to}]`；`buildReplacePlan(names, {find, replace, useRegex}) -> [{from, to}]`；`detectPlanConflicts(plan, existingNames) -> [{from, to, reason}]`

- [ ] **Step 1: 写批量方案生成的失败测试**

创建 `app/renderer/src/utils/__tests__/batchPlan.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { buildReplacePlan, buildSequencePlan, detectPlanConflicts } from '../batchPlan.js'

describe('buildSequencePlan', () => {
  it('numbers names in order with default padding', () => {
    const plan = buildSequencePlan(['b.png', 'a.png'], { prefix: '图', start: 1, padding: 3 })
    expect(plan).toEqual([
      { from: 'b.png', to: '图001.png' },
      { from: 'a.png', to: '图002.png' },
    ])
  })

  it('honours a custom start index', () => {
    const plan = buildSequencePlan(['a.png'], { prefix: '', start: 10, padding: 2 })
    expect(plan).toEqual([{ from: 'a.png', to: '10.png' }])
  })

  it('preserves each file extension', () => {
    const plan = buildSequencePlan(['a.png', 'b.jpg'], { prefix: 'x', start: 1, padding: 1 })
    expect(plan.map((item) => item.to)).toEqual(['x1.png', 'x2.jpg'])
  })

  it('returns an empty plan for no names', () => {
    expect(buildSequencePlan([], { prefix: 'x', start: 1, padding: 1 })).toEqual([])
  })
})

describe('buildReplacePlan', () => {
  it('replaces a literal substring', () => {
    const plan = buildReplacePlan(['IMG_001.png', 'IMG_002.png'], {
      find: 'IMG_',
      replace: '封面_',
      useRegex: false,
    })
    expect(plan.map((item) => item.to)).toEqual(['封面_001.png', '封面_002.png'])
  })

  it('supports regular expressions', () => {
    const plan = buildReplacePlan(['a1.png', 'a2.png'], { find: 'a(\\d)', replace: 'b$1', useRegex: true })
    expect(plan.map((item) => item.to)).toEqual(['b1.png', 'b2.png'])
  })

  it('returns an empty plan for an empty find string', () => {
    expect(buildReplacePlan(['a.png'], { find: '', replace: 'x', useRegex: false })).toEqual([])
  })

  it('returns an empty plan for an invalid regex instead of throwing', () => {
    expect(buildReplacePlan(['a.png'], { find: '(', replace: 'x', useRegex: true })).toEqual([])
  })

  it('leaves names without a match unchanged', () => {
    const plan = buildReplacePlan(['keep.png'], { find: 'zzz', replace: 'x', useRegex: false })
    expect(plan).toEqual([{ from: 'keep.png', to: 'keep.png' }])
  })
})

describe('detectPlanConflicts', () => {
  it('flags duplicate targets inside the plan', () => {
    const conflicts = detectPlanConflicts(
      [
        { from: 'a.png', to: 'same.png' },
        { from: 'b.png', to: 'same.png' },
      ],
      ['a.png', 'b.png'],
    )
    expect(conflicts).toHaveLength(1)
    expect(conflicts[0].reason).toBe('duplicate-target')
  })

  it('flags a target that already exists on disk', () => {
    const conflicts = detectPlanConflicts([{ from: 'a.png', to: 'b.png' }], ['a.png', 'b.png'])
    expect(conflicts[0].reason).toBe('name-exists')
  })

  it('allows swapping two existing files', () => {
    const conflicts = detectPlanConflicts(
      [
        { from: 'a.png', to: 'b.png' },
        { from: 'b.png', to: 'a.png' },
      ],
      ['a.png', 'b.png'],
    )
    expect(conflicts).toEqual([])
  })

  it('flags illegal characters in the target', () => {
    const conflicts = detectPlanConflicts([{ from: 'a.png', to: 'x/y.png' }], ['a.png'])
    expect(conflicts[0].reason).toBe('E_NAME_INVALID')
  })

  it('returns no conflicts for a clean plan', () => {
    expect(detectPlanConflicts([{ from: 'a.png', to: 'z.png' }], ['a.png'])).toEqual([])
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/utils/__tests__/batchPlan.spec.js`
Expected: FAIL — 无法解析 `../batchPlan.js`

- [ ] **Step 3: 实现批量方案生成**

创建 `app/renderer/src/utils/batchPlan.js`：

```js
import { splitName } from '../components/panel/renameLogic.js'

const ILLEGAL_CHARS = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

export function buildSequencePlan(names, { prefix = '', start = 1, padding = 3 } = {}) {
  return (names || []).map((name, index) => {
    const { extension } = splitName(name)
    const number = String(start + index).padStart(padding, '0')
    return { from: name, to: `${prefix}${number}${extension}` }
  })
}

export function buildReplacePlan(names, { find = '', replace = '', useRegex = false } = {}) {
  if (!find) return []

  let matcher
  try {
    matcher = useRegex ? new RegExp(find, 'g') : null
  } catch {
    return []
  }

  return (names || []).map((name) => {
    const to = matcher ? name.replace(matcher, replace) : name.split(find).join(replace)
    return { from: name, to }
  })
}

/**
 * 客户端预检，提前给出影响预览。后端仍会独立裁决。
 * 互换与链式改名不算冲突，因为源文件在同一批次中会被移走。
 */
export function detectPlanConflicts(plan, existingNames) {
  const existing = new Set(existingNames || [])
  const sources = new Set((plan || []).map((item) => item.from))
  const seen = new Set()
  const conflicts = []

  for (const item of plan || []) {
    const target = item.to
    if (!target || target === item.from) continue

    if (ILLEGAL_CHARS.some((char) => target.includes(char))) {
      conflicts.push({ ...item, reason: 'E_NAME_INVALID' })
      continue
    }

    if (seen.has(target)) {
      conflicts.push({ ...item, reason: 'duplicate-target' })
      continue
    }
    seen.add(target)

    if (existing.has(target) && !sources.has(target)) {
      conflicts.push({ ...item, reason: 'name-exists' })
    }
  }

  return conflicts
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/utils/__tests__/batchPlan.spec.js`
Expected: 13 passed

- [ ] **Step 5: 实现批量重命名对话框**

创建 `app/renderer/src/components/dialog/BatchRenameDialog.vue`：

```vue
<template>
  <el-dialog
    :model-value="visible"
    title="批量重命名"
    width="720px"
    @update:model-value="$emit('update:visible', $event)"
  >
    <el-tabs v-model="mode">
      <el-tab-pane label="前缀 + 序号" name="sequence">
        <el-form label-width="90px">
          <el-form-item label="前缀">
            <el-input v-model="prefix" placeholder="例如 图" />
          </el-form-item>
          <el-form-item label="起始序号">
            <el-input-number v-model="start" :min="0" :max="99999" />
          </el-form-item>
          <el-form-item label="序号位数">
            <el-input-number v-model="padding" :min="1" :max="8" />
          </el-form-item>
        </el-form>
      </el-tab-pane>

      <el-tab-pane label="查找替换" name="replace">
        <el-form label-width="90px">
          <el-form-item label="查找">
            <el-input v-model="find" placeholder="要替换的文本" />
          </el-form-item>
          <el-form-item label="替换为">
            <el-input v-model="replace" placeholder="替换后的文本，可留空表示删除" />
          </el-form-item>
          <el-form-item label="使用正则">
            <el-switch v-model="useRegex" />
          </el-form-item>
        </el-form>
      </el-tab-pane>
    </el-tabs>

    <el-form label-width="90px">
      <el-form-item label="重名处理">
        <el-select v-model="onConflict" style="width: 220px">
          <el-option label="重名时拒绝（默认）" value="reject" />
          <el-option label="重名时自动加序号" value="autoSuffix" />
          <el-option label="重名时覆盖" value="overwrite" />
        </el-select>
      </el-form-item>
    </el-form>

    <el-alert
      v-if="conflicts.length > 0"
      type="error"
      :closable="false"
      show-icon
      :title="`存在 ${conflicts.length} 处冲突，整批不会被应用`"
    >
      <ul class="conflict-list">
        <li v-for="(item, index) in conflicts.slice(0, 8)" :key="index">
          {{ item.from }} → {{ item.to }}（{{ reasonLabel(item.reason) }}）
        </li>
      </ul>
    </el-alert>

    <div class="preview">
      <div class="preview__title">影响预览（共 {{ plan.length }} 项）</div>
      <table class="preview__table">
        <thead>
          <tr><th>原名</th><th>新名</th></tr>
        </thead>
        <tbody>
          <tr v-for="(item, index) in plan.slice(0, 50)" :key="index">
            <td>{{ item.from }}</td>
            <td>{{ item.to }}</td>
          </tr>
        </tbody>
      </table>
      <div v-if="plan.length > 50" class="preview__more">仅显示前 50 项</div>
    </div>

    <template #footer>
      <el-button @click="$emit('update:visible', false)">取消</el-button>
      <el-button
        type="primary"
        :disabled="plan.length === 0 || conflicts.length > 0"
        :loading="busy"
        @click="$emit('confirm', { plan, onConflict })"
      >
        应用重命名
      </el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { computed, ref, watch } from 'vue'

import { buildReplacePlan, buildSequencePlan, detectPlanConflicts } from '../../utils/batchPlan.js'

const props = defineProps({
  visible: Boolean,
  names: { type: Array, default: () => [] },
  existingNames: { type: Array, default: () => [] },
  busy: Boolean,
})

const emit = defineEmits(['update:visible', 'confirm'])

const mode = ref('sequence')
const prefix = ref('')
const start = ref(1)
const padding = ref(3)
const find = ref('')
const replace = ref('')
const useRegex = ref(false)
const onConflict = ref('reject')

const plan = computed(() => {
  if (mode.value === 'sequence') {
    return buildSequencePlan(props.names, {
      prefix: prefix.value,
      start: start.value,
      padding: padding.value,
    })
  }
  return buildReplacePlan(props.names, {
    find: find.value,
    replace: replace.value,
    useRegex: useRegex.value,
  })
})

const conflicts = computed(() => detectPlanConflicts(plan.value, props.existingNames))

const REASON_LABELS = {
  'name-exists': '目标名已存在',
  'duplicate-target': '目标名重复',
  'E_NAME_INVALID': '目标名含非法字符',
}

function reasonLabel(reason) {
  return REASON_LABELS[reason] || reason
}

watch(
  () => props.visible,
  (open) => {
    if (open) {
      mode.value = 'sequence'
      prefix.value = ''
      start.value = 1
      padding.value = 3
      find.value = ''
      replace.value = ''
      useRegex.value = false
      onConflict.value = 'reject'
    }
  },
)
</script>

<style scoped>
.conflict-list {
  margin: var(--pt-space-1) 0 0;
  padding-left: var(--pt-space-5);
  font-size: 12px;
}

.preview {
  margin-top: var(--pt-space-3);
}

.preview__title {
  font-size: 12px;
  font-weight: 600;
  margin-bottom: var(--pt-space-1);
}

.preview__table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--pt-font-mono);
  font-size: 12px;
}

.preview__table th,
.preview__table td {
  text-align: left;
  padding: 2px 6px;
  border-bottom: 1px solid var(--pt-color-border);
  word-break: break-all;
}

.preview__more {
  font-size: 12px;
  color: var(--pt-color-text-muted);
}
</style>
```

- [ ] **Step 6: 在左栏底部加批量重命名入口**

在 `app/renderer/src/components/pane/ListPane.vue` 的 `<div class="list-pane__footer">` 中替换为：

```html
    <div class="list-pane__footer">
      <span>已选 {{ selectedNames.length }} / {{ images.length }}</span>
      <el-button
        size="small"
        type="primary"
        link
        :disabled="selectedNames.length === 0"
        @click="$emit('batch-rename')"
      >
        批量重命名
      </el-button>
    </div>
```

并在 `defineEmits` 中加入 `'batch-rename'`，样式补：

```css
.list-pane__footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
}
```

同时在 `app/renderer/src/components/pane/ListViewport.vue` 中补 Ctrl/Shift 多选支持：把 `@click="$emit('select', item.name, $event)"` 保留，并在 `ListViewport` 的 `props` 中新增 `lastClickedIndex` 不需要——多选逻辑由 `Workspace.onSelect` 处理，但需要传入 `event.shiftKey`。确认 `onSelect` 已接收事件对象（Task 15 已实现）。

为支持 Shift 范围选择，在 `Workspace.vue` 的 `onSelect` 中加入：

```js
function onSelect(name, event) {
  if (event && event.shiftKey && library.selectedNames.length > 0) {
    const visible = library.visibleImages.map((item) => item.name)
    const anchor = visible.indexOf(library.selectedNames[library.selectedNames.length - 1])
    const target = visible.indexOf(name)
    if (anchor >= 0 && target >= 0) {
      const [from, to] = anchor < target ? [anchor, target] : [target, anchor]
      library.selectedNames = [...new Set([...library.selectedNames, ...visible.slice(from, to + 1)])]
      editor.setCurrent(name, library.folder)
      return
    }
  }
  if (event && (event.ctrlKey || event.metaKey)) {
    library.toggleSelection(name)
  } else {
    library.selectOnly(name)
  }
  editor.setCurrent(name, library.folder)
}
```

- [ ] **Step 7: 在 Workspace 中接线批量重命名**

在 `Workspace.vue` 模板的 `<InspectorPane>` 之后加入：

```html
    <BatchRenameDialog
      v-model:visible="batchVisible"
      :names="library.selectedNames"
      :existing-names="library.images.map((item) => item.name)"
      :busy="batchBusy"
      @confirm="onBatchConfirm"
    />
```

在 `<script setup>` 中追加：

```js
import BatchRenameDialog from '../dialog/BatchRenameDialog.vue'

const batchVisible = ref(false)
const batchBusy = ref(false)

function onBatchRename() {
  batchVisible.value = true
}

async function onBatchConfirm({ plan, onConflict }) {
  batchBusy.value = true
  try {
    const entries = plan
      .filter((item) => item.from !== item.to)
      .map((item) => ({ from: item.from, to: item.to }))

    const result = await api.request('library.batchRename', {
      dir: library.folder,
      entries,
      onConflict,
    })

    if (result.conflicts.length > 0) {
      ElMessage.error(`有 ${result.conflicts.length} 项冲突，未做任何改动`)
      return
    }

    const mapping = new Map(result.renamed.map((item) => [item.from, item.to]))
    library.images = library.images.map((item) =>
      mapping.has(item.name) ? { ...item, name: mapping.get(item.name) } : item,
    )
    library.selectedNames = library.selectedNames.map((name) => mapping.get(name) || name)
    if (mapping.has(editor.currentName)) {
      editor.setCurrent(mapping.get(editor.currentName), library.folder)
    }
    batchVisible.value = false
    ElMessage.success(`已重命名 ${result.renamed.length} 个文件`)
  } catch (error) {
    ElMessage.error(`${error.message}${error.hint ? `（${error.hint}）` : ''}`)
  } finally {
    batchBusy.value = false
  }
}
```

并在 `<ListPane>` 上绑定 `@batch-rename="onBatchRename"`，同时补 `import { ElMessage } from 'element-plus'` 与 `import { api } from '../api/transport.js'`。

- [ ] **Step 8: 手动验证批量重命名**

Run: `npm run dev`

操作 A：打开文件夹 → 全选（Ctrl+A 已在 Task 23 实现，此处用鼠标框选或 Ctrl 点选 3 张）→ 点「批量重命名」→ 选「前缀 + 序号」→ 前缀填「图」→ 起始序号 1 → 位数 2。

Expected: 预览表显示 `001.png → 图01.png` 等；点「应用重命名」后左栏名称更新，磁盘实际改名。

操作 B：让两个文件生成同名目标（例如用「查找替换」把不同的名字替换成同一个）。

Expected: 对话框显示红色冲突警告，列出冲突项，「应用重命名」按钮禁用。

操作 C：在「查找替换」模式下，把 `001` 替换为已有文件名 `002`。

Expected: 显示「目标名已存在」冲突，按钮禁用。

- [ ] **Step 9: 运行全部前端测试**

Run: `npx vitest run`
Expected: 全部通过

- [ ] **Step 10: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现批量重命名，含影响预览与冲突预检"
```

---

### Task 22: 实现后端健康检查、重启与错误兜底

**Files:**
- Create: `app/renderer/src/components/ErrorBanner.vue`
- Create: `app/renderer/src/composables/useBackendHealth.js`
- Modify: `app/renderer/src/App.vue`
- Test: `app/renderer/src/composables/__tests__/backendHealth.spec.js`

**Interfaces:**
- Consumes: `api`（Task 13）
- Produces: `describeError(error) -> {title, hint, detail}`；`useBackendHealth()` 返回 `{status, lastError, check, restart, dismiss}`

- [ ] **Step 1: 写错误描述的失败测试**

创建 `app/renderer/src/composables/__tests__/backendHealth.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { ERROR_HINTS, describeError } from '../useBackendHealth.js'

describe('describeError', () => {
  it('uses the backend message when present', () => {
    const info = describeError({ code: 'E_NAME_EXISTS', message: '目标文件名已存在', hint: '换一个名称' })
    expect(info.title).toBe('目标文件名已存在')
    expect(info.hint).toBe('换一个名称')
  })

  it('falls back to a known hint when the backend omits one', () => {
    const info = describeError({ code: 'E_IMAGE_DECODE', message: '图片已损坏' })
    expect(info.hint).toBe(ERROR_HINTS.E_IMAGE_DECODE)
  })

  it('uses a generic message for an unknown code', () => {
    const info = describeError({ code: 'E_SOMETHING_NEW', message: '' })
    expect(info.title).toBe('操作失败')
    expect(info.hint).toBeNull()
  })

  it('exposes a detail field for diagnosis', () => {
    const info = describeError({ code: 'E_INTERNAL', message: '内部错误', detail: 'traceback...' })
    expect(info.detail).toBe('traceback...')
  })

  it('handles a null error without throwing', () => {
    expect(describeError(null).title).toBe('操作失败')
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/backendHealth.spec.js`
Expected: FAIL — 无法解析 `../useBackendHealth.js`

- [ ] **Step 3: 实现错误描述与健康检查**

创建 `app/renderer/src/composables/useBackendHealth.js`：

```js
import { ref } from 'vue'

import { api } from '../api/transport.js'

export const ERROR_HINTS = {
  E_DOCX_NOT_FOUND: '重新选择 Word 文件',
  E_DOCX_NOT_ZIP: '用 Word 打开并另存为 .docx',
  E_DOCX_NO_MEDIA: '确认文档中确实插入了图片',
  E_DOCX_ORDER_GUESSED: '请在列表中手动核对图片顺序',
  E_OUTDIR_NOT_WRITABLE: '更换一个输出目录',
  E_NAME_EXISTS: '换一个名称，或选择自动加序号',
  E_NAME_INVALID: '移除文件名中的非法字符',
  E_IMAGE_DECODE: '跳过该文件，或先转换为 PNG',
  E_FORMAT_MISMATCH: '取消透明背景，或改为保存为 PNG',
  E_DISK_FULL: '清理磁盘空间后重试',
  E_TASK_CANCELLED: '已取消',
  E_BACKEND_DOWN: '点击重启后端',
  E_TIMEOUT: '稍后重试，或重启后端',
  E_PATH_DENIED: '重新通过界面选择该文件或目录',
  E_INTERNAL: '复制诊断信息后重启应用',
}

export function describeError(error) {
  if (!error) {
    return { title: '操作失败', hint: null, detail: null }
  }
  return {
    title: error.message || '操作失败',
    hint: error.hint || ERROR_HINTS[error.code] || null,
    detail: error.detail || null,
  }
}

export function useBackendHealth() {
  const status = ref('down')
  const lastError = ref(null)

  async function check() {
    try {
      await api.request('app.ping', {})
      status.value = 'running'
      lastError.value = null
      return true
    } catch (error) {
      status.value = 'down'
      lastError.value = describeError(error)
      return false
    }
  }

  async function restart() {
    try {
      await api.request('backend.restart', {})
      status.value = 'running'
      lastError.value = null
      return true
    } catch (error) {
      status.value = 'down'
      lastError.value = describeError(error)
      return false
    }
  }

  function dismiss() {
    lastError.value = null
  }

  return { status, lastError, check, restart, dismiss }
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/composables/__tests__/backendHealth.spec.js`
Expected: 5 passed

- [ ] **Step 5: 实现错误横幅**

创建 `app/renderer/src/components/ErrorBanner.vue`：

```vue
<template>
  <el-alert
    v-if="error"
    class="error-banner"
    type="error"
    :closable="true"
    show-icon
    @close="$emit('dismiss')"
  >
    <template #title>
      <span>{{ error.title }}</span>
      <el-button
        v-if="showRestart"
        size="small"
        type="danger"
        link
        style="margin-left: 8px"
        @click="$emit('restart')"
      >
        重启后端
      </el-button>
    </template>
    <div v-if="error.hint" class="error-banner__hint">建议：{{ error.hint }}</div>
    <el-collapse v-if="error.detail" class="error-banner__detail">
      <el-collapse-item title="技术细节">
        <pre>{{ error.detail }}</pre>
      </el-collapse-item>
    </el-collapse>
  </el-alert>
</template>

<script setup>
defineProps({
  error: { type: Object, default: null },
  showRestart: { type: Boolean, default: false },
})

defineEmits(['dismiss', 'restart'])
</script>

<style scoped>
.error-banner {
  margin: var(--pt-space-2) var(--pt-space-3) 0;
  flex: none;
}

.error-banner__hint {
  font-size: 12px;
}

.error-banner__detail pre {
  max-height: 160px;
  overflow: auto;
  margin: 0;
  font-size: 11px;
  white-space: pre-wrap;
  word-break: break-all;
}
</style>
```

- [ ] **Step 6: 在 App.vue 中接入健康检查**

在 `app/renderer/src/App.vue` 的 `<script setup>` 中替换后端状态管理部分为：

```js
import ErrorBanner from './components/ErrorBanner.vue'
import { describeError, useBackendHealth } from './composables/useBackendHealth.js'

const health = useBackendHealth()
const pageError = ref(null)
const showRestart = ref(false)

let unsubscribeStatus = null
let unsubscribeExit = null

async function onRestartBackend() {
  const ok = await health.restart()
  if (ok) {
    ElMessage.success('后端已重启')
    showRestart.value = false
    pageError.value = null
  } else {
    ElMessage.error('重启失败')
  }
}

function reportError(error) {
  pageError.value = describeError(error)
  showRestart.value = error && (error.code === 'E_BACKEND_DOWN' || error.code === 'E_TIMEOUT')
}

onMounted(async () => {
  unsubscribeStatus = api.onBackendStatus((status) => {
    health.status.value = status
  })
  unsubscribeExit = api.onBackendExit(() => {
    health.status.value = 'down'
    showRestart.value = true
    pageError.value = describeError({
      code: 'E_BACKEND_DOWN',
      message: '后端进程已退出',
      hint: ERROR_HINTS.E_BACKEND_DOWN,
    })
    ElMessage.error('后端进程已退出')
  })

  // 每 15 秒做一次轻量健康检查
  const ok = await health.check()
  if (!ok && health.lastError.value) {
    pageError.value = health.lastError.value
    showRestart.value = true
  }
  healthTimer = setInterval(() => {
    if (health.status.value !== 'running') health.check()
  }, 15000)
})

let healthTimer = null

onBeforeUnmount(() => {
  if (unsubscribeStatus) unsubscribeStatus()
  if (unsubscribeExit) unsubscribeExit()
  if (healthTimer) clearInterval(healthTimer)
})
```

模板中在 `<Toolbar>` 之后加入：

```html
    <ErrorBanner
      :error="pageError"
      :show-restart="showRestart"
      @dismiss="pageError = null"
      @restart="onRestartBackend"
    />
```

并把 `StatusBar` 的 `backend-status` 绑定改为 `health.status`。

同时把所有已有的 `catch (error) { ElMessage.error(error.message) }` 改为 `catch (error) { reportError(error); ElMessage.error(error.message) }`，使错误同时进入横幅与瞬时提示。

把 `ERROR_HINTS` 加入从 `useBackendHealth.js` 的导入列表。

- [ ] **Step 7: 手动验证后端崩溃恢复**

Run: `npm run dev`

操作：在任务管理器中结束 `python.exe` 进程。

Expected:
- 界面顶部出现红色横幅「后端进程已退出」并带「重启后端」按钮
- 点击「重启后端」后横幅消失，状态栏显示「后端已连接」
- 再次打开文件夹可正常工作

- [ ] **Step 8: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现后端健康检查、重启与结构化错误横幅"
```

---

### Task 23: 实现键盘导航与快捷键体系

**Files:**
- Create: `app/renderer/src/composables/useShortcuts.js`
- Modify: `app/renderer/src/App.vue`
- Test: `app/renderer/src/composables/__tests__/shortcuts.spec.js`

**Interfaces:**
- Consumes: `library`、`editor` store
- Produces: `matchShortcut(event) -> string | null`（返回动作名）；`nextIndex(current, total, delta) -> number`；`useShortcuts(handlers)`

- [ ] **Step 1: 写快捷键匹配的失败测试**

创建 `app/renderer/src/composables/__tests__/shortcuts.spec.js`：

```js
import { describe, expect, it } from 'vitest'

import { isEditableTarget, matchShortcut, nextIndex } from '../useShortcuts.js'

function key(overrides) {
  return { key: '', ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, ...overrides }
}

describe('matchShortcut', () => {
  it('matches arrow navigation', () => {
    expect(matchShortcut(key({ key: 'ArrowDown' }))).toBe('next-image')
    expect(matchShortcut(key({ key: 'ArrowUp' }))).toBe('prev-image')
    expect(matchShortcut(key({ key: 'ArrowRight' }))).toBe('next-image')
    expect(matchShortcut(key({ key: 'ArrowLeft' }))).toBe('prev-image')
  })

  it('matches save shortcuts', () => {
    expect(matchShortcut(key({ key: 's', ctrlKey: true }))).toBe('save-current')
    expect(matchShortcut(key({ key: 'S', ctrlKey: true, shiftKey: true }))).toBe('save-all')
  })

  it('matches undo and redo', () => {
    expect(matchShortcut(key({ key: 'z', ctrlKey: true }))).toBe('undo')
    expect(matchShortcut(key({ key: 'z', ctrlKey: true, shiftKey: true }))).toBe('redo')
    expect(matchShortcut(key({ key: 'y', ctrlKey: true }))).toBe('redo')
  })

  it('matches select all, rename focus and delete', () => {
    expect(matchShortcut(key({ key: 'a', ctrlKey: true }))).toBe('select-all')
    expect(matchShortcut(key({ key: 'F2' }))).toBe('focus-rename')
    expect(matchShortcut(key({ key: 'Delete' }))).toBe('delete')
  })

  it('returns null for unbound keys', () => {
    expect(matchShortcut(key({ key: 'q' }))).toBeNull()
    expect(matchShortcut(key({ key: 'Escape' }))).toBeNull()
  })

  it('does not treat a bare s as save', () => {
    expect(matchShortcut(key({ key: 's' }))).toBeNull()
  })
})

describe('isEditableTarget', () => {
  it('detects inputs and contenteditable', () => {
    expect(isEditableTarget({ tagName: 'INPUT' })).toBe(true)
    expect(isEditableTarget({ tagName: 'TEXTAREA' })).toBe(true)
    expect(isEditableTarget({ tagName: 'DIV', isContentEditable: true })).toBe(true)
    expect(isEditableTarget({ tagName: 'DIV', isContentEditable: false })).toBe(false)
    expect(isEditableTarget(null)).toBe(false)
  })
})

describe('nextIndex', () => {
  it('advances and retreats within bounds', () => {
    expect(nextIndex(0, 5, 1)).toBe(1)
    expect(nextIndex(4, 5, -1)).toBe(3)
  })

  it('clamps at both ends', () => {
    expect(nextIndex(4, 5, 1)).toBe(4)
    expect(nextIndex(0, 5, -1)).toBe(0)
  })

  it('returns -1 for an empty list', () => {
    expect(nextIndex(0, 0, 1)).toBe(-1)
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

Run: `npx vitest run app/renderer/src/composables/__tests__/shortcuts.spec.js`
Expected: FAIL — 无法解析 `../useShortcuts.js`

- [ ] **Step 3: 实现快捷键匹配**

创建 `app/renderer/src/composables/useShortcuts.js`：

```js
import { onBeforeUnmount, onMounted } from 'vue'

import { useEditorStore } from '../stores/editor.js'
import { useLibraryStore } from '../stores/library.js'

export function isEditableTarget(target) {
  if (!target) return false
  if (target.isContentEditable) return true
  const tag = target.tagName
  return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT'
}

export function matchShortcut(event) {
  const mod = event.ctrlKey || event.metaKey
  const key = event.key

  if (mod && key.toLowerCase() === 's') return event.shiftKey ? 'save-all' : 'save-current'
  if (mod && key.toLowerCase() === 'z') return event.shiftKey ? 'redo' : 'undo'
  if (mod && key.toLowerCase() === 'y') return 'redo'
  if (mod && key.toLowerCase() === 'a') return 'select-all'

  if (key === 'Delete') return 'delete'
  if (key === 'F2') return 'focus-rename'
  if (key === 'ArrowDown' || key === 'ArrowRight') return 'next-image'
  if (key === 'ArrowUp' || key === 'ArrowLeft') return 'prev-image'

  return null
}

export function nextIndex(current, total, delta) {
  if (!total) return -1
  const candidate = (current < 0 ? 0 : current) + delta
  return Math.min(total - 1, Math.max(0, candidate))
}

/**
 * 绑定全局快捷键。输入框聚焦时只放行保存与撤销类操作，
 * 避免在输入文件名时被方向键改图。
 */
export function useShortcuts(handlers) {
  const library = useLibraryStore()
  const editor = useEditorStore()

  function onKeyDown(event) {
    const action = matchShortcut(event)
    if (!action) return

    const editable = isEditableTarget(event.target)
    const allowedWhileEditing = ['save-current', 'save-all', 'undo', 'redo']

    if (editable && !allowedWhileEditing.includes(action)) return

    event.preventDefault()

    if (action === 'next-image' || action === 'prev-image') {
      const visible = library.visibleImages.map((item) => item.name)
      const current = visible.indexOf(editor.currentName)
      const target = nextIndex(current, visible.length, action === 'next-image' ? 1 : -1)
      if (target >= 0) {
        const name = visible[target]
        library.selectOnly(name)
        editor.setCurrent(name, library.folder)
        if (handlers.onImageChanged) handlers.onImageChanged(name)
      }
      return
    }

    if (action === 'select-all') {
      library.selectedNames = library.visibleImages.map((item) => item.name)
      return
    }

    const handler = handlers[action]
    if (handler) handler()
  }

  onMounted(() => window.addEventListener('keydown', onKeyDown))
  onBeforeUnmount(() => window.removeEventListener('keydown', onKeyDown))
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `npx vitest run app/renderer/src/composables/__tests__/shortcuts.spec.js`
Expected: 10 passed

- [ ] **Step 5: 在 App.vue 中启用快捷键**

在 `app/renderer/src/App.vue` 的 `<script setup>` 中加入：

```js
import { useShortcuts } from './composables/useShortcuts.js'

useShortcuts({
  'save-current': onSaveCurrent,
  'save-all': onSaveAll,
  undo: () => editor.undo(),
  redo: () => editor.redo(),
  'focus-rename': () => {
    if (workspaceRef.value) workspaceRef.value.focusRename()
  },
  delete: () => ElMessage.info('删除功能在本版本中未启用，请手动在文件夹中操作'),
  onImageChanged: () => {
    if (workspaceRef.value) workspaceRef.value.resetPreview()
  },
})
```

在 `Workspace.vue` 的 `defineExpose` 中补：

```js
defineExpose({
  scrollToTop: () => listPaneRef.value?.scrollToTop(),
  focusRename: () => inspectorRef.value?.focusRename(),
  resetPreview: () => previewRef.value?.reset(),
})
```

同时把 `Workspace.vue` 中此前自行绑定的 `onKeyDown`（Task 17 添加的 Ctrl+Z / Ctrl+Y / Escape / F2 处理）**删除**，避免与 `useShortcuts` 重复触发。保留 Escape 结束裁剪的分支：

```js
function onKeyDown(event) {
  if (event.key === 'Escape' && cropping.value) {
    cropping.value = false
  }
}
```

- [ ] **Step 6: 手动验证快捷键**

Run: `npm run dev`

操作：打开文件夹 → 连续按 `↓` 与 `↑`。

Expected: 当前图片在列表中上下移动，中栏预览同步更新，不会越界。

操作：选图 → 按 `Ctrl+Z` 与 `Ctrl+Y`。
Expected: 编辑被撤销/重做。

操作：焦点在搜索框内 → 按 `↓`。
Expected: 光标在输入框内移动，**不**切换图片。

操作：按 `F2`。
Expected: 焦点跳到重命名输入框。

- [ ] **Step 7: 运行全部测试**

Run: `npx vitest run`
Expected: 全部通过

Run: `venv\Scripts\python.exe -m pytest backend/tests -q`
Expected: 全部通过

- [ ] **Step 8: 提交**

```bash
git add app/renderer/src/
git commit -m "feat: 实现键盘导航与快捷键体系"
```

---

## 里程碑 M6：打包与发布

### Task 24: 配置 PyInstaller 打包后端

**Files:**
- Create: `backend.spec`
- Modify: `package.json`（`backend` 脚本）
- Modify: `.gitignore`

**Interfaces:**
- Consumes: `backend/__main__.py`（Task 1–10）
- Produces: `build/backend/backend.exe` 与同目录的依赖文件

- [ ] **Step 1: 安装打包依赖**

Run: `venv\Scripts\python.exe -m pip install pyinstaller==6.11.1`
Expected: 安装成功

- [ ] **Step 2: 编写打包配置**

创建 `backend.spec`：

```python
# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 配置：以 --onedir 方式打包后端。

选择 onedir 而非 onefile：启动更快，且显著降低杀毒软件误报。
console=False 使打包后不弹出黑色控制台窗口，靠 stdio 管道通信。
"""

block_cipher = None

a = Analysis(
    ['backend/__main__.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'backend.transport_stdio',
        'backend.transport_http',
        'backend.selftest',
        'backend.rpc',
        'backend.cache',
        'backend.security',
        'backend.errors',
        'backend.services.docx_extract',
        'backend.services.image_ops',
        'backend.services.library',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'test', 'unittest', 'pydoc_data'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='backend',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='backend',
)
```

- [ ] **Step 3: 更新 npm 脚本**

把 `package.json` 的 `backend` 脚本替换为：

```json
    "backend": "pyinstaller --noconfirm --clean backend.spec --distpath build",
```

- [ ] **Step 4: 更新 .gitignore**

在 `.gitignore` 中确认存在以下条目，不存在则追加：

```
build/
dist/
release/
backend.spec.bak
```

注意：`build/` 已在 Task 1 追加。若 `build/` 已被忽略则无需重复。

- [ ] **Step 5: 执行打包**

Run: `npm run backend`
Expected: 打包完成，生成 `build/backend/backend.exe`

- [ ] **Step 6: 验证打包产物可独立运行**

Run:
```powershell
'{"id":1,"method":"app.ping","params":{}}' | .\build\backend\backend.exe --stdio
```
Expected: 输出单行 JSON 含 `python` 与 `pillow` 字段，且**没有**控制台窗口弹出（因为 `console=False`，回显可能为空——若为空，改用下述脚本验证）

创建临时验证脚本 `scripts/verify-backend-exe.mjs`：

```js
import { spawn } from 'node:child_process'

const child = spawn('build/backend/backend.exe', ['--stdio'], {
  stdio: ['pipe', 'pipe', 'inherit'],
  windowsHide: true,
})

let buffer = ''
child.stdout.setEncoding('utf-8')
child.stdout.on('data', (chunk) => {
  buffer += chunk
  const index = buffer.indexOf('\n')
  if (index >= 0) {
    console.log('response:', buffer.slice(0, index))
    child.kill()
    process.exit(0)
  }
})

child.stdin.write(JSON.stringify({ id: 1, method: 'app.ping', params: {} }) + '\n')

setTimeout(() => {
  console.error('no response within 15s')
  child.kill()
  process.exit(1)
}, 15000)
```

Run: `node scripts/verify-backend-exe.mjs`
Expected: 打印 `response: {"id":1,"result":{...}}`，退出码 0

验证后**保留**该脚本（它是有用的打包冒烟测试），加入版本管理。

- [ ] **Step 7: 确认打包体积与依赖无 tkinter**

Run: `Get-ChildItem build/backend -Recurse -Filter "*tk*" | Select-Object -First 5 FullName`
Expected: 无输出（tkinter 已被排除）

Run: `(Get-ChildItem build/backend -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB`
Expected: 打印体积，量级为几十 MB

- [ ] **Step 8: 提交**

```bash
git add backend.spec package.json .gitignore scripts/verify-backend-exe.mjs
git commit -m "chore: 配置 PyInstaller onedir 打包后端并加冒烟验证脚本"
```

---

### Task 25: 配置 electron-builder 产出安装包与便携版

**Files:**
- Create: `electron-builder.yml`
- Modify: `package.json`
- Modify: `app/main/paths.js`（确认打包路径解析）
- Create: `build/installer.nsh`（可选，若需自定义安装行为）

**Interfaces:**
- Consumes: `build/backend/`（Task 24）、`app/renderer/dist/`（Task 13 起可构建）
- Produces: `release/` 下的 NSIS 安装包与 portable 可执行文件

- [ ] **Step 1: 编写 electron-builder 配置**

创建 `electron-builder.yml`：

```yaml
appId: com.mandyluer.picture-tools
productName: Word图片提取与重命名工具
copyright: Copyright © 2026

directories:
  output: release
  buildResources: build-resources

files:
  - app/main/**/*
  - app/preload/**/*
  - app/renderer/dist/**/*
  - package.json

extraResources:
  - from: build/backend
    to: backend
    filter:
      - '**/*'

win:
  target:
    - target: nsis
      arch: [x64]
    - target: portable
      arch: [x64]
  icon: 奶牛猫.ico
  artifactName: ${productName}-${version}-${arch}.${ext}

nsis:
  oneClick: false
  perMachine: false
  allowToChangeInstallationDirectory: true
  createDesktopShortcut: true
  createStartMenuShortcut: true
  shortcutName: Word图片提取与重命名工具
  installerIcon: 奶牛猫.ico
  uninstallerIcon: 奶牛猫.ico
  deleteAppDataOnUninstall: false

portable:
  artifactName: ${productName}-${version}-portable.${ext}
```

- [ ] **Step 2: 在 package.json 中补充构建脚本与元信息**

把 `package.json` 的 `scripts` 补齐为：

```json
  "scripts": {
    "dev": "concurrently -k \"npm:dev:renderer\" \"npm:dev:electron\"",
    "dev:renderer": "vite",
    "dev:electron": "wait-on tcp:5173 && cross-env ELECTRON_DEV=1 electron .",
    "test": "npm run test:backend && npm run test:unit",
    "test:backend": "venv\\Scripts\\python.exe -m pytest backend/tests -q",
    "test:unit": "vitest run",
    "test:e2e": "playwright test",
    "backend": "pyinstaller --noconfirm --clean backend.spec --distpath build",
    "build:renderer": "vite build",
    "build": "npm run build:renderer && npm run backend",
    "dist": "npm run build && electron-builder --win --config electron-builder.yml"
  },
```

并在 `package.json` 顶层补充：

```json
  "author": "MandyLuer",
  "license": "MIT",
```

- [ ] **Step 3: 创建 buildResources 目录占位**

Run: `New-Item -ItemType Directory -Force -Path build-resources | Out-Null; New-Item -ItemType File -Force -Path build-resources/.gitkeep | Out-Null`

Expected: 目录创建成功

在 `.gitignore` 中确认 `build-resources/.gitkeep` 不会被忽略（若有 `build-resources/` 规则则移除或加例外）。

- [ ] **Step 4: 执行完整构建**

Run: `npm run build`
Expected: 前端构建成功，后端打包成功，`app/renderer/dist/index.html` 与 `build/backend/backend.exe` 均存在

- [ ] **Step 5: 执行打包**

Run: `npx electron-builder --win --config electron-builder.yml`
Expected: `release/` 下生成 NSIS 安装包（`.exe`）与 portable 可执行文件

- [ ] **Step 6: 验证打包后的应用可运行**

操作：运行 `release/` 下名字含 `portable` 的可执行文件。

Expected:
- 应用窗口打开，界面正常显示三栏布局
- 状态栏显示「后端已连接」
- 能打开文件夹、看到缩略图列表
- 能旋转一张图并保存成功

若状态栏显示「后端未连接」，检查 `app/main/paths.js` 的 `resolveBackendLaunch` 是否指向 `path.join(resourcesPath, 'backend', 'backend.exe')`，并确认 `extraResources` 已把 `build/backend` 复制为 `backend`。

- [ ] **Step 7: 提交**

```bash
git add electron-builder.yml package.json build-resources/
git commit -m "chore: 配置 electron-builder 产出 NSIS 安装包与便携版"
```

---

### Task 26: 编写端到端测试与更新文档

**Files:**
- Create: `playwright.config.js`
- Create: `tests/e2e/helpers.js`
- Create: `tests/e2e/extract-and-rename.spec.js`
- Create: `tests/e2e/edit-and-save.spec.js`
- Modify: `README.md`

**Interfaces:**
- Consumes: 打包后的应用（Task 25）或开发模式
- Produces: `npm run test:e2e` 可运行的端到端测试

- [ ] **Step 1: 配置 Playwright**

创建 `playwright.config.js`：

```js
import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests/e2e',
  timeout: 90000,
  expect: { timeout: 15000 },
  fullyParallel: false,
  workers: 1,
  reporter: [['list']],
  use: {
    trace: 'retain-on-failure',
  },
})
```

安装浏览器依赖不需要（Electron 自带 Chromium），但需要 Playwright 的 Electron 支持：

Run: `npm install --save-dev playwright`
Expected: 安装成功

- [ ] **Step 2: 编写测试辅助工具**

创建 `tests/e2e/helpers.js`：

```js
import { _electron as electron } from 'playwright'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'

const APP_ROOT = path.resolve(import.meta.dirname, '..', '..')

/**
 * 启动打包前的主进程（开发模式），返回 { app, page, close }。
 * 后端用仓库 venv 中的 Python，避免依赖打包产物。
 */
export async function launchApp() {
  const app = await electron.launch({
    args: [APP_ROOT],
    cwd: APP_ROOT,
    env: {
      ...process.env,
      ELECTRON_DEV: '',
      PICTURE_TOOLS_PYTHON: path.join(APP_ROOT, 'venv', 'Scripts', 'python.exe'),
    },
  })

  const page = await app.firstWindow()
  await page.waitForLoadState('domcontentloaded')

  return {
    app,
    page,
    close: async () => {
      await app.close()
    },
  }
}

export async function makeTempDir(prefix) {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix))
}

export async function copyFixture(name, targetDir) {
  const source = path.join(APP_ROOT, 'backend', 'tests', 'fixtures', name)
  const target = path.join(targetDir, name)
  await fs.copyFile(source, target)
  return target
}

export async function makeImage(targetDir, name, size = [40, 20], color = [200, 100, 50]) {
  const { PNG } = await import('pngjs').catch(() => ({ PNG: null }))
  if (!PNG) throw new Error('需要 pngjs 才能生成测试图片，请运行 npm i -D pngjs')
  const png = new PNG({ width: size[0], height: size[1] })
  for (let i = 0; i < png.data.length; i += 4) {
    png.data[i] = color[0]
    png.data[i + 1] = color[1]
    png.data[i + 2] = color[2]
    png.data[i + 3] = 255
  }
  const target = path.join(targetDir, name)
  await fs.writeFile(target, PNG.sync.write(png))
  return target
}

export function readPngSize(buffer) {
  // PNG 尺寸位于 IHDR：偏移 16 起的 8 字节为大端宽高
  return [buffer.readUInt32BE(16), buffer.readUInt32BE(20)]
}
```

Run: `npm install --save-dev pngjs`
Expected: 安装成功（`helpers.js` 需要它生成测试图片）

- [ ] **Step 3: 编写提取与重命名端到端测试**

创建 `tests/e2e/extract-and-rename.spec.js`：

```js
import fs from 'node:fs/promises'
import path from 'node:path'
import { expect, test } from '@playwright/test'

import { copyFixture, launchApp, makeTempDir } from './helpers.js'

test('提取 docx 图片并按文档顺序命名', async () => {
  const workDir = await makeTempDir('pt-extract-')
  const docxPath = await copyFixture('reordered.docx', workDir)
  const outDir = path.join(workDir, 'extracted')

  const { page, close } = await launchApp()
  try {
    // 直接用后端接口驱动提取，避免依赖原生文件对话框
    const result = await page.evaluate(
      async ([docx, out]) => {
        const response = await window.api.request('docx.extract', {
          path: docx,
          outDir: out,
          pattern: '{index:03d}',
          startIndex: 1,
          onConflict: 'autoSuffix',
        })
        return response
      },
      [docxPath, outDir],
    )

    expect(result.ok).toBe(true)
    expect(result.result.written).toEqual(['001.png', '002.png', '003.png'])
    expect(result.result.orderSource).toBe('xml')

    const files = await fs.readdir(outDir)
    expect(files.sort()).toEqual(['001.png', '002.png', '003.png'])
  } finally {
    await close()
  }
})

test('提取不删除输出目录中已有的文件', async () => {
  const workDir = await makeTempDir('pt-preserve-')
  const docxPath = await copyFixture('plain.docx', workDir)
  const outDir = path.join(workDir, 'extracted')
  await fs.mkdir(outDir, { recursive: true })
  const keepPath = path.join(outDir, 'my_notes.txt')
  await fs.writeFile(keepPath, 'keep me', 'utf-8')

  const { page, close } = await launchApp()
  try {
    await page.evaluate(
      async ([docx, out]) => {
        await window.api.request('docx.extract', { path: docx, outDir: out })
      },
      [docxPath, outDir],
    )

    expect(await fs.readFile(keepPath, 'utf-8')).toBe('keep me')
  } finally {
    await close()
  }
})

test('重命名后磁盘文件名确实改变', async () => {
  const workDir = await makeTempDir('pt-rename-')
  const docxPath = await copyFixture('plain.docx', workDir)
  const outDir = path.join(workDir, 'extracted')

  const { page, close } = await launchApp()
  try {
    await page.evaluate(
      async ([docx, out]) => {
        await window.api.request('docx.extract', { path: docx, outDir: out })
      },
      [docxPath, outDir],
    )

    const result = await page.evaluate(
      async ([dir]) => {
        return window.api.request('library.rename', {
          dir,
          from: '001.png',
          to: '封面',
          onConflict: 'reject',
        })
      },
      [outDir],
    )

    expect(result.ok).toBe(true)
    const files = await fs.readdir(outDir)
    expect(files).toContain('封面.png')
    expect(files).not.toContain('001.png')
  } finally {
    await close()
  }
})

test('重名时默认拒绝且不覆盖已有文件', async () => {
  const workDir = await makeTempDir('pt-conflict-')
  const docxPath = await copyFixture('plain.docx', workDir)
  const outDir = path.join(workDir, 'extracted')

  const { page, close } = await launchApp()
  try {
    await page.evaluate(
      async ([docx, out]) => {
        await window.api.request('docx.extract', { path: docx, outDir: out })
      },
      [docxPath, outDir],
    )

    const before = await fs.readFile(path.join(outDir, '002.png'))

    const result = await page.evaluate(
      async ([dir]) => {
        return window.api.request('library.rename', {
          dir,
          from: '001.png',
          to: '002',
          onConflict: 'reject',
        })
      },
      [outDir],
    )

    expect(result.ok).toBe(false)
    expect(result.error.code).toBe('E_NAME_EXISTS')
    expect(Buffer.compare(await fs.readFile(path.join(outDir, '002.png')), before)).toBe(0)
  } finally {
    await close()
  }
})
```

- [ ] **Step 4: 编写编辑与保存端到端测试**

创建 `tests/e2e/edit-and-save.spec.js`：

```js
import fs from 'node:fs/promises'
import path from 'node:path'
import { expect, test } from '@playwright/test'

import { launchApp, makeImage, makeTempDir, readPngSize } from './helpers.js'

test('旋转后保存会改变磁盘上的像素尺寸', async () => {
  const workDir = await makeTempDir('pt-edit-')
  const imagePath = await makeImage(workDir, '001.png', [40, 20])
  const before = readPngSize(await fs.readFile(imagePath))
  expect(before).toEqual([40, 20])

  const { page, close } = await launchApp()
  try {
    const result = await page.evaluate(
      async ([target]) => {
        return window.api.request('image.save', {
          path: target,
          targetPath: target,
          ops: [{ type: 'rotate', angle: 90 }],
          format: 'png',
        })
      },
      [imagePath],
    )

    expect(result.ok).toBe(true)
    expect(readPngSize(await fs.readFile(imagePath))).toEqual([20, 40])
  } finally {
    await close()
  }
})

test('裁剪使用源图像坐标并产生预期尺寸', async () => {
  const workDir = await makeTempDir('pt-crop-')
  const imagePath = await makeImage(workDir, '001.png', [40, 20])

  const { page, close } = await launchApp()
  try {
    const result = await page.evaluate(
      async ([target]) => {
        return window.api.request('image.save', {
          path: target,
          targetPath: target,
          ops: [{ type: 'crop', rect: { left: 10, top: 5, right: 30, bottom: 15 } }],
          format: 'png',
        })
      },
      [imagePath],
    )

    expect(result.ok).toBe(true)
    expect(readPngSize(await fs.readFile(imagePath))).toEqual([20, 10])
  } finally {
    await close()
  }
})

test('裁剪后再保存，原文件名保持不变', async () => {
  const workDir = await makeTempDir('pt-cropname-')
  const imagePath = await makeImage(workDir, '我的封面.png', [40, 20])

  const { page, close } = await launchApp()
  try {
    await page.evaluate(
      async ([target]) => {
        return window.api.request('image.save', {
          path: target,
          targetPath: target,
          ops: [{ type: 'crop', rect: { left: 0, top: 0, right: 20, bottom: 10 } }],
          format: 'png',
        })
      },
      [imagePath],
    )

    const files = await fs.readdir(workDir)
    // 关键回归断言：不产生 _cropped 文件，且原名保留
    expect(files).toEqual(['我的封面.png'])
    expect(files.some((name) => name.includes('_cropped'))).toBe(false)
  } finally {
    await close()
  }
})

test('签名在旋转四次后回到基线', async () => {
  const workDir = await makeTempDir('pt-signature-')
  const imagePath = await makeImage(workDir, '001.png', [40, 20])

  const { page, close } = await launchApp()
  try {
    const result = await page.evaluate(
      async ([target]) => {
        const baseline = await window.api.request('image.signature', { path: target, ops: [] })
        const spun = await window.api.request('image.signature', {
          path: target,
          ops: [
            { type: 'rotate', angle: 90 },
            { type: 'rotate', angle: 90 },
            { type: 'rotate', angle: 90 },
            { type: 'rotate', angle: 90 },
          ],
        })
        return { baseline: baseline.result.signature, spun: spun.result.signature }
      },
      [imagePath],
    )

    expect(result.spun).toBe(result.baseline)
  } finally {
    await close()
  }
})
```

- [ ] **Step 5: 运行端到端测试**

Run: `npm run test:e2e`
Expected: 全部通过

若因 Electron 沙箱或环境限制导致启动失败，记录具体错误并在 `playwright.config.js` 中改用 `args: [APP_ROOT, '--no-sandbox']` 后重试。

- [ ] **Step 6: 更新 README**

覆盖写入 `README.md`：

```markdown
# Word 图片提取与重命名工具

## 介绍

从 Word（.docx）文档中按插入顺序提取内嵌图片，并对图片进行重命名与编辑。
支持旋转、水平/垂直翻转、画布内裁剪与压缩，编辑采用非破坏性模型：改动先在预览中生效，
点击保存时才写回磁盘，且支持撤销/重做。

应用形态为 Windows 桌面客户端：Vue 前端 + Electron 外壳 + Python 后端。

## 开发环境

| 依赖 | 版本 |
|---|---|
| Python | 3.10（仓库自带 `venv`） |
| Node.js | 18 以上（实测 24.9.0） |

安装依赖：

```bash
venv\Scripts\python.exe -m pip install -r requirements-dev.txt
npm install
```

## 运行

```bash
npm run dev
```

## 测试

```bash
npm test              # 后端 pytest + 前端 vitest
npm run test:e2e      # Playwright 驱动 Electron 的端到端测试
venv\Scripts\python.exe -m backend --selftest   # 后端自检，退出码 0 表示通过
```

## 打包

```bash
npm run build         # 构建前端 + 打包后端
npm run dist          # 产出 NSIS 安装包与便携版，输出到 release/
```

## 目录结构

| 目录 | 说明 |
|---|---|
| `backend/` | Python 后端：docx 解析、图片操作、目录与重命名服务 |
| `app/main/` | Electron 主进程：窗口、子进程监管、appimg 协议、IPC |
| `app/preload/` | contextBridge 白名单 API |
| `app/renderer/` | Vue 3 + Element Plus 三栏工作台 |
| `backend/tests/fixtures/` | 测试用样例 docx |
| `docs/superpowers/` | 设计文档与实施计划 |

## 与旧版本的关系

仓库根目录的 `image-editer.py` 是历史 Tkinter 实现，功能已被本客户端完全覆盖，
保留仅为回退参考，不再维护。

## 主要修复

1. **裁剪后文件名丢失**——编辑改为操作栈模型，保存时沿用当前文件名，不再派生 `_cropped` 名称。
2. **提取时清空输出目录**——提取改为冲突安全命名，绝不删除目录中已有文件。
3. **`imghdr` 在 Python 3.13 被移除**——改为读取文件头魔数自主识别格式。
4. **排序对纯文字文件名异常**——改为自然排序，数字名按数值排、文字名按字典序排。
5. **编辑无撤销且即写即弃**——新增操作栈、撤销/重做与未保存退出提醒。
```

- [ ] **Step 7: 更新 requirements.txt 说明**

在 `requirements.txt` 末尾追加注释说明：

```
# 说明：lxml 与 python-docx 已移除。
# 现有代码从未使用它们（lxml 未被导入，python-docx 的 Document 导入后未被调用），
# docx 解析只依赖标准库 zipfile 与 xml.etree.ElementTree。
```

- [ ] **Step 8: 运行完整测试套件**

Run: `npm test`
Expected: 后端 pytest 与前端 vitest 全部通过

Run: `npm run test:e2e`
Expected: 全部通过

- [ ] **Step 9: 提交**

```bash
git add playwright.config.js tests/ README.md requirements.txt
git commit -m "test: 新增端到端测试并更新文档至 Electron 客户端版本"
```

---

## 完成标准

全部 26 个任务完成后，以下条件必须同时成立：

1. `venv\Scripts\python.exe -m pytest backend/tests -q` 零失败。
2. `npx vitest run` 零失败。
3. `npm run test:e2e` 零失败。
4. `venv\Scripts\python.exe -m backend --selftest` 退出码为 0。
5. `npm run dist` 在 `release/` 产出可安装的 NSIS 包与可运行的便携版。
6. 在便携版中完成一次完整流程：导入 `reordered.docx` → 提取出顺序正确的三张图 → 旋转其中一张 → 撤销 → 重做 → 重命名为中文名 → 保存 → 磁盘上文件名与像素均已更新。
7. 输出目录中预先放入的其他文件在提取后仍然存在。
8. 编辑后保存不产生任何 `_cropped` 或 `_compressed` 文件。
9. 关闭窗口时有未保存改动则弹出确认。
