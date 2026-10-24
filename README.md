# 拾图 Shitu · Word 图片工作台

Vue 3 + FastAPI 前后端分离重构，使用 pywebview 提供桌面窗口，使用 PyInstaller 分发客户端。旧版 Tkinter 实现可通过 Git 历史查看。

## 功能

- 按 DOCX 正文中的图片引用顺序提取图片，兼容 DrawingML 和旧式 VML，重复引用只提取一次。
- 图片预览旁显示上下文和命名参考：同段文字、上下各两段及替代文字；疑似题注优先显示，点击填入文件名后确认保存。同一图片的多处引用分别保留描述，编辑后继续保留描述。
- 导入已有图片或图片文件夹，逐张预览、方向键切换、滚轮缩放、拖动平移。
- 重命名（保留格式扩展名）、左右旋转、水平/垂直翻转、拖动选区裁剪、JPEG 质量压缩和 PNG 无损优化。
- 旋转、翻转、裁剪、压缩直接更新当前工作区图片，保留名称和列表位置，不生成副本；预览与缩略图自动刷新。导出全部图片 ZIP，桌面版支持直接导出到目录。
- 图片列表在固定窗口内独立滚动，方向键切图时保持当前项可见；列表每张图片支持删除。清空和删除使用界面内确认框，取消时保留内容。
- 加宽图片预览区，支持收起参考描述和「大图模式」（Esc 退出），桌面窗口默认 1440 × 960。
- 加载、导出、清空、删除和编辑后会在页面顶部弹出结果提示，成功提示自动消失，失败提示显示原因；取消操作不提示成功。桌面 ZIP 导出在保存完成后提示，浏览器 ZIP 下载提示已开始下载。
- 文件名校验、重名保护、导入失败回滚。导出到目录遇到同名文件会拒绝操作，不清空或覆盖原有文件。

## 开发运行

需要 Python 3.10+、Node.js 22.12+。建议创建独立虚拟环境。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m backend
```

另开终端启动前端：

```powershell
cd frontend
npm ci
npm run dev
```

访问 http://127.0.0.1:5173 。Vite 将 `/api` 代理到本机 8000 端口，前后端独立开发。首次检出没有 lock 文件时先执行 `npm install`。

## 浏览器运行构建版

```powershell
cd frontend
npm ci
npm run build
cd ..
python -m backend
```

访问 http://127.0.0.1:8000 。后端同时提供构建后的静态文件。

## 桌面运行和打包

```powershell
python -m pip install -r requirements/desktop.txt
cd frontend
npm ci
npm run build
cd ..
python desktop.py
```

Windows 一键打包：

```powershell
.\scripts\build-desktop.ps1
```

软件名为「拾图 Shitu」，图标原图和 ICO 位于 `assets/branding/`。打包脚本会同步生成浏览器及桌面图标，并输出 `dist/Shitu-2.3.0-Windows-x64.zip`；Windows 文件属性包含产品名称和版本。

或在目标平台执行：

```powershell
python -m PyInstaller --noconfirm --distpath dist/2.3.0 packaging/desktop.spec
```

分发整个 `dist/2.3.0/Shitu/` 文件夹，Windows 入口为 `拾图 Shitu.exe`。按版本输出目录，避免覆盖正在运行的旧版。用户不需要安装 Python 或 Node.js。Windows 需要 WebView2 Runtime；macOS/Linux 需要对应的 pywebview 原生 WebView 依赖。必须在各目标操作系统分别构建，不能跨平台生成二进制；签名、公证和安装包封装按发布平台另行配置。新打包使用 `packaging/desktop.spec`。

## 目录与接口

```text
picture-tools/
├── assets/           品牌图标原图及 Windows ICO
├── backend/          Python API 和图片处理
├── frontend/         Vue 前端（dist 为前端构建输出）
├── packaging/        PyInstaller 打包配置
├── requirements/     Web、桌面依赖清单
├── scripts/          构建脚本
├── tests/            自动化测试
├── docs/             设计与历史方案
├── dist/             客户端及分发 ZIP（不纳入 Git）
├── desktop.py        桌面启动入口
├── requirements.txt  默认后端依赖入口
└── README.md
```

`.venv/` 和 `frontend/node_modules/` 是当前开发依赖；`build/`、`.pytest_cache/`、`__pycache__/` 是可再生成的临时目录。`.worktrees/` 中的 Electron 分支是独立 Git 工作区，保留供原有工作继续使用。

- `backend/service.py`：与界面无关的提取、重命名、编辑和导出逻辑。
- `backend/app.py`：FastAPI 应用工厂和 REST API。
- `frontend/`：Vue 单文件组件、样式和 Vite 配置。
- `desktop.py`：随机本机端口、每次启动访问令牌、原生目录对话框和退出清理。
- `tests/`：提取顺序、重命名、编辑、导出、权限和导入回滚测试。

接口：`POST /api/import`（multipart files）、`GET/DELETE /api/images`、`PATCH/DELETE /api/images/{id}`、`POST /api/images/{id}/edit`、`GET /api/images/{id}/file`、`GET /api/export`。运行后可访问 `/docs` 查看接口定义。

## 数据和使用边界

使用描述命名：重新导入 Word → 选中图片 → 在右侧「文档参考描述」查看上下文 → 点击「填入名称」→ 按需修改 → 点击「保存名称」。候选名称会清理文件名中的非法字符并截短；同名文件仍会提示冲突。

描述按 Word XML 的段落结构关联，搜索最近的上下各两段，遇到其他图片、表格单元格或结构边界时停止。上下文只是命名参考，无法保证是图片的正式题注；浮动布局、复杂表格需人工核对，不读取图片内部文字，也不包含页眉页脚描述。

文件位于进程临时工作区，退出后清理；退出前务必导出。桌面版关闭时会提示。浏览器刷新保留当前进程工作区，服务重启后不保留。每个服务进程对应一个共享工作区，当前版本面向单机单用户，不应直接作为公网多用户服务。

编辑、单张删除和清空作用于工作区，源 Word 文档和导入前的本地图片不受影响。编辑内容需导出后保存到磁盘。删除当前图片后自动选中下一张（末尾删除时选上一张），删除最后一张或清空后预览与名称一起复位。

单文件最多 50 MB，工作区及 DOCX 解压总量最多 250 MB。支持 DOCX，不支持旧式 `.doc`。提取顺序指正文 XML 顺序，浮动图片的视觉排列、页眉页脚及外链图片不在范围内。无法由 Pillow 解码的矢量媒体会跳过。GIF/TIFF 编辑处理当前首帧，导入和直接导出保留原始文件。压缩不保证每次生成更小的文件。

## 测试

```powershell
python -m pip install pytest httpx
python -m pytest tests -q
cd frontend
npm test
npm run build
```

安装桌面依赖后，可运行 `python desktop.py --gui-smoke-test`。此检查会启动真实 WebView，验证 Vue 界面、品牌图标和桌面桥接接口后自动关闭；`--smoke-test` 仅检查本地 HTTP 服务。打包后的 EXE 同样支持上述两个参数。

添加 `--test-image-list` 会使用 80 张合成测试图片验证独立滚动和页面高度；同时添加 `--compact-window` 可验证最小窗口。打包脚本会自动执行前端回归测试，以及普通窗口、最小窗口的图片列表检查。

客户端日志在 `%LOCALAPPDATA%/Shitu/logs/desktop.log`（源码运行时在项目 `logs/`）。若窗口初始化超过 30 秒，会生成 `startup-threads.log` 供排查。Windows 使用 WebView2 引擎；桥接仅开放目录导入、目录导出和 ZIP 导出三个方法。

当前 Web 和桌面依赖分别位于 `requirements/web.txt` 和 `requirements/desktop.txt`。
