# 拾图 Shitu

本地图片与证书截图工作台，当前版本 **2.6.0**。前端使用 Vue 3，后端使用 FastAPI，桌面版由 pywebview 和 PyInstaller 提供。支持网页预览和 Windows 客户端分发，旧版 Tkinter 实现可通过 Git 历史查看。

## 功能入口

| 菜单 | 功能 |
| --- | --- |
| Word 提图 | 按 DOCX 正文引用顺序提取图片，显示附近文字和命名参考 |
| 本地图片 | 导入图片或文件夹，重命名、旋转、翻转、裁剪和压缩 |
| 证书截图 | 批量添加证书，选择常用或自定义字段，生成真实 Windows 证书窗口截图 |

顶部菜单下方只显示对应操作；左侧是独立滚动的图片列表，右侧用于预览和编辑，右上方提供导出和清空。菜单共用图片工作区，切换时保留图片、名称草稿及编辑状态。

- 支持方向键切图、滚轮缩放、拖动平移、收起参考描述，以及大图模式（Esc 退出）。
- 旋转、翻转、裁剪和压缩直接更新当前工作区图片，保留名称、顺序和文档参考描述，不生成编辑副本。
- 单张删除和清空工作区均需要界面确认，不删除原始文档或导入前的本地图片。
- 加载、编辑、删除、清空和导出后显示操作提示；取消操作不提示成功。

## 使用客户端

Windows 分发包为 `dist/Shitu-2.6.0-Windows-x64.zip`。解压整个文件夹后运行 **拾图 Shitu.exe**，请保留同目录的依赖文件，不要只复制 EXE。

客户端不要求安装 Python 或 Node.js，需要 Windows WebView2 Runtime。程序默认窗口为 1440 × 960，最小窗口为 800 × 650。

### Word 提图与命名

1. 选择「Word 提图」，导入 `.docx` 文件。
2. 在左侧选择图片，在右侧「文档参考描述」查看同段文字、上下各两段及替代文字。
3. 点击命名参考填入名称，按需修改，再点击「保存名称」。扩展名保持不变。
4. 按需编辑图片，然后导出保存。

兼容 DrawingML 和旧式 VML；同一图片的重复引用只提取一次，多处引用的描述分别保留。候选名称会清理非法字符并截短，重名时提示冲突。

### 证书批量截图

1. 选择顶部「证书截图」，点击「证书自动截图」。**每次打开都从空证书队列开始**，清除上轮证书选择、导入错误和弹窗处理记录；已生成的图片和当前页面中的字段设置保留。
2. 点击「批量添加证书」，选择 PEM 或 DER 格式的 `.crt`、`.cer`、`.pem` 公钥证书。队列最多 **50 张**，每张最多 **2 MB**，相同证书按 SHA-256 去重；无效文件逐个提示原因。
3. 选择截图项目，最多 **24 项**：默认提供常规信息、颁发者、公钥、证书路径；还可选择序列号、签名算法、使用者、有效期、指纹、密钥用法等字段。
4. 自定义字段填写 Windows 证书「详细信息」中的字段名称，例如「证书策略」。按名称精确匹配，不区分英文大小写。常用字段提供简体中文、繁体中文及英文别名；其他语言可输入系统实际显示的名称。
5. 点击「开始截图」，保持 Windows 桌面解锁，避免操作自动打开的证书窗口。逐张处理，每张最多等待 75 秒；「停止后续处理」会等待当前证书结束，再跳过剩余证书。
6. 同一证书的全部所选项目完成后才加入图片列表；失败时不导入该证书的部分截图，并继续下一张。截图按「证书名称_字段.png」命名，同名输出自动编号。
7. 任务结束后自动关闭配置弹窗，工作区保留摘要及失败详情，明确显示证书文件、失败字段和具体原因。成功/失败的证书份数与生成的图片张数分开统计。

证书导入兼容 SM2 等 Python 解析库暂不支持的公钥参数：不因公钥元数据解码失败而拒绝整张证书，保留原始证书供 Windows 查看，界面会说明元数据解析限制。

证书截图在独立进程中通过控件定位完成，不使用固定鼠标坐标。只操作任务自己的只读证书窗口，不安装证书、不修改系统信任设置；证书链使用 Windows 本机证书库和缓存，不联网补链。窗口中的链状态会如实保留，截图不代表额外验证了证书有效性。截图需要可交互的 Windows 桌面；网页模式也可使用，但 Python 后端必须运行在该 Windows 本机。

不支持私钥文件、PFX、多证书 PEM 集合，也不自动打开颁发者证书本身。字段缺失会明确报错；窗口初始化或截图导入失败时显示对应阶段，不虚构失败字段。

### 图片导出

点击「导出全部图片」，选择以下方式：

- **保存到文件夹**：保留当前文件名、格式和编辑结果。客户端使用原生目录选择框；网页优先使用浏览器目录接口，接口不可用或被 WebView 拦截时，可由 Windows 本机 Python 服务打开系统选择框。也可直接点击「使用系统文件夹选择框」。
- **保存为 ZIP 压缩包**：客户端选择 ZIP 保存位置；网页支持保存选择器时直接写入指定 ZIP 文件，否则通过浏览器下载。

目录导出遇到同名文件会停止，避免覆盖，请选择空目录。系统选择框仅允许本机请求，最多等待 180 秒；取消不写入文件。直接保存成功后才提示完成；浏览器下载回退只提示已开始下载。浏览器写入中断时会提示已保存数量，未完成的文件可能不完整。

## 开发运行

需要 Python 3.10+、Node.js 22.12+。以下命令在项目根目录运行，示例使用 PowerShell。

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

访问 [开发页面](http://127.0.0.1:5173)。Vite 将 `/api` 代理到本机 8000 端口，前后端可分别开发。

### 网页构建预览

```powershell
cd frontend
npm ci
npm run build
cd ..
python -m backend
```

访问 [网页预览](http://127.0.0.1:8000)，由后端提供构建后的静态文件。`?menu=word`、`?menu=local`、`?menu=certificate` 可指定初始菜单；证书入口会打开新的空队列配置窗口。

通常先通过网页确认改动，再按需打包客户端。项目代码中的默认后端端口为 8000，开发调试时临时启动的其他端口不属于分发配置。

### 桌面源码运行与打包

```powershell
python -m pip install -r requirements/desktop.txt
cd frontend
npm ci
npm run build
cd ..
python desktop.py
```

在已激活虚拟环境的终端执行 Windows 打包脚本：

```powershell
.\scripts\build-desktop.ps1
```

脚本会准备图标、安装前端锁定依赖、运行前端测试、构建网页和客户端，再检查正常/小窗口下的图片列表，最后生成：

```text
dist/2.6.0/Shitu/拾图 Shitu.exe
dist/Shitu-2.6.0-Windows-x64.zip
```

按版本输出目录，避免覆盖正在运行的旧版。已有前端构建资源时，也可仅执行：

```powershell
python -m PyInstaller --noconfirm --distpath dist/2.6.0 packaging/desktop.spec
```

图标资源位于 `assets/branding/`，Windows 文件属性包含产品名称和版本。客户端必须在目标操作系统上分别构建；当前发布和原生证书截图验证面向 Windows x64。签名、安装器等发布流程需另行配置。

## 测试与日志

```powershell
python -m pip install pytest httpx
python -m pytest tests -q
cd frontend
npm test
npm run build
```

安装桌面依赖后，可在项目根目录检查真实 WebView：

```powershell
python desktop.py --gui-smoke-test
python desktop.py --gui-smoke-test --test-image-list
python desktop.py --gui-smoke-test --test-image-list --compact-window
python desktop.py --gui-smoke-test --test-certificate "data/example.crt"
python desktop.py --gui-smoke-test --test-certificate "tests/fixtures/sm2.crt"
```

- `--gui-smoke-test` 启动真实窗口，验证界面、图片和桌面桥接后自动退出。
- `--test-image-list` 创建 80 张合成图片，检查列表独立滚动和页面尺寸。
- `--test-certificate` 执行四类真实证书截图，需要未锁定的 Windows 桌面。
- `--smoke-test` 只检查本地 HTTP 服务。打包后的 EXE 同样支持这些参数。

2.6.0 发布前已通过 **59 项 Python 测试、27 项前端测试**，以及打包后正常/小窗口启动、RSA 和 SM2 证书截图检查。`tests/fixtures/sm2.crt` 是自动化测试用自签名公钥证书，仓库不包含对应私钥。

客户端日志位于 `%LOCALAPPDATA%/Shitu/logs/desktop.log`，源码桌面运行时位于项目 `logs/`。窗口初始化超过 30 秒会生成 `startup-threads.log`。网页源码运行的后端错误输出到启动终端；服务器错误返回 JSON，前端也能处理非 JSON 错误响应，不直接展示 JSON 解析异常。

## 目录与接口

```text
picture-tools/
├── assets/branding/       品牌图标与 Windows ICO
├── backend/              API、Word/图片处理、证书任务与系统目录选择
├── frontend/             Vue 页面、样式、测试和 Vite 配置
├── packaging/            PyInstaller 配置与 Windows 版本信息
├── requirements/         Web、桌面依赖清单
├── scripts/              图标准备与客户端构建脚本
├── tests/                Python 测试及公开测试证书
├── docs/                 设计与历史方案
├── data/                 本地输入文档/证书（不提交）
├── logs/                 本地日志和调试产物（不提交）
├── dist/                 客户端与分发 ZIP（不提交）
├── desktop.py            桌面入口及独立任务入口
├── requirements.txt      默认 Web 后端依赖入口
└── CHANGELOG.md          版本更新记录
```

`.venv/`、`node_modules/`、`frontend/dist/`、`build/`、Python 缓存等均可重新生成，不纳入 Git。`.worktrees/` 中已有的 Electron 分支工作区保留，历史方案不代表当前发布架构。

| 接口 | 用途 |
| --- | --- |
| `POST /api/import` | 导入 Word 或图片，multipart `files` |
| `GET/DELETE /api/images` | 列出/清空工作区图片 |
| `PATCH/DELETE /api/images/{id}` | 重命名/删除图片 |
| `POST /api/images/{id}/edit` | 原图编辑 |
| `GET /api/images/{id}/file` | 原始图片、预览或缩略图 |
| `GET /api/export` | 全部图片 ZIP |
| `POST /api/export/folder` | 本机系统目录选择与导出，不接收客户端指定路径 |
| `POST /api/certificates/import` | 批量导入公钥证书 |
| `GET/DELETE /api/certificates` | 列出/清空证书队列 |
| `DELETE /api/certificates/{id}` | 移除单张证书 |
| `POST /api/certificate-jobs` | 启动批量任务，指定证书 ID 和字段 |
| `GET /api/certificate-jobs` | 最近一次任务 |
| `GET /api/certificate-jobs/{id}` | 查询进度及逐证书结果 |
| `POST /api/certificate-jobs/{id}/cancel` | 停止后续处理 |

后端运行后可访问 `/docs` 查看接口定义。桌面版采用随机本机端口和每次启动的访问令牌，桥接仅开放目录导入、目录导出、ZIP 导出、证书选择与证书截图五个方法，内部应用和窗口对象不暴露给前端。

## 数据和使用边界

- 文件位于进程临时工作区，退出后清理。浏览器刷新保留同一后端进程中的图片，服务重启后不保留；退出或重启前请导出。桌面关闭时会提示未导出的工作区。
- 每个后端进程对应一个共享工作区，面向单机单用户，不应直接作为公网多用户服务。
- 图片单文件最多 50 MB，图片工作区和 DOCX 解压总量最多 250 MB。
- 支持 `.docx`，不支持旧式 `.doc`；提取按正文 XML 引用顺序，不保证浮动图片的视觉顺序，不读取页眉页脚、外链图片或图片内部文字。Pillow 无法解码的矢量媒体会跳过。
- 文字描述按段落结构关联，遇到其他图片、表格单元格或结构边界时停止。它只是命名参考，复杂浮动布局和表格需人工核对。
- GIF/TIFF 编辑处理当前首帧，直接导出未编辑的图片保留原始文件；JPEG 质量可调，PNG 使用无损优化，压缩不保证每次都减小文件。
- Windows 证书窗口若缺少字段或无法定位，任务会报告原因；其他系统语言和系统证书提供程序的显示能力需按实际环境核对。

版本变化见 [CHANGELOG.md](CHANGELOG.md)。
