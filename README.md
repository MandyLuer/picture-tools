# 拾图 Shitu

本地图片与证书截图工作台，当前版本 **2.8.0**。前端使用 Vue 3，后端使用 FastAPI，桌面版由 pywebview 和 PyInstaller 提供。支持网页预览和 Windows 客户端分发，旧版 Tkinter 实现可通过 Git 历史查看。

## 功能入口

| 菜单 | 功能 |
| --- | --- |
| Word 提图 | 按 Word 正文引用顺序提取图片，显示附近文字和命名参考；支持 `.docx` 与 `.doc` |
| 本地图片 | 导入图片或文件夹，重命名、旋转、翻转、裁剪和压缩 |
| 证书截图 | 批量添加证书，选择常用或自定义字段，生成真实 Windows 证书窗口截图 |

顶部菜单下方只显示对应操作；左侧是独立滚动的图片列表，右侧用于预览和编辑，右上方提供 AI 配置、导出和清空。图片名称旁提供「AI 生成名称」和「批量命名」。菜单共用图片工作区，切换时保留图片、名称草稿及编辑状态。

- 支持方向键切图、滚轮缩放、拖动平移、收起参考描述，以及大图模式（Esc 退出）。
- 重命名只更新当前名称，保留当前缩放与选区；相邻图片异步预加载，预览按图片版本缓存（后端缓存最多 64 MB）。编辑后自动失效，避免显示旧图；大图预览解码不阻塞名称更新。
- 旋转、翻转、裁剪和压缩直接更新当前工作区图片，保留名称、顺序和文档参考描述，不生成编辑副本。
- 单张删除和清空工作区均需要界面确认，不删除原始文档或导入前的本地图片。
- 加载、编辑、删除、清空和导出后显示操作提示；取消操作不提示成功。

## AI 智能命名

本地图片支持两种减少重复文件的保存方式：网页端在「本地图片」点击「关联原文件夹」，通过系统选择框导入；桌面端「选择本地图片文件夹」会自动关联原文件。普通文件上传不会取得原文件位置。

- 勾选名称下方「原位置重命名」后，点击「保存名称」直接修改原文件名，不需要导出。若图片已有编辑，请使用文件夹导出保存编辑结果。批量 AI 应用仍先更新工作区名称，可逐张保存到原位置。
- 导出窗口默认保留原图。主动勾选「保存到文件夹后删除关联原图」，选择新的目标文件夹后，全部导出并校验成功才删除关联原图；ZIP 不支持删除原图。原图变化、同名冲突或写入失败时保留原图；部分原图无法删除会列出文件名。
- 关联仅在当前后端会话中有效；工作区删除和清空仍只删除临时图片，不删除关联原图。

此功能已包含在 2.7.0 客户端和网页中。点击工作区右上方「AI 配置」：

1. 选择配置：本地模型（Ollama / LM Studio 等）、OpenAI、DeepSeek 或其他兼容 API。各配置独立保存，随时切换。
2. 填写接口**基础地址**和可用模型名称，不要填写完整 `/chat/completions` 路径。Ollama 默认地址为 `http://127.0.0.1:11434/v1`，OpenAI 为 `https://api.openai.com/v1`，DeepSeek 为 `https://api.deepseek.com`。本地服务不要求认证时，密钥可以留空。
3. 选择「发送图片」「使用文档描述」、顺序编号、前缀、命名要求和超时时间，点击「保存配置」。纯文本模型应关闭发送图片；若图片没有可用描述，会提示启用视觉模型。
4. 可点击「测试连接」，只发送合成测试图或简短文本，不发送工作区内容。
5. 关闭配置窗口，在图片名称旁点击「AI 生成名称」，直接将当前图片的建议名称填入输入框，并显示命名依据。可继续修改，点击「保存名称」后才生效。尚未配置模型时，会自动打开配置窗口。
6. 多张处理可点击名称旁的「批量命名」，每批最多 100 张。对照原名、建议名和命名依据，修改候选名称并勾选，再点击「应用所选名称」。模型标记为需要人工核对的结果默认不勾选；支持停止后续处理，刷新后重新打开批量命名可恢复最近任务。

生成建议不会修改文件名。应用前会检查工作区、图片版本及名称是否变化，并校验最终重名；校验不通过时整批不应用。应用只更新工作区名称，保留扩展名、图片内容、选中项和缩放状态。

所有模型请求由 Python 后端发起，使用 OpenAI-compatible Chat Completions 协议。发送图片时使用最长边不超过 1280 像素的 JPEG 副本，不修改原图；文档描述按配置发送。使用远程 API 时，这些内容会发送给对应服务。是否支持图片取决于所选模型。ChatGPT 的聊天订阅与 API 额度分开，需要单独可用的 API Key。

配置保存于源码运行目录的 `.local/ai-settings.json`，桌面分发运行时位于 `%LOCALAPPDATA%/Shitu/ai-settings.json`。密钥默认只在本次后端运行期间保留；Windows 可勾选「记住密钥」使用 DPAPI 加密保存。设置接口不回传密钥，切换接口地址会清除旧地址的密钥，避免带到新服务。请在界面配置密钥，不要提交到 Git 或贴入聊天。

接口依据：[OpenAI Chat Completions](https://developers.openai.com/api/reference/resources/chat/subresources/completions/methods/create)、[DeepSeek 视觉输入](https://api-docs.deepseek.com/guides/vision/)、[Ollama OpenAI 兼容接口](https://docs.ollama.com/api/openai-compatibility)。本次完成了本机模拟模型的 HTTP 与浏览器全流程联调；真实服务需配置你自己的模型名称和密钥后测试。

## PDF 合并

在「PDF 合并」菜单中多选 PDF，可上下调整顺序并填写输出名称。点击「合并并保存」后，Windows 目录选择框默认定位到第一份 PDF 所在目录。同名时自动编号，原文件保留。

## 使用客户端

Windows 分发包为 `dist/Shitu-2.8.0-Windows-x64.zip`。解压整个文件夹后运行 **拾图 Shitu.exe**，请保留同目录的依赖文件，不要只复制 EXE。

客户端不要求安装 Python 或 Node.js，需要 Windows WebView2 Runtime。程序默认窗口为 1440 × 960，最小窗口为 800 × 650。

### Word 提图与命名

1. 选择「Word 提图」，导入 `.docx` 或 `.doc` 文件。旧式 `.doc`（Word 97–2003 二进制或 RTF）会在独立子进程中转换为 `.docx`，再与 `.docx` 走同一条提取流程；转换需要本机安装 Microsoft Word 或 LibreOffice，未安装时会提示先另存为 `.docx`。
2. 在左侧选择图片，在右侧「文档参考描述」查看同段文字、上下各两段及替代文字。
3. 点击命名参考填入名称，按需修改，再点击「保存名称」。扩展名保持不变。
4. 按需编辑图片，然后导出保存。

兼容 DrawingML 和旧式 VML；同一图片的重复引用只提取一次，多处引用的描述分别保留。候选名称会清理非法字符并截短，重名时提示冲突。文档加密、损坏或转换超时会给出具体原因，该批导入整体回滚，不影响工作区已有图片。

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
dist/2.8.0/Shitu/拾图 Shitu.exe
dist/Shitu-2.8.0-Windows-x64.zip
```

按版本输出目录，避免覆盖正在运行的旧版。已有前端构建资源时，也可仅执行：

```powershell
python -m PyInstaller --noconfirm --distpath dist/2.8.0 packaging/desktop.spec
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

2.7.0 发布前已通过 **98 项 Python 测试、37 项前端测试**，以及打包后 HTTP 启动、正常/小窗口图片列表检查。AI 命名与原文件保存流程已用模拟模型和临时图片完成网页联调。此前 2.6.2 已验证真实 DOC/DOCX 图片与题注提取、RSA 和 SM2 证书截图；本次未重复执行这些原生流程。`tests/fixtures/sm2.crt` 是自动化测试用自签名公钥证书，仓库不包含对应私钥。

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

`.local/`（AI 本机配置）、`.venv/`、`node_modules/`、`frontend/dist/`、`build/`、Python 缓存等均可重新生成，不纳入 Git。`.worktrees/` 中已有的 Electron 分支工作区保留，历史方案不代表当前发布架构。

| 接口 | 用途 |
| --- | --- |
| `GET/PUT /api/ai/settings` | 读取脱敏配置 / 保存模型配置 |
| `POST /api/ai/test` | 用合成输入测试模型连接 |
| `GET/POST /api/ai/jobs` | 恢复最近任务 / 生成命名建议 |
| `GET /api/ai/jobs/{id}` | 查询命名进度及候选结果 |
| `POST /api/ai/jobs/{id}/cancel` | 停止后续命名 |
| `POST /api/ai/apply` | 校验并应用所选名称 |
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
- 图片与 Word 文档单文件最多 50 MB；图片工作区最多 250 MB，Word 文档解压总量同样最多 250 MB。
- 支持 `.docx` 与旧式 `.doc`（Word 97–2003 二进制或 RTF）。`.doc` 依赖本机安装的 Microsoft Word 或 LibreOffice，在独立子进程中转换后按 `.docx` 提取；文档加密、损坏或转换超过 90 秒会失败并提示原因。提取按正文 XML 引用顺序，不保证浮动图片的视觉顺序，不读取页眉页脚、外链图片或图片内部文字。Pillow 无法解码的矢量媒体会跳过。
- 文字描述按段落结构关联，遇到其他图片、表格单元格或结构边界时停止。它只是命名参考，复杂浮动布局和表格需人工核对。
- GIF/TIFF 编辑处理当前首帧，直接导出未编辑的图片保留原始文件；JPEG 质量可调，PNG 使用无损优化，压缩不保证每次都减小文件。
- Windows 证书窗口若缺少字段或无法定位，任务会报告原因；其他系统语言和系统证书提供程序的显示能力需按实际环境核对。

版本变化见 [CHANGELOG.md](CHANGELOG.md)。
