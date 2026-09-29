# Word 图片提取与重命名工具 — Electron 客户端重构设计

- 文档日期：2026-09-29
- 状态：已通过分段评审，待用户审阅本文档
- 目标平台：Windows（仅）
- 现有实现：`image-editer.py`（Python 3 + Tkinter，1093 行，单文件）

## 1. 背景与目标

### 1.1 现有实现

仓库当前是一个单文件 Tkinter 桌面程序，`image-editer.py` 内包含两个类：

- `ZoomableImage`（第 12–461 行）：画布上的图片显示与编辑，支持滚轮缩放、拖拽平移、旋转、水平/垂直翻转、框选裁剪、压缩。
- `WordImageExtractorApp`（第 463–1088 行）：主窗口，两个选项卡——「从 Word 提取图片」与「图片重命名」。

已实现的能力：

| 能力 | 现有实现位置 | 说明 |
|---|---|---|
| 从 docx 提取图片 | `extract_images` | 解压 `word/media/`，按文档顺序重命名输出为 `001.jpg` 形式 |
| 图片保序 | `get_image_order_from_docx` | 解析 `document.xml` + `document.xml.rels`，含 `w:pict`/VML 回退，再回退到正则，最后按文件名数字排序 |
| 图片重命名 | `save_rename` | 单张编辑，输入框 + 上一张/下一张，按文件名数字排序 |
| 旋转 | `rotate_image` | 左转/右转 90°，基于 `original_image` 旋转 |
| 翻转 | `flip_horizontal` / `flip_vertical` | 水平/垂直翻转 |
| 裁剪 | `start_cropping` / `crop_image` | 画布上框选，确认后另存 |
| 压缩 | `compress_image` | 按质量另存 |

### 1.2 已记录的缺陷（README TODO 与代码事实）

1. **裁剪后文件名丢失**：`save_cropped_image`（第 246 行）从当前文件路径派生新文件名，用户重命名后的名称在裁剪流程中被覆盖或丢失。
2. **输出目录被清空**：`extract_images`（第 722–729 行）在写入前无条件删除输出目录下的所有文件，用户放在该目录里的其他文件会被删除。
3. **编辑产生碎片文件**：旋转、翻转、裁剪、压缩都不修改原文件，而是在旁边生成 `_cropped`、`_compressed` 副本，文件夹很快变乱，且没有聚合入口。
4. **`imghdr` 已废弃**：`imghdr` 模块在 Python 3.13 被移除，第 6 行的 `import imghdr` 与第 740、930、1036 行的调用在新 Python 上会直接失败。
5. **排序对非数字文件名会抛异常**：`load_image_files`（第 937 行）用 `int("".join(filter(str.isdigit, x)) or "0")` 排序，纯文字文件名依赖 `or "0"` 兜底，处理不严谨。
6. **未保存改动无保护**：所有编辑即时写盘，关闭窗口即丢失，且没有撤销/重做。
7. **界面局限**：Tkinter 无法实现遮罩式裁剪、虚拟滚动列表、深色主题、快捷键体系等现代交互。

### 1.3 本次目标

1. **保留能力**：提取、保序、重命名、旋转、翻转、裁剪、压缩、打开文件夹等能力全部保留；允许重构这些能力的**行为**以修正上述缺陷。
2. **重新设计界面**：三栏工作台（图片列表 + 预览编辑 + 属性面板），基于 Element Plus 与统一设计令牌，支持浅色/深色主题。
3. **技术栈**：Python 后端 + JS 前端 + Electron 打包为 Windows 客户端。

### 1.4 非目标

- 不做 macOS / Linux 构建（架构不阻碍，但本次不验证、不交付）。
- 不做多用户、账号、云同步。
- 不支持旧版 `.doc`（二进制格式），仅支持 `.docx`。
- 不做图片内容识别、自动命名等智能功能。
- 不做与 Word 的双向写回（不修改原始 docx）。

## 2. 架构总览

### 2.1 进程与分层

```
┌──────────────────────────────────────────────────────────────┐
│ Electron 主进程 (Node)                                        │
│  · 窗口与生命周期、原生文件/目录对话框、应用菜单               │
│  · 拉起并监管 Python 后端子进程（崩溃重启、退出清理）          │
│  · 注册 appimg:// 自定义协议，按白名单提供本地图片            │
│  · contextBridge 暴露白名单 API，关闭 nodeIntegration         │
└───────────────┬──────────────────────────┬───────────────────┘
                │ IPC (invoke/on)          │ stdio 行分隔 JSON-RPC
┌───────────────▼──────────────────┐       │
│ 渲染进程 (Vue 3 + Element Plus)   │       │
│  · 三栏工作台、Pinia stores       │       │
│  · transport.js 抽象传输层        │       │
│    (stdio 通道 | 开发期 HTTP)     │       │
└──────────────────────────────────┘       │
                                ┌──────────▼────────────────────┐
                                │ Python 后端 (无 GUI 依赖)      │
                                │  __main__.py: --stdio         │
                                │             | --serve         │
                                │             | --selftest      │
                                │  rpc.py: 静态方法表分发        │
                                │  services/: docx_extract,     │
                                │  image_ops, library           │
                                │  security.py: 路径白名单       │
                                └───────────────────────────────┘
```

### 2.2 关键决策

**后端作为无 GUI 的独立子进程。** 剥离所有 Tkinter 依赖，核心逻辑成为普通函数。收益有三：可在 Electron 内以 stdio 运行，可在浏览器开发模式下以 HTTP 调用，可被 `--selftest` 与 pytest 直接调用。这是"同一份逻辑三处可达"，也是让回归测试变便宜的前提。

**图片不走 base64。** 预览大图若经 base64 传入渲染进程会明显拖慢加载并放大内存占用。主进程注册 `appimg://` 自定义协议，仅允许读取后端已授权的目录内文件，渲染进程用 `<img src="appimg://...">` 直接加载，复用 Chromium 自身的流式解码与缓存。

**双传输模式。** `--stdio` 为生产路径（无端口、无防火墙弹窗、随 Electron 生命周期）；`--serve` 为开发路径（本地 HTTP，允许前端在纯浏览器中调试与热更新）。两者共用同一个 `rpc.py` 分发器，传输层是唯一差异。

### 2.3 目录结构

```
picture-tools/
├─ app/
│  ├─ main/
│  │  ├─ main.js          窗口、菜单、生命周期
│  │  ├─ backend.js       启动/监控/重启 Python 子进程
│  │  ├─ ipc.js           渲染进程 ↔ 后端的请求转发
│  │  └─ protocol.js      注册 appimg:// 协议
│  ├─ preload/preload.js  contextBridge 白名单 API
│  └─ renderer/           Vue 3 + Vite + Element Plus
│     ├─ src/main.js
│     ├─ src/App.vue
│     ├─ src/api/         transport.js（stdio / http）
│     ├─ src/stores/      library.js / editor.js / tasks.js
│     ├─ src/components/  三栏与各面板组件
│     └─ src/styles/      设计令牌（浅色/深色）
├─ backend/
│  ├─ __main__.py         入口：--stdio | --serve | --selftest
│  ├─ rpc.py              JSON-RPC 分发
│  ├─ transport_stdio.py  行分隔 JSON over stdin/stdout
│  ├─ transport_http.py   开发用 HTTP（127.0.0.1 + 随机端口 + token）
│  ├─ services/
│  │  ├─ docx_extract.py  提取与保序
│  │  ├─ image_ops.py     旋转/翻转/裁剪/压缩/另存
│  │  └─ library.py       扫描、排序、重命名、冲突处理
│  ├─ security.py         路径白名单
│  ├─ errors.py           错误码表
│  ├─ cache.py            预览渲染缓存（LRU + 落盘目录）
│  └─ tests/              pytest + 样例 docx
├─ docs/superpowers/specs/
├─ image-editer.py        旧 Tkinter 版本（M1–M5 期间保留）
├─ package.json
├─ vite.config.js
├─ electron-builder.yml
└─ 奶牛猫.ico
```

## 3. 后端服务层

后端由三个服务组成，**对 Electron 与 HTTP 均无感知**；`rpc.py` 负责方法名到函数的映射，两种传输调用同一分发器。所有函数只接收与返回可 JSON 序列化的普通数据。

### 3.1 `services/docx_extract.py`

替代现有 `extract_images` 与 `get_image_order_from_docx`。

- `inspect_document(path) -> {media: [{relId, partName, bytes, sniffedFormat}], order: [...], orderSource: "xml" | "regex" | "filename" | "guess"}`
- 保序沿用现有三层策略（`document.xml` + rels 解析 → 正则回退 → 文件名数字排序），但**额外上报排序来源**，使界面能在排序为推测结果时给出警告。
- **替换 `imghdr`**：改为读取文件头魔数自主识别 PNG / JPEG / GIF / BMP / TIFF / WebP，避免 Python 3.13 起模块缺失。
- **不再清空输出目录**：写入时采用冲突安全命名，返回值明确区分 `written`、`skipped`、`renamed` 三类结果。

### 3.2 `services/image_ops.py`

替代 `ZoomableImage` 中的编辑方法。设计为「Pillow 图像 + 操作栈」的纯函数，因此同一份操作栈既可重放也可断言。

支持的原子操作：

| 操作 | 参数 | 说明 |
|---|---|---|
| `rotate` | `angle` | 90 的整数倍，`expand=True` |
| `flip` | `axis: horizontal \| vertical` | 转置 |
| `crop` | `rect`（源图像坐标系） | 坐标由前端计算并回传 |
| `compress` | `quality`, `format` | 仅影响编码，不改变像素 |

操作栈**只包含上表四项**。缩放不属于编辑操作——它是预览的视图状态，不进 `ops`、不参与签名、不影响保存结果；裁剪已足够改变输出尺寸。后续若确需独立的缩放编辑，须另行评审并扩充本表。

接口：

- `apply_ops(source_path, ops) -> Image`：**「编辑后的图片长什么样」的唯一真相来源**。
- `save(source_path, target_path, ops, format, quality)`：只在保存时编码一次。
- `image.signature(path, ops) -> str`：对「源文件 + 操作栈」计算稳定哈希，用于脏标记。

**裁剪前移到前端**：浏览器负责遮罩与选框交互，仅把源图像坐标回传后端。这同时修掉了「裁剪后文件名丢失」——后端不再自行推导文件名，也不会再生成任何 `_cropped` 文件。

### 3.3 `services/library.py`

替代 `load_image_files`、排序逻辑、`save_rename`、`open_image_folder`。

- `list_images(folder) -> [{name, size, mtime, width, height}]`：按文件名前缀数字排序，对非数字名有稳定回退，不再抛异常。
- `rename(folder, from, to, onConflict) -> {path}`：冲突策略显式化——`reject`（默认）/ `autoSuffix` / `overwrite`。不存在静默覆盖。
- `batch_rename(folder, entries, onConflict) -> {renamed, conflicts}`：**两阶段提交**（先规划、后应用），批量中途遇到冲突不会留下改了一半的目录。
- `reveal_in_explorer(path)`：仅 Windows，`os.startfile`。对应 RPC 方法 `system.reveal`（服务层函数名与 RPC 方法名不必一致，映射关系在 `rpc.py` 中一处声明）。

### 3.4 RPC 方法表

| 方法 | 参数 | 返回 |
|---|---|---|
| `app.ping` | — | `{version, python, pillow}` |
| `docx.inspect` | `{path}` | 媒体清单 + 顺序 + `orderSource` |
| `docx.extract` | `{path, outDir, pattern, startIndex}` | `{written[], skipped[], renamed[]}` |
| `library.list` | `{dir}` | 图片条目数组 |
| `library.rename` | `{dir, from, to, onConflict}` | `{path}` |
| `library.batchRename` | `{dir, entries, onConflict}` | `{renamed[], conflicts[]}` |
| `image.preview` | `{path, ops, maxEdge}` | `{cacheKey}`（经 `appimg://` 渲染） |
| `image.signature` | `{path, ops}` | `{signature}` |
| `image.save` | `{path, targetPath, ops, format, quality}` | `{path, bytes}` |
| `image.batchSave` | `{items[], outDir, format, quality}` | `{saved[], failed[]}` |
| `system.reveal` | `{path}` | `{ok}` |
| `task.cancel` | `{taskId}` | `{ok}` |

上表全部由 **Python 后端**实现。以下方法由 **Electron 主进程**实现（渲染进程看来是同一套调用接口，`transport.js` 内部按方法名前缀路由）：

| 方法 | 参数 | 返回 |
|---|---|---|
| `dialog.pickDocx` | — | `{path}` 或 `{cancelled: true}` |
| `dialog.pickFolder` | `{title}` | `{path}` 或 `{cancelled: true}` |
| `shell.openFolder` | `{path}` | `{ok}` |
| `backend.restart` | — | `{ok}` |

约定：

- 原生文件/目录选择器由 Electron 主进程实现（不走 Python），以便使用标准 Windows 对话框并让后端无需感知界面。
- 长任务（`docx.extract`、`image.batchSave`）立即返回 `taskId`，随后通过 `task.progress` / `task.done` 通知推送进度与结果，不阻塞单次请求。

## 4. 前端界面设计

### 4.1 骨架

顶部工具栏 + 三栏主体 + 底部状态栏。

- 工具栏：左侧「导入 Word」「打开文件夹」；中间搜索框、排序方式、缩放比例；右侧「保存」「全部保存」、主题切换。
- 状态栏：图片总数、已选中数量、未保存改动数量、后端连接状态。

### 4.2 左栏 · 图片列表（260–320px，可拖拽调宽，可折叠）

- 列表项：缩略图 + 文件名 + 尺寸。缩略图由后端按需生成，不加载原图。
- 多选（Ctrl / Shift）、搜索过滤、排序切换（原始顺序 / 文件名 / 修改时间 / 尺寸）。
- 有未保存改动的项右侧显示圆点标记。
- 底部固定「已选 N 张」与批量重命名入口。
- 虚拟滚动，大文件夹下仅渲染可见项。

### 4.3 中栏 · 预览与编辑区（自适应剩余宽度）

- 画布支持滚轮缩放、按住空格或中键拖拽平移，缩放范围 10%–400%，左上角浮层显示当前缩放比例。
- **裁剪为画布内直接拉框**：进入裁剪模式后出现半透明遮罩，框外变暗，八个控制点可调整，Enter 确认、Esc 取消。这是现有 Tkinter 版无法实现的交互。
- 底部工具条：旋转 90°、水平/垂直翻转、裁剪、压缩。点击后**立即在预览中体现但不写盘**，每步进入操作栈，可撤销/重做。
- 空状态、加载中、加载失败三种状态各有占位表现，不再只绘制一句灰字。

### 4.4 右栏 · 属性与操作面板（约 320px）

三个折叠分区：

1. **文件信息**——完整路径、原始尺寸、文件大小、格式、修改时间。
2. **重命名**——文件名输入框（不含扩展名，扩展名以固定标签显示在右侧）、冲突策略下拉、「应用」。支持选区批量重命名，提供「前缀 + 起始序号 + 位数」与「查找替换」两种模式，并实时显示重命名前后对照。
3. **编辑参数**——压缩质量滑块（附预计体积变化提示）、旋转/翻转按钮、裁剪尺寸精确输入（像素值）。与中栏底部工具条是同一组操作的两个入口，状态保持同步。

### 4.5 视觉语言

- 基于 Element Plus 的 CSS 变量构建一层设计令牌（颜色、间距、字号、圆角、阴影），统一置于 `src/styles/`，组件内不写魔法数字。
- 中性灰底 + 单一强调色（用于选中态与主操作按钮）；危险操作使用红色且必须二次确认。
- 浅色与深色两套主题只替换令牌值。
- 字体优先 `"Microsoft YaHei UI", "Segoe UI"`，保证中文界面字重正常。

### 4.6 键盘可达性

| 快捷键 | 行为 |
|---|---|
| `←` `→` / `↑` `↓` | 切换上/下一张图片 |
| `Ctrl+S` | 保存当前 |
| `Ctrl+Shift+S` | 全部保存 |
| `Ctrl+Z` / `Ctrl+Y` | 撤销 / 重做 |
| `Delete` | 进入删除确认 |
| `F2` | 聚焦重命名输入框 |
| `Ctrl+A` | 列表中全选 |

## 5. 状态管理、保存模型与撤销

### 5.1 三层状态划分

| Store | 职责 | 不负责 |
|---|---|---|
| `stores/library.js` | 文件列表、选中集合、排序、搜索、磁盘指纹 | 任何编辑内容 |
| `stores/editor.js` | 当前图片路径、`ops` 操作栈、`historyIndex`、已保存基线 | 文件夹级信息 |
| `stores/tasks.js` | 任务 id、类型、进度、可取消标志、结果与错误 | 编辑内容 |

`editor.js` 是编辑的唯一真相来源；中栏预览与右栏参数面板均从它派生，不存在两处各存一份的可能。

### 5.2 脏标记

后端 `image.signature(path, ops)` 对「源文件 + 操作栈」计算稳定哈希。渲染进程保存已保存基线签名，当前签名与基线不一致即为未保存。此方案使「旋转四次 90° 又转回原状」能自动回到干净状态，并驱动左栏圆点标记。前端对签名请求加防抖。

### 5.3 预览更新

改动 `ops` 后请求 `image.preview{path, ops, maxEdge}`，后端将渲染结果写入临时渲染缓存并返回 `cacheKey`，前端只替换一次 `<img>` 的 `src`。相同签名命中缓存直接返回。请求带自增序号，**过期响应到达时丢弃**，避免快速连点旋转导致画面回跳。

### 5.4 保存流程（三步）

1. **预检**：格式与扩展名是否匹配（如 JPEG 不支持透明通道，需先合成白底）、目标是否可写、冲突策略是否明确。
2. **写入**：后端先写临时文件再原子替换目标文件。任一步失败，原文件保持不变，不会出现写坏一半的图片。
3. **提交**：成功后更新基线签名，脏标记自动清除；失败则保留 `ops` 栈供重试或另存。

### 5.5 批量保存

先弹出影响预览对话框（将覆盖 N 个文件、其中 M 个会转换格式、是否存在冲突），确认后任务化执行，逐张推进度、可中途取消，结束给出「成功 / 跳过 / 失败」三类清单，失败项可一键定位。

### 5.6 撤销 / 重做

操作栈在渲染进程维护，`historyIndex` 仅移动指针，后端无需感知撤销。栈上限 100 步，超出丢弃最早一步（不丢弃当前状态）。切图、重命名、保存均不清空历史。**关闭窗口或存在未保存改动时必须弹确认**，不再静默丢弃。

### 5.7 内存控制

后端维护 LRU 预览缓存（按总字节数上限，默认 256MB），超出淘汰最久未用项。渲染进程只持有当前一张已解码位图与可见的十余个缩略图。

## 6. 错误处理、安全与性能

### 6.1 错误码表

替代现有「后端抛什么前端弹什么」的模式。结构化错误格式：`{code, message, detail, hint, context}`——`message` 为面向用户的中文说明，`detail` 为技术细节（默认折叠、可复制），`hint` 为下一步建议。

| 错误码 | 用户提示 | 建议动作 |
|---|---|---|
| `E_DOCX_NOT_FOUND` | 找不到这个 Word 文件 | 重新选择 |
| `E_DOCX_NOT_ZIP` | 该文件不是有效的 .docx（.doc 旧格式不支持） | 用 Word 另存为 .docx |
| `E_DOCX_NO_MEDIA` | 文档中没有嵌入图片 | 确认内容 |
| `E_DOCX_ORDER_GUESSED` | 未能读取图片插入顺序，已按文件名排序 | 警告级，不阻断；可手动排序 |
| `E_OUTDIR_NOT_WRITABLE` | 输出目录不可写（可能被占用或无权限） | 更换目录 |
| `E_NAME_EXISTS` | 目标文件名已存在 | 换名或选择冲突策略 |
| `E_NAME_INVALID` | 文件名含非法字符 `\ / : * ? " < > \|` | 自动给出合法名 |
| `E_IMAGE_DECODE` | 图片已损坏或格式不受支持 | 跳过该文件 |
| `E_FORMAT_MISMATCH` | JPEG 不支持透明背景 | 合成白底或取消 |
| `E_DISK_FULL` | 磁盘空间不足 | 清理空间 |
| `E_TASK_CANCELLED` | 已取消（非错误态展示） | — |
| `E_BACKEND_DOWN` | 后端进程已退出 | 一键重启 |
| `E_INTERNAL` | 内部错误 | 复制诊断信息 |

**硬性约束**：任何单张图片的失败都不能中断整批任务。批量操作结束后给出「成功 / 跳过 / 失败」三份清单，失败项可逐条查看原因并定位文件。

### 6.2 安全边界（四层）

1. **后端路径白名单**：后端只允许访问用户在原生对话框中选中的路径，以及已授权目录下的子路径。所有路径先 `Path.resolve()` 再做前缀校验，防 `..` 穿越与符号链接逃逸；其余一律拒绝。
2. **Electron 侧**：开启 `contextIsolation` 与 `sandbox`，关闭 `nodeIntegration`；仅通过 preload 的 `contextBridge` 暴露白名单方法；`will-navigate` 与 `setWindowOpenHandler` 一律拦截，禁止加载任何远程内容；设置 CSP，禁止 `unsafe-eval`。
3. **`appimg://` 协议**：主进程解析请求时校验路径在白名单内且扩展名属于图片格式，否则返回 403，不复用 `file://` 的宽松语义。
4. **HTTP 开发模式**：仅绑定 `127.0.0.1`，端口随机，每次启动生成一次性 token 置于请求头，无 token 一律 401。`--stdio` 模式完全不开端口。RPC 分发器为**静态方法名到函数的映射表**，不做任何动态导入或 `eval`。

### 6.3 性能

- **虚拟滚动**：图片列表在数千文件规模下只渲染可见项。
- **缩略图懒加载**：滚动进入视野才请求生成，首批仅加载前 80 个。
- **提取并行化**：解码与写盘使用线程池（默认 4，上限 8）；docx XML 解析保持单线程，以免顺序错乱。
- **取消机制**：长任务携带取消标志，后端在每张图的处理边界检查；收到取消后先完整写完当前文件再退出，不留半成品。
- **切图资源释放**：切图时先释放上一张解码位图引用再加载新图，与后端 LRU 缓存形成双保险。

性能目标：

| 指标 | 目标 |
|---|---|
| 首屏（点击图标到可操作） | ≤ 3 秒 |
| 打开 500 张图的文件夹到列表可滚动 | ≤ 1 秒 |
| 大图（4000×3000）预览响应 | ≤ 500 毫秒 |

### 6.4 Windows 专有处理

- **中文与空格路径**：子进程参数以数组传递，绝不拼接 shell 字符串；Python 侧统一 `UTF-8`，读写文件显式指定 `encoding="utf-8"`。
- **长路径**：Python 侧对超过 260 字符的路径使用 `\\?\` 前缀，UI 层对超长输出路径提前警告。
- **不弹黑窗**：打包后的后端为 `console=False` 的窗口化程序，通过 stdio 管道通信；仅开发模式显示控制台。
- **降低杀软误报**：使用 PyInstaller `--onedir` 目录模式分发（而非单文件模式），同时加快启动。
- **高 DPI**：在 manifest 声明 DPI 感知，避免 125% / 150% 缩放下的模糊窗口。

## 7. 测试策略

### 7.1 后端（pytest，主要投入）

核心逻辑均为纯函数，测试成本低、价值高。

- `test_docx_extract.py`——保序提取正确性，三条回退路径分别可测。样例 docx 三个：普通图文文档、含 VML/`pict` 旧式图片的文档、**图片顺序与 `media` 文件名顺序故意不一致**的文档（最关键，证明保序逻辑真的生效而非碰巧正确）。
- `test_image_ops.py`——`apply_ops` 的确定性与可重放性；空操作栈返回原图；旋转四次 90° 回到原始尺寸；裁剪坐标映射（给定源图坐标，断言输出像素尺寸与内容）；JPEG 透明通道合成白底。
- `test_library.py`——纯文字文件名排序不抛异常；三种冲突策略行为；批量重命名两阶段提交在冲突时**不产生半成品**。
- `test_security.py`——`..` 穿越、符号链接逃逸、白名单外路径一律拒绝。
- `test_signature.py`——签名对操作栈顺序敏感；「转回原状」能回到干净状态。

### 7.2 前端（vitest）

仅覆盖纯逻辑：脏标记计算、撤销栈边界、冲突策略判断、路径格式化。不做组件快照测试。

### 7.3 端到端（Playwright 驱动 Electron）

两条主干：

1. 导入 docx → 提取 → 列表出现 → 重命名 → 保存 → 断言磁盘文件名确实改变。
2. 打开文件夹 → 旋转与裁剪 → 撤销 → 重做 → 保存 → 断言像素发生变化。

## 8. 打包与分发

```
npm run dev      并发启动 Vite dev server + Electron（后端用系统 Python，热更新）
npm run backend  PyInstaller --onedir 打包后端为 backend/backend.exe
npm run build    构建前端 → 打包后端 → electron-builder 产出 Windows 产物
npm run dist     产出 NSIS 安装包 + 免安装 portable 版
```

要点：

- 复用仓库已有的 `奶牛猫.ico`。
- 版本号统一取自 `package.json`，替换现有 `build.spec` 中手写的 `'WordImageExtractor V1.1'`，避免产物名与代码版本脱节。
- 打包后前端资源走 `file://`，后端可执行文件路径通过 `process.resourcesPath` 解析。
- 采用 `--onedir` 而非单文件模式，启动更快且显著降低杀软误报。
- NSIS 安装包含开始菜单与桌面快捷方式、可选安装路径；portable 版可直接从 U 盘运行。

## 9. 迁移计划

| 里程碑 | 交付内容 | 验收标准 |
|---|---|---|
| **M1 骨架** | 目录结构、Electron 空窗口、Vue 三栏静态布局、设计令牌、`app.ping` 打通 | 界面可启动，状态栏显示后端已连接 |
| **M2 后端逻辑** | `docx_extract` / `image_ops` / `library` 三服务 + pytest 全绿 | 第 7.1 节列出的五个测试文件全部存在且 `pytest` 零失败；`python -m backend --selftest` 退出码为 0 |
| **M3 提取流程** | 原生选择文件、提取（含进度与取消）、左栏列表与缩略图 | 样例 docx 提取结果与旧版一致且顺序正确 |
| **M4 编辑与保存** | 预览、旋转/翻转、画布内裁剪、压缩、非破坏性保存、撤销重做 | 裁剪后文件名不再丢失；不再生成 `_cropped` 文件 |
| **M5 重命名与批量** | 单张重命名、批量重命名、查找替换、冲突策略 | 批量冲突时无半成品；具备撤销与影响预览 |
| **M6 打包发布** | `npm run build` 产出安装包与 portable 版、图标、错误兜底 | 在干净 Windows 机器上双击安装后可跑通主干 |

### 9.1 旧版 `image-editer.py` 的处置

M1–M5 期间原样保留在仓库根目录，随时可回退，且不被任何新代码引用。M6 经用户实际验收后另行决定：移入 `legacy/` 并在 README 标注「已由 Electron 版取代」，或删除。本文档不预设结论。

## 10. 环境与依赖

现状实测：

| 项 | 版本 |
|---|---|
| Python（系统与 `venv`） | 3.10.8 |
| Node | v24.9.0 |
| npm | 10.9.8 |

后端依赖：`pillow` 为必需；`pyinstaller` 为打包必需；开发期新增 `pytest`。`lxml` 与 `python-docx` 均**移除**：现有代码虽在 `requirements.txt` 中声明了它们，但全文未使用（`lxml` 未被导入，`python-docx` 的 `Document` 导入了却从未调用），提取逻辑只依赖标准库 `zipfile` 与 `xml.etree.ElementTree`。移除可减小打包体积。若后续发现确有 XML 解析需求，再按需引入并更新本节。

前端依赖：`vue`、`vite`、`element-plus`、`pinia`、`electron`、`electron-builder`，开发期 `vitest`、`playwright`。

## 11. 风险与缓解

| 风险 | 影响 | 缓解 |
|---|---|---|
| 排序回退路径在某些真实文档上仍失准 | 提取顺序与文档不符 | 上报 `orderSource`，界面告警；提供手动拖拽调整顺序 |
| PyInstaller 产物被误杀或体积过大 | 用户无法安装或启动 | 采用 `--onedir`；在干净机器上验证；必要时提供便携版 |
| 大文件夹（数千张）性能不足 | 界面卡顿 | 虚拟滚动 + 缩略图懒加载 + LRU 缓存，并在 M3 用大数据集实测 |
| 非破坏性编辑的资源占用 | 内存增长 | 操作栈只存参数不存位图；LRU 缓存设字节上限 |
| 后端子进程意外退出 | 功能中断 | 主进程监管与自动重启；`E_BACKEND_DOWN` 提供一键重启 |
