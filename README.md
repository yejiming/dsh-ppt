# DSH PPT · 一句话生成专业可编辑 PPT

**中文** | [English](README.en.md)

<p align="center">
  <img src="assets/banner.jpeg" alt="DSH PPT Banner" width="100%">
</p>

<p align="center">
  <img src="https://img.shields.io/github/v/release/yejiming/dsh-ppt?style=flat-square" alt="Version">
  &nbsp;
  <img src="https://img.shields.io/github/stars/yejiming/dsh-ppt?style=flat-square" alt="Stars">
  &nbsp;
  <img src="https://img.shields.io/npm/v/@yejiming%2Fdsh-ppt?style=flat-square&label=npm" alt="npm">
  &nbsp;
  <img src="https://img.shields.io/badge/license-MIT-blue?style=flat-square" alt="License">
</p>
<p align="center">
  <strong>专为职场人打造的 DeepSeek 智能 PPT 创作助手</strong><br>
  <em>一句话生成 · 原生可编辑 PPTX · 智能排版与配色 · 免版权配图 · 自动数据图表 · AI 视觉审稿质检</em>
</p>

<p align="center">

[产品亮点](#产品亮点) · [快速上手](#快速上手) · [依赖与环境要求](#依赖与环境要求) · [内置主题库](#内置主题库) · [放映动效与字体](#放映动效与字体) · [使用场景](#使用场景) · [产物说明](#产物说明) · [多 Agent 协作流程](#多-agent-协作流程) · [常见问题](#常见问题) · [开源协议](#开源协议)

</p>

## 产品亮点

做 PPT 总是耗费大量时间找模板、调对齐、排版面？市面上的 AI PPT 要么生成无法修改的“整页死图”，要么模板僵硬千篇一律。

**DSH PPT 让 PPT 制作回归内容本身：**

- 📝 **真正原生可编辑**：生成的不是整页死图，而是标准的 `.pptx` 文件。文字、图形、图表均可在 PowerPoint、Keynote 或 WPS 中随意二次修改与微调。
- 🎨 **专业级版式与审美**：告别套路化模板，AI 根据主题内容智能定制色彩搭配、字体层级与页面构图，呈现高级商务质感。
- 🎭 **12 套内置主题库**：每套主题都是一份经过校验的设计配方——配色取自 IBM Carbon、Tailwind、Radix、Microsoft Fluent 2 等公开色板，字体只使用本机真能渲染的家族，并附带动效节奏、装饰手法与逐页版式配方。调用 `ppt_themes` 直接选用，配色对比度全部经 WCAG AA（≥4.5:1）实算验证。
- 📈 **智能图表与精选配图**：自动搜索合规免版权商业配图，并根据数据自动绘制高品质图表，告别枯燥的大段纯文本。
- 👁️ **AI 视觉闭环审稿**：AI 会把做好的 PPT 真正“渲染并审阅一遍”，自动发现并纠正文字溢出、重叠与排版瑕疵，确保出片品质。
- ✨ **放映文字动画**：文字与图形可以按页配置入场动画——淡入、飞入、缩放、旋转、擦除等 29 种效果，支持按段落逐条构建、点击步编排与整页转场。
- 🔤 **任意本机字体**：不再局限于内置字体表，插件会枚举本机全部已安装字体（含 PANOSE、字距族、字符集与嵌入许可），也能把磁盘上的字体文件一键装进当前用户字体目录。
- 🔒 **本地安全与隐私保护**：所有制作与渲染过程在本地安全运行，素材来源与授权清晰记录，商业汇报更放心。

## 快速上手

### 1. 运行环境准备

- **DeepSeek Harness 0.2.0-rc.2 或更高版本**（本插件的 DSH 依赖声明为 `^0.2.0-rc.2`，与运行时版本一致；版本不匹配时插件会被拒绝加载）
- 电脑已安装以下任意办公软件（用于高保真渲染与视觉质检）：
  - **macOS**：Keynote 或 Microsoft PowerPoint
  - **Windows**：Microsoft PowerPoint
  - **通用 / Linux**：LibreOffice

### 2. 快速安装

在终端执行以下命令直接从 npm 安装插件：

```bash
# 通过 DSH 插件管理器直接安装
dsh plugin --profile web add @yejiming/dsh-ppt
dsh plugin --profile headless add @yejiming/dsh-ppt
```

### 3. 开始创作

#### 方式一：Web 界面（推荐）
启动 Web 控制台后，新建会话并选择 **「PPT 模式」**，像和设计师沟通一样直接向 AI 提出需求：
```bash
dsh --profile web
```

#### 方式二：命令行极速生成
也可以直接在命令行输入一句话指令，快速生成整套演示文稿：
```bash
dsh --profile headless "为管理层制作一份6页的AI项目季度复盘与规划汇报"
```

## 依赖与环境要求

插件在 DSH 宿主进程里运行：生成、回渲与质检全部落在本机工作目录内，离开本机的只有图片检索和 `browser_visit`。下面每一项都对应源码里的真实行为，不是推荐配置。

### 必需

| 项目 | 要求 | 说明 |
| :--- | :--- | :--- |
| DeepSeek Harness | `0.2.0-rc.2` 或更高 | 插件声明 19 项 `@deepseek-ai/*` 对等依赖（其中 18 项为 `^0.2.0-rc.2`，`@deepseek-ai/cordis` 为 `^4.0.1`）。版本不匹配时插件被拒绝加载，PPT 模式与 21 项工具都不会挂载 |
| Node.js | `>= 20.19.0` | `package.json` 的 `engines.node` 声明 |
| 操作系统 | Windows / macOS / Linux | `package.json` 的 `os` 只声明 `win32`、`darwin`、`linux`；其它平台在插件加载期抛 `PPT_PLATFORM_UNSUPPORTED` |
| Chromium 系浏览器 | 本机已安装 Chrome、Chromium 或 Edge | HTML 校验渲染、`browser_visit` 与 HTML 预览截图都靠它。不需要 `playwright install`：插件只探测本机可执行文件，再用 `chromium.launch({ executablePath })` 显式启动 |
| `sharp` | 必需，但声明为可选（见下一节） | 图片元数据校验、回渲图归一化与联系表拼版 |

运行时 npm 依赖：

| 依赖 | 加载方式 | 用途 |
| :--- | :--- | :--- |
| `pptxgenjs` | 静态导入 | 把校验后的页面结构写成原生 PPTX |
| `fflate` | 静态导入 | 读写 PPTX/OOXML 的 ZIP 容器，含转场与文字动画的 XML 改写 |
| `fontkit` | 静态导入 | 解析字体文件、判定字形覆盖、生成字体子集 |
| `zod` | 静态导入 | 校验大纲、Art Direction 与质检载荷 |
| `schemastery` | 静态导入 | 声明插件 Config schema |
| `playwright-core` | 静态导入 | 驱动本机 Chromium 系浏览器 |
| `happy-dom` | 动态 `import()` | 只在 `html_create` 校验 HTML 时按需加载，不参与宿主启动 |

除 `happy-dom` 外，以上依赖全部在模块顶层静态导入。它们是**加载期硬依赖**：任何一项缺失都会让插件整体加载失败，而不是某个工具单独降级。

### 可选（增强能力）

**`sharp`：声明为可选，实际必需。** `package.json` 把它放进 `peerDependencies`，并在 `peerDependenciesMeta` 里标记 `optional: true`；但 `src/image-search.ts`、`src/python.ts`、`src/ppt-image.ts`、`src/quality.ts` 都在顶层静态 `import sharp from 'sharp'`，且没有 try/catch 兜底。缺失时插件同样整体加载失败，与这份可选声明矛盾。**请按必需依赖提供 `sharp`。**

**PPT 渲染器。** 缺了不影响出 PPTX，但会失去全部逐页回渲证据与视觉质检，`ppt_image` 返回 `not_available`。候选程序与回退顺序见下一节。

**Python 分析运行时。** 只有 `python_execute` 需要，用于绘制数据图表与图像处理。可执行文件默认为 Windows `python`、macOS / Linux `python3`，可用配置项 `pythonExecutable` 覆盖。必须先装好 `matplotlib`、`Pillow`、`opencv-python`：每次执行前会跑一次 `import matplotlib, PIL, cv2` 预检，失败抛 `PYTHON_DEPENDENCY_MISSING`。执行环境固定 `MPLBACKEND=Agg`，不弹窗。

### 外部程序与回退顺序

| 程序 | 用途 | Windows | macOS | Linux | 缺失时的行为 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| Chrome / Chromium / Edge | HTML 校验渲染、`browser_visit`、HTML 预览截图 | `%ProgramFiles%`、`%ProgramFiles(x86)%`、`%LOCALAPPDATA%` 下的 `Google\Chrome\Application\chrome.exe`、`Chromium\Application\chrome.exe`、`Microsoft\Edge\Application\msedge.exe` | `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`、`/Applications/Chromium.app/Contents/MacOS/Chromium`、`~/Applications/Google Chrome.app/...` | `/usr/bin/google-chrome`、`/usr/bin/google-chrome-stable`、`/usr/bin/chromium`、`/usr/bin/chromium-browser`、`/snap/bin/chromium` | 抛 `BROWSER_NOT_READY`，诊断项 `browser` 为 `failed` |
| PowerShell | Windows 上驱动 PowerPoint COM 完成回渲 | `powershell.exe`、`powershell`、`pwsh.exe`、`pwsh` | — | — | Windows 的 `powerpoint` 后端不可用，回落 LibreOffice |
| Microsoft PowerPoint（桌面版） | 把 PPTX 回渲为逐页 PNG | 经 COM 启动；探测阶段不要求 `POWERPNT.EXE` 路径存在 | 经 AppleScript 启动，输出走 `screencapture` 截屏 | — | 该后端不可用 |
| Keynote | macOS 首选回渲：直接导出每页 PNG | — | `/Applications/Keynote.app`、`~/Applications/Keynote.app` | — | macOS 回落 LibreOffice |
| `osascript` | macOS 上驱动 Keynote / PowerPoint | — | `/usr/bin/osascript` | — | macOS 的 `keynote` 与 `powerpoint` 后端都不可用 |
| `screencapture` | macOS 的 PowerPoint 兜底：截屏放映窗口 | — | `/usr/sbin/screencapture` | — | macOS 的 `powerpoint` 后端不可用 |
| LibreOffice（`soffice`） | 跨平台回渲第一步：PPTX → PDF | `soffice.exe`、`soffice`，以及 `%ProgramFiles%`、`%ProgramFiles(x86)%`、`%LOCALAPPDATA%` 下的 `LibreOffice\program\soffice.exe` | `soffice`、`/Applications/LibreOffice.app/Contents/MacOS/soffice` | `soffice`、`libreoffice`、`/usr/bin/soffice`、`/usr/bin/libreoffice` | LibreOffice 后端不可用 |
| Poppler（`pdftoppm`） | 回渲第二步：PDF → PNG | `pdftoppm.exe`、`pdftoppm` | `pdftoppm` | `pdftoppm` | **整个 LibreOffice 后端不可用**：`soffice` 与 `pdftoppm` 必须同时存在 |
| Python | 数据图表与图像处理 | `python` | `python3` | `python3` | `python_execute` 抛 `PYTHON_DEPENDENCY_MISSING` |

浏览器探测顺序：配置项 `browserExecutable` → `playwright-core` 记录的浏览器路径 → 上表中的系统候选。

`backend=auto` 时的回渲顺序：

- **macOS**：`keynote` → `libreoffice` → `powerpoint`（截屏兜底，会额外提示动画与多显示器配置可能影响截帧）
- **Windows**：`powerpoint`（PowerPoint COM，由 PowerShell 驱动）→ `libreoffice`
- **Linux**：只有 `libreoffice`

所有候选都不可用时**不抛错**：`ppt_image` 返回 `not_available`，质检报告写入 `render_status: not_available` 并追加一条 `RENDERER_NOT_AVAILABLE` 警告，自动视觉层同步为 `not_available`。也就是说没有渲染器仍能产出 PPTX，但拿不到任何逐页回渲证据。

### 网络与隐私

只有两个出口：

| 操作 | 目标 | 凭据 | 失败时的行为 |
| :--- | :--- | :--- | :--- |
| `image_search` | 先请求 `api.openverse.org/v1/images/`，结果不足时补 `commons.wikimedia.org/w/api.php` | 不需要 API key | 不抛错：返回已拿到的结果并附 `degraded:unavailable` 之类的告警，同时给出三条替代路径（`read_image` 复用已冻结素材、`browser_visit` 从公开页面抓取、注入自定义 provider）。只有显式打开 `strictOnUnavailable` 才抛 `IMAGE_SEARCH_FAILED` |
| `browser_visit` | 模型指定的任意公网 http(s) 页面 | 不需要凭据 | 抛 `BROWSER_URL_BLOCKED`（URL 被判定为不可访问）或 `BROWSER_NOT_READY`（没有可用浏览器） |

`image_search` 的请求约束：15 秒超时、`redirect: 'error'`（不跟随重定向）、单响应上限 2 MiB、结果按查询缓存 10 分钟（降级结果不缓存，下次会真正重试）。

`browser_visit` 与 `image_search` 的候选图片 URL 共用同一套 SSRF 防护：只允许 `http` / `https`；拒绝带用户名或密码的 URL；只允许 80 / 443 端口；拒绝 `localhost`、`*.localhost`、`*.local`、`*.internal`；域名解析后逐个校验地址，环回、私有段、链路本地（含云元数据地址）、CGNAT、组播与保留段一律拒绝。浏览器上下文里的每个请求都经过同一条拦截路由，命中即中止。

HTML 侧没有外链：`html_create` 拒绝非本地资源引用和含远程 / data / file 资源的 CSS；PPT 模式里 `dsh-tool-web` 的 `fetch` 已关闭，该工具只做搜索。

其余能力——PPTX 生成、字体解析与子集化、回渲、自动质检、联系表拼版——全部在本机工作目录内完成。

### 平台差异矩阵

| 能力 | Windows | macOS | Linux | 不满足时的行为 |
| :--- | :--- | :--- | :--- | :--- |
| HTML 渲染与浏览器访问 | Chrome / Chromium / Edge 任一 | 同左 | 同左 | 抛 `BROWSER_NOT_READY` |
| PPTX 回渲（`auto`） | PowerPoint COM → LibreOffice | Keynote → LibreOffice → PowerPoint 截屏 | 仅 LibreOffice | 不抛错；`render_status=not_available` 加 `RENDERER_NOT_AVAILABLE` |
| LibreOffice 后端前置条件 | `soffice` 与 `pdftoppm` 同时存在 | 同左 | 同左 | 该后端不可用；Linux 上等于回渲能力归零 |
| Python 图表与图像处理 | `python` 加 matplotlib / Pillow / opencv-python | `python3` 加同上 | `python3` 加同上 | 抛 `PYTHON_DEPENDENCY_MISSING` |
| 系统字体目录 | `%SystemRoot%\Fonts`、`%LOCALAPPDATA%\Microsoft\Windows\Fonts` | `/System/Library/Fonts`、`/System/Library/Fonts/Supplemental`、`/System/Library/AssetsV2/com_apple_MobileAsset_Font7`、`...Font8`、`/Library/Fonts`、`~/Library/Fonts` | `/usr/share/fonts`、`/usr/local/share/fonts`、`~/.local/share/fonts`、`~/.fonts` | 配置项 `fontDirs` 只在这批目录之外**追加**；HTML 声明了未安装的批准主字体时抛 `PPT_DEPENDENCY_MISSING` |
| 字体安装（自定义字体） | 复制到 `%LOCALAPPDATA%\Microsoft\Windows\Fonts`，再用 `reg.exe` 写 HKCU 注册值 | 复制到 `~/Library/Fonts` | 复制到 `~/.local/share/fonts`，再尽力执行 `fc-cache -f` | 只影响自定义字体的可用性，不影响出片；Windows 上缺 `reg.exe` 时该步骤报错 |

### 对照：安装前自检

**Windows（PowerShell）**

```powershell
node -v
dsh --version

# Chromium 系浏览器：任一存在即可
Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
Test-Path "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
Test-Path "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"

# 回渲后端：PowerShell 加 PowerPoint，或 LibreOffice 加 Poppler
Get-Command powershell.exe
Test-Path "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE"
Get-Command soffice.exe, pdftoppm.exe -ErrorAction SilentlyContinue

# Python 分析依赖
python -c "import matplotlib, PIL, cv2; print('python deps ok')"

# sharp（在插件被安装到的 profile 目录里执行）
node --input-type=module -e "import('sharp').then(() => console.log('sharp ok')).catch(() => { console.error('sharp missing'); process.exit(1) })"
```

**macOS**

```bash
node -v
dsh --version

# 浏览器：任一存在即可
ls -d "/Applications/Google Chrome.app" /Applications/Chromium.app 2>/dev/null

# 回渲后端：Keynote / PowerPoint，或 LibreOffice 加 Poppler
ls -d /Applications/Keynote.app "/Applications/Microsoft PowerPoint.app" 2>/dev/null
which osascript screencapture pdftoppm soffice

# Python 分析依赖
python3 -c "import matplotlib, PIL, cv2; print('python deps ok')"

# sharp
node --input-type=module -e "import('sharp').then(() => console.log('sharp ok')).catch(() => { console.error('sharp missing'); process.exit(1) })"
```

**Linux**

```bash
node -v
dsh --version

# 浏览器：任一存在即可
which google-chrome google-chrome-stable chromium chromium-browser

# 回渲后端：soffice 与 pdftoppm 必须同时存在
which soffice libreoffice pdftoppm

# Python 分析依赖
python3 -c "import matplotlib, PIL, cv2; print('python deps ok')"

# sharp
node --input-type=module -e "import('sharp').then(() => console.log('sharp ok')).catch(() => { console.error('sharp missing'); process.exit(1) })"
```

以上全绿后，PPT 模式的运行诊断（`platform`、`tools`、`browser`、`python`、`fonts`、`renderer`、`attachments`、`vision_model` 共八项）应当全部为 `ready`，其中 `vision_model` 还要求当前模型声明接受图片输入；任何一项 `failed` 都会在开工前被指出来。

## 内置主题库

插件内置 12 套设计主题。主题不是套壳模板，而是**一份可校验的设计配方**：配色、四个字体角色、逐页版式倾向、节奏限制，以及这套主题允许的装饰手法。它给 Art Direction 一个经过验证的起点，而不是替 AI 做设计——`concept`、`audience_effect`、每页的 `job` 与 `takeaway` 仍然由当前 Agent 按实际内容写。

| 主题 id | 名称 | 适用场景 | 底色（主 / 反色） | 强调色 |
| :--- | :--- | :--- | :--- | :--- |
| `carbon-blueprint` | 碳素蓝图 | 技术方案、架构评审、RFC 评审 | `#FFFFFF` / `#161616` | `#0043CE` |
| `slate-review` | 板岩复盘 | 季度复盘、经营分析、OKR 回顾 | `#F8FAFC` / `#0F172A` | `#1D4ED8` |
| `midnight-raise` | 午夜融资 | 融资路演、战略汇报、董事会 | `#0B1220` / `#FFFFFF` | `#78A9FF` |
| `paper-ink` | 纸墨 | 学术报告、白皮书、深度长文 | `#F7F5F0` / `#111318` | `#9F1853` |
| `editorial-serif` | 学刊 | 学术答辩、研究报告、政策解读 | `#F7F3F2` / `#171414` | `#8A3800` |
| `telemetry-teal` | 遥测青 | 运维复盘、数据看板汇报、性能评审 | `#FFFFFF` / `#081A1C` | `#005D5D` |
| `graphite-minimal` | 石墨极简 | 产品规格、内部对齐、工程文档 | `#FCFCFC` / `#202020` | `#3A5BC7` |
| `indigo-launch` | 靛蓝发布 | 产品发布、新版本宣讲、大会 Keynote | `#1E1B4B` / `#EEF2FF` | `#A5B4FC` |
| `amber-academy` | 琥珀课堂 | 培训、工作坊、新人上手 | `#FFFBEB` / `#451A03` | `#92400E` |
| `crimson-brief` | 绯红简报 | 营销提案、立项申请、一页纸决策 | `#FFFFFF` / `#4C0519` | `#BE123C` |
| `fluent-azure` | 流蓝企业 | 企业汇报、客户方案、招标应答 | `#FFFFFF` / `#0A2E4A` | `#0F6CBD` |
| `moss-annual` | 苔绿年报 | 年度报告、ESG 披露、长期规划 | `#F0FDF4` / `#14532D` | `#15803D` |

配色不是拍脑袋定的：每套主题的色值都取自公开设计系统（IBM Carbon、Tailwind、Radix、Microsoft Fluent 2、Open Color），并在代码里以 `palette_source` 记录了出处。每套主题还带一个**反色强调色**——因为同一个强调色不可能同时在一个浅底和一个深底上都达到 4.5:1，这是颜色本身的物理限制，所以主题成对给出，逐页方案会指明该页用哪一个。

### 怎么用

```text
ppt_themes()                                              # 列出全部主题与场景标签
ppt_themes(scene="融资路演")                                # 按场景筛选
ppt_themes(theme_id="carbon-blueprint")                    # 读取完整配方
ppt_themes(theme_id="carbon-blueprint", page_types=[...])   # 取逐页视觉方案
ppt_outline(..., theme_id="carbon-blueprint", art_direction={...})
```

`ppt_outline` 带上 `theme_id` 后会做一致性校验，任何偏离都以告警列出，而不是静默通过：

- `THEME_PALETTE_DRIFT`：出现了主题未定义的颜色
- `THEME_ACCENT_REPLACED`：换掉了主题强调色
- `THEME_TYPOGRAPHY_REPLACED`：换掉了主题字体角色
- `THEME_ACCENT_DRIFT` / `THEME_FONT_DRIFT`：某页 `style.accent` 或 `style.font` 不在主题内
- `THEME_PLAN_MISSING`：选了主题却没给 `art_direction`，主题无从生效

### 配色来源与品牌色

12 套主题覆盖常见商务场景；要走客户品牌规范时，用 `browser_visit` 读取公开的品牌规范页取色，再按同一套纪律自建主题：正文与强调色对每一层底色都要 ≥4.5:1，深色页准备一个独立的反色强调色。`ppt_themes` 的返回值里带着这套规则，可以直接照着做。

### 设计约束（主题配方遵守，HTML 也必须遵守）

- 画布 1280×720，页边距水平 72px / 垂直 64px，内容区 1136×592，12 栏、列间距 24px，圆角只用 4 / 8 / 12px
- 字号层级：封面主标题 72px、页标题 44px、小节标题 32px、导语 26px、正文 24px、辅助 16px、指标数字 88px、代码 17px
- 装饰是封闭集合：1px / 2px 边框、圆角、色阶差、opacity、z-index、字号对比、栅格对齐、纯色或 `linear-gradient` 背景
- 投影（`box-shadow`）、`transform`、`filter`、`clip-path`、`content` 一律禁止——立体感只能靠边框与色阶，这样 HTML 预览与 PPTX 产物才不会出现保真断层

## 放映动效与字体

### 文字动画

`ppt_create` 接受可选的 `effects` 参数，为每一页声明放映效果。转场与入场动画都在 pptxgenjs 写出包之后、原子提交之前注入，只改写 `ppt/slides/slideN.xml`，其余部件逐字节保持原样；某一页没有声明动效就完全不动。

```jsonc
{
  "page": 1,
  "transition": { "type": "fade", "speed": "slow" },
  "animations": [
    { "target": "title-1", "effect": "fly-in", "direction": "up", "duration_ms": 600 },
    { "target": "body-1", "effect": "fade", "by_paragraph": true, "start": "after-previous" }
  ]
}
```

- `target` 是 IR 里的元素 id，也就是 PPTX 中该形状的名字。
- 入场效果共 29 种：`appear`、`flash-once`、`fade`、`dissolve`、`wedge`、`wipe`、`blinds`、`checkerboard`、`random-bars`、`box`、`circle`、`diamond`、`plus`、`split`、`strips`、`wheel`、`zoom`、`fly-in`、`crawl`、`peek`、`stretch`、`swivel`、`spiral`、`bounce`、`credits`、`float-in`、`grow-turn`、`rise-up`、`unfold`。
- 其中 `zoom`、`fly-in`、`crawl`、`peek`、`stretch`、`swivel`、`spiral`、`bounce`、`credits`、`float-in`、`grow-turn`、`rise-up`、`unfold` 属于**位移与缩放类**效果，文字会真正移动，而不只是显隐切换。
- `direction` 仅对渲染方向的效果有效：`wipe`、`fly-in`、`crawl`、`peek`、`blinds`、`checkerboard`、`random-bars`、`box`、`circle`、`diamond`、`plus`、`stretch`、`swivel`。
- `start` 取 `on-click`（默认）、`with-previous` 或 `after-previous`；一页的第一个动画总是开启一个点击步。
- 每页最多 24 条动画，同一元素在同一构建单位下只能出现一次；`by_paragraph` 会让它按段落逐条播放。

内部实现不靠规范散文，而以 PowerPoint 实际写出的时间线为准：`presetID` 的编号空间按 `presetClass` 划分，`presetSubtype` 是方向位掩码（`1`=上、`2`=右、`4`=下、`8`=左、`16`=入场、`32`=出场），真正决定渲染的是 `p:animEffect` 的 filter 与 `p:anim` 的初末值。

### 字体：自由选择与安装

`ppt_fonts` 有两个取样范围：

- `scope=registry`（默认）：插件内置的推荐字体表，并按角色给出确定性推荐。
- `scope=installed`：**本机实际安装的全部字体面**，每一项都带 PANOSE、pitch/family 字节、charset、字重、字形数、拉丁与中日韩覆盖，以及字体自身的嵌入许可位（`fs_type` / `embeddable`）。

`ppt_outline` 的 `title_font` / `body_font` 因此不再受内置表限制：**任何本机已安装的字体家族都可以直接指定**。只有当某个家族不在内置表里时，插件才会去扫描系统字体目录（要读取每个字体文件），因此常见路径没有额外开销。

把磁盘上的字体文件装进当前用户字体目录，之后就能在稿件里直接引用：

```jsonc
{ "scope": "installed", "install_path": "assets/fonts/BrandSans.ttf", "dry_run": true }
```

- 安装是**用户级**的，不需要管理员权限：Windows 写入 `%LOCALAPPDATA%\Microsoft\Windows\Fonts` 并登记 `HKCU` 下的字体项，macOS 写入 `~/Library/Fonts`，Linux 写入 `~/.local/share/fonts` 并尽力刷新 `fc-cache`。
- 先跑 `dry_run: true`，只会计算目标路径与家族名，不落盘、不写注册表。
- 返回值带 `uninstall_hint`，便于手工卸载。

关于把字体**嵌进 PPTX 文件本身**：PowerPoint 的原生嵌入部件（`ppt/fonts/*.fntdata`）是 EOT 容器，内部为 MicroType Express 压缩的 TrueType 数据，本插件不生成该格式。需要跨机器保真时，请用 `ppt_fonts` 把字体装到目标机器，或随产物分发字体文件后在目标机器安装。

## 使用场景

| 场景分类 | 提示词示例 |
| :--- | :--- |
| 📊 **业务复盘 / 述职报告** | *"制作一份 6 页的 Q3 电商运营复盘 PPT，重点突出 GMV 增长、转化漏斗分析及下季度策略"* |
| 💼 **商业计划 / 方案提案** | *"为客户撰写一份关于企业私有化 AI 知识库的解决方案 PPT，8 页，商务科技风"* |
| 📢 **产品发布 / 功能推介** | *"围绕新上线的移动端应用设计一份产品发布宣讲 PPT，突出核心卖点与用户体验升级"* |
| 🎓 **培训分享 / 团队规范** | *"制作一份 5 页的敏捷开发入门与团队协作规范分享 PPT，风格轻松明快、条理清晰"* |

## 产物说明

每次任务完成后，系统会在 `ppt-output/` 目录下生成独立的项目文件夹：

- 📄 **`deck.pptx`**：最终生成的原生 PPT 文件，直接双击用 PowerPoint / Keynote / WPS 打开即可放映或自由编辑。
- 🖼️ **`preview/`**：各页面的高清截图预览，方便在手机、群聊或文档中快速预览与确认效果。
- 📁 **`assets/`**：PPT 中使用的图片素材与高清图表文件，并附带合规版权来源清单。

## 多 Agent 协作流程

一份 PPT 也可以由多个 agent 分工完成：一个搜图、一个找材料、一个出框架（大纲 + Art Direction）、主 agent 整合并交付。插件只提供固定的 21 项工具与一套强校验的工件契约，分工由会话层编排。

- 🔍 **搜图 agent**：只往候选池写图片候选，不碰工件目录。搜索与落盘是两件事，冻结素材是主 agent 的活。
- 📚 **材料 agent**：只写带来源的研究笔记与结构化数据点。
- 🧭 **框架 agent**：只产出大纲与 Art Direction 载荷，不生成任何成品文件。
- 🎯 **主 agent**：独占 `ppt_outline` / `html_create` / `ppt_create` / `ppt_image`，按 `outline → html → pptx → 回渲 → 逐页审稿 → finalize` 的质量门顺序交付。

角色输入输出、写入边界（谁写 `outline.json`、谁写 `deck.html`、谁只产出候选清单）、依赖顺序与失败降级路径（含「图源不可用 → 退回零外链纯矢量设计」的真实案例）见 **[多 Agent 协作流程](docs/agent-workflow.md)**。

## 常见问题

<details>
<summary><b>Q: 生成的 PPT 在我的 Office / WPS 里打开会排版错乱吗？</b></summary>
不会。系统内置了多平台字体兼容适配方案（优先采用微软雅黑、苹方、思源黑体等主流系统字体），确保在不同设备与办公软件中打开时文字排版依然稳定美观。
</details>

<details>
<summary><b>Q: 可以把生成的 PPT 导出为 PDF 或直接全屏演示吗？</b></summary>
当然可以！生成的 <code>.pptx</code> 为标准演示文稿格式，支持在 PowerPoint、Keynote 或 WPS 中直接进入演讲者模式放映，或一键导出为 PDF / 演讲备注。
</details>

<details>
<summary><b>Q: 如果对生成的部分页面不满意，该如何调整？</b></summary>
有两种便捷方式：
1. **直接对话修改**：在会话中告诉 AI 调整意见（如“第 3 页换成柱状图对比”、“整体色系调整为商务蓝”）；
2. **本地自由编辑**：直接打开 <code>deck.pptx</code>，像普通 PPT 一样随意修改文字、替换图片或移动排版位置。
</details>

## 开源协议

本项目基于 [MIT License](LICENSE) 开源。
