# 多 Agent 协作流程

把「一份 PPT 由多个 agent 分工完成」变成可照做的流程。

分工形态：一个 agent 搜图、一个 agent 找材料、一个 agent 出框架（大纲 + Art Direction）、主 agent 整合并把 PPTX 交付出去。

本文只描述插件里真实存在的工具面。文档里的每个工具名都来自 `src/schemas.ts` 的 `PPT_MODE_TOOL_NAMES`，逐字核对过。

---

## 0. 先看硬边界：PPT 模式的工具面是固定的 21 项

| 工具 | 用途 |
| :--- | :--- |
| `read` | 读文本文件 |
| `write` | 新建或整份替换文本文件 |
| `edit` | 定点改文本文件 |
| `read_image` | 把 PNG/JPEG/WebP/GIF 读进上下文（审稿靠它） |
| `todo_write` | 列任务清单，显示进度 |
| `ask_user_question` | 用途、受众、页数、语言、品牌等澄清 |
| `web_search` | 联网检索 |
| `pwsh` / `bash` | 本机 shell；**Windows 上是 `pwsh`，macOS 与 Linux 上是 `bash`** |
| `python` | 有界非交互 Python（数据分析、Matplotlib Agg 图表、Pillow/OpenCV 图像处理） |
| `browser_visit` | 打开公开 HTTP(S) 页面或插件生成的本地 HTML 预览（只读研究页） |
| `browser_find` | 在只读研究页里找可见文本/交互元素，返回带版本号的引用 |
| `browser_click` | 点 `browser_find` 返回的引用 |
| `browser_scroll_down` / `browser_scroll_up` | 在只读研究页里滚动 |
| `image_search` | 免费匿名的 Openverse 检索，自动回退 Wikimedia Commons；不需要 API key |
| `ppt_fonts` | 只读查询**插件批准注册表内当前可用**的字体（不是本机全量字体清单） |
| `ppt_themes` | 只读查询内置主题库：目录、单套完整配方、以及按页型展开的逐页视觉方案（构图、疏密、底色角色、标题处理、帧策略与具体色值） |
| `ppt_outline` | 校验大纲与 Art Direction，原子生成 `outline.json` 与 `design-plan.json`；给了 `theme_id` 时同时校验主题一致性 |
| `html_create` | 校验受限静态 HTML，原子写 `deck.html`，并逐页渲染 PNG 预览 |
| `ppt_create` | 把 HTML 设计稿转成可编辑 PPTX；也负责 `finalize_visual_review=true` |
| `ppt_image` | 用真实渲染器打开 PPTX，输出逐页 PNG 与联系表，并可刷新机器质量报告 |

两条必须记住的限制：

1. **PPT 模式本身不含「派生 teammate」或跨 agent 通信的工具。** 上表 21 项里没有派生、消息投递、共享任务板这一类工具。分工是**会话层编排**——由具备 Agent Teams 的会话派生子 agent，或人工并行开多个会话。插件不负责调度，只负责给出这 21 项工具和一套强校验的工件契约。角色之间的隔离靠任务书里的写入范围约定，不靠工具面隔离。
2. 无论角色怎么划，**能用的工具都不会超出上表**。表外工具调用不会成功。

---

## 1. 工件目录与暂存区

`ppt_outline` 是**唯一**会分配工件目录的东西，返回的路径形如：

```
ppt-output/<title-slug>/          ← <root>，下称工件目录
  outline.json
  design-plan.json
  deck.html
  deck.pptx
  assets/images/                  ← 冻结后的图片素材
  assets/source-manifest.json     ← 素材来源记账
  preview/                        ← HTML 预览、pptx 逐页图、联系表、高风险页副本
  report.json                     ← 机器质量报告（四层）
  visual-review.json              ← 人工/模型审稿结论
```

`<title-slug>` 由 `artifact_title` 归一化而来；同名目录已存在时自动加 `-2`、`-3`。

框架定稿**之前**没有工件目录，所以三个子角色一律在暂存区工作：

```
ppt-output/_staging/<topic>/
  research/    ← 材料 agent 独占
  images/      ← 搜图 agent 独占
  frame/       ← 框架 agent 独占
```

三个子目录互不重叠，是并行的前提。

---

## 2. 角色表

### 2.1 材料 agent

| 项 | 内容 |
| :--- | :--- |
| 输入 | 用户需求原文；主 agent 转达的澄清结论（用途、受众、页数、语言、品牌） |
| 输出 | `ppt-output/_staging/<topic>/research/sources.md`：结论、事实、每条事实的来源链接；可选 `research/facts.json`：结构化数据点，每条带来源 |
| 工具 | `web_search`、`browser_visit`、`browser_find`、`browser_scroll_down`、`browser_scroll_up`、`browser_click`、`read`、`write`、`pwsh`/`bash`、`python` |
| 写入边界 | 只写 `research/` 子目录 |
| 禁止 | 不写工件目录下任何文件；不为凑页数编造数据；核不动的数据点标 `unverified`，不要假装确定 |

### 2.2 搜图 agent

| 项 | 内容 |
| :--- | :--- |
| 输入 | 框架 agent 或主 agent 给的配图需求（逐页的图片意图与检索词）；材料 agent 的来源清单 |
| 输出 | 候选池 `ppt-output/_staging/<topic>/images/candidates.json`：每条候选记录来源页、图片 URL、提供方、许可名称、许可 URL、作者、检索词；可选把缩略图下到 `images/raw/` 供 `read_image` 目检 |
| 工具 | `image_search`、`browser_visit`、`read`、`read_image`、`write`、`pwsh`/`bash`、`python` |
| 写入边界 | 只写 `images/` 子目录 |
| 禁止 | 不写 `<root>/assets/images/`；不写 `assets/source-manifest.json`；框架定稿前不把任何候选写进大纲载荷 |

关于 `image_search` 的能力边界，说清楚：

- 它返回的是**候选元数据**（图片来源页、图片 URL、提供方、许可名称等），**不把图片落到工件目录**，也不做许可核验——候选里带的许可信息按其字面记录即可，不要当成已核验结论。
- 因此「搜索」和「冻结」是两件事。搜索属于搜图 agent，**冻结属于主 agent**（见 4.3）。
- 它只有两个免费来源：Openverse，失败时回退 Wikimedia Commons。两个都拿不到结果就是「图源不可用」，走 5.1 的降级路径。

### 2.3 框架 agent（大纲 + Art Direction）

| 项 | 内容 |
| :--- | :--- |
| 输入 | 用户需求；材料 agent 的 `sources.md` / `facts.json`；搜图 agent 的候选池（用来判断哪几页适合上图片锚点） |
| 输出 | `ppt-output/_staging/<topic>/frame/outline-payload.json`（slides 数组，每页恰好 `page`、`type`、`title`、`content`、`style`）与 `frame/art-direction.json`（deck 级 concept/audience_effect/palette/typography/rhythm + 逐页 Art Direction） |
| 工具 | `read`、`write`、`ppt_themes`、`ppt_fonts`、`web_search`（补背景）、`python`（配平篇幅与数据） |
| 写入边界 | 只写 `frame/` 子目录 |
| 禁止 | 不调用 `ppt_outline`、`html_create`、`ppt_create`、`ppt_image`；不写 `deck.html`；不把候选池里的具体文件当成 `asset` 引用 |

框架 agent 的稿子必须逐页写全这 8 个字段，缺一项会被 Art Direction 校验打回：

`job`、`takeaway`、`composition`、`density`、`background_role`、`title_treatment`、`visual_anchor`、`frame_policy`。

先用 `ppt_themes` 定主题再排页，比从零设计省事得多：传 `theme_id` 与 `page_types` 会直接返回逐页的 `composition` / `density` / `background_role` / `title_treatment` / `frame_policy` 与该页应使用的具体色值（含反色页要换用的 `accent_inverted`）。主题仍是起点：`concept`、`audience_effect` 与每页 `job`、`takeaway` 必须按实际内容自己写。需要品牌色时，用 `browser_visit` 读公开品牌规范取色，并按同一纪律自建（正文与强调色对每一层底色都 ≥4.5:1，深色页准备独立的反色强调色）。

另外三条会直接触发校验失败或告警的硬约束，框架 agent 交稿前自查：

1. `art_direction.slides` 的页数必须等于大纲页数，`page` 必须从 1 连续递增。
2. `rhythm.background_sequence` 必须逐页对上该页的 `background_role`（数量和顺序都要一致）。
3. `typography` 的四个角色（display / body / latin / code）的字体族必须来自 `ppt_fonts` 返回的批准注册表；给了表外字体会被确定性替换，替换结果以 `ppt_outline` 返回值为准。

还有两条会被质量层记为设计告警，别踩：`frame_policy = grouped` 的页数别超过 `rhythm.max_grouped_frame_slides`；相邻页不要连续重复同一种 `composition`，除非该页显式标了「有意重复」。

### 2.4 主 agent（整合与交付）

| 项 | 内容 |
| :--- | :--- |
| 输入 | 三方产物 + 澄清结论 |
| 输出 | `<root>/outline.json`、`design-plan.json`、`deck.html`、`deck.pptx`、`preview/**`、`report.json`、`visual-review.json` |
| 工具 | 全部 21 项；**但只有主 agent 调用 `ppt_outline`、`html_create`、`ppt_create`、`ppt_image`** |
| 写入边界 | 工件目录的唯一 owner |
| 禁止 | 手写或手改 `outline.json` / `design-plan.json` / `deck.html` / `report.json`；手工修补 OOXML；用 `rasterize-element` 兜底（除非用户显式授权） |

---

## 3. 写入边界总表

| 路径 | 唯一写者 | 写入方式 |
| :--- | :--- | :--- |
| `ppt-output/_staging/<topic>/research/**` | 材料 agent | `write` |
| `ppt-output/_staging/<topic>/images/**` | 搜图 agent | `write` / `pwsh` / `python` |
| `ppt-output/_staging/<topic>/frame/**` | 框架 agent | `write` |
| `<root>/outline.json`、`<root>/design-plan.json` | `ppt_outline` | 主 agent 调用，原子写；任何人不手写 |
| `<root>/deck.html`、`<root>/preview/*.png` | `html_create` | 主 agent 调用，原子写；目录内已有 `deck.html` 会直接拒绝 |
| `<root>/deck.pptx` | `ppt_create` | 主 agent 调用 |
| `<root>/preview/pptx/**`、`preview/contact-sheet-NNN.png`、`preview/high-risk-page-NNN.png` | `ppt_image` | 主 agent 调用 |
| `<root>/report.json` | 机器质量层 | `ppt_create` / `ppt_image(refresh_quality=true)` / finalize 时写 |
| `<root>/visual-review.json` | 主 agent | `write` 一次，之后 finalize 只读不写 |
| `<root>/assets/images/**`、`<root>/assets/source-manifest.json` | 主 agent | 下载落盘 + `write`，必须在 `html_create` 之前 |

一句话记法：**子角色只碰暂存区，工件目录里的每个文件都只有一个写者**。

---

## 4. 依赖顺序与质量门

```
Gate A  澄清（主 agent，ask_user_question）
   ↓
Gate B  并行：材料 agent 出 research/  ·  搜图 agent 出 images/candidates.json
   ↓
Gate C  框架 agent 交 frame/ 两份载荷  →  主 agent 调 ppt_outline  拿 <root> / outline_path / design_plan_path / design_status / fonts
   ↓
Gate D  冻结素材：候选 → <root>/assets/images/ + source-manifest.json
   ↓
Gate E  html_create(outline_path, design_plan_path, html)  → deck.html + preview/*.png
   ↓
Gate F  ppt_create(html_path, outline_path, output_path)   → deck.pptx
   ↓
Gate G  ppt_image(pptx_path, backend=auto, refresh_quality=true)  → 逐页图 + 联系表 + 重算机器四层
   ↓
Gate H  read_image 逐页审稿 → write visual-review.json → ppt_create(..., finalize_visual_review=true)
```

### 4.1 Gate C：`ppt_outline` 的返回值要逐项看

`ppt_outline` 成功后会返回：`artifact_dir`、`outline_path`、`design_plan_path`（带 Art Direction 时才有）、`design_status`（`directed` 或 `legacy`）、`page_count`、`type_counts`、`fonts`、`warnings`、`blocking_warnings`。

判断条件：

- `design_status = legacy`：Art Direction 没生效，退回了旧模式，设计保真层只会按 legacy 模式评估。先排查 art_direction 载荷是否真的传了、是否校验通过，再往下走。
- `warnings` 非空：逐条读完，重点是字体回退。**此后一切字体以返回的 `fonts` 为准**，HTML 里不要再写没被返回的字体族。
- `blocking_warnings` 非空：这份大纲有必须先解决的问题，别进入下一阶段。
- `page_count` 与 `type_counts`：页数与需求不符当场返工，不要拖到 HTML 或 PPTX 阶段才发现。

### 4.2 Gate C 之前的铁律：搜图产出只能进候选池

`ppt_outline` 没成功返回路径之前，`<root>` 不存在。此时：

- 候选池文件只能待在 `_staging/<topic>/images/`，可以被框架 agent 读来定配图意图，但**不得**被当成 `asset` 写进任何大纲载荷。
- 大纲里表达配图意图的正确做法是写 `query`（检索词），把具体文件留到 Gate D 再定。大纲的图片项是「`query` 或 `asset` 二选一」的强校验结构。
- 判断条件很直白：**框架 agent 交稿时，载荷里出现的任何本地文件路径都应该是错的**——那时候还没有目录能装它。

### 4.3 Gate D 为什么要单独一步

`html_create` 会校验 HTML 里的本地素材引用：`img[src]`、`svg image` 的 `href` / `xlink:href`、以及 CSS `url(...)`，指向的文件必须真实存在，且不能逃出工件目录。所以素材必须先落到 `<root>/assets/images/`，HTML 才能引用。

顺带把外链这件事说透：外链图片在 HTML 校验阶段不会被拦，但在 `ppt_create` 阶段会被拒——PPTX 内的图片必须是本地冻结文件，外部关系（`TargetMode="External"`）被明令禁止。**所以不要用外链凑数，那只是把失败推迟到更贵的阶段。**

冻结时同步写 `<root>/assets/source-manifest.json`，每条素材至少记全：原始图片 URL、来源页、抓取时间、作者、许可名称、许可 URL、本地相对路径、sha256。许可名称抓不到就写 `unknown`，不猜。

### 4.4 Gate H 的审稿硬规则

`visual-review.json` 不是自由文本，它有结构强校验。写成下面这样才可能被 finalize 接受：

| 字段 | 要求 |
| :--- | :--- |
| `version` | 必须 `1` |
| `status` | `passed` / `failed` / `not_available` 三选一（`not_performed` 会被拒） |
| `checklist` | 字符串数组，最多 30 条 |
| `reviewed_assets` | 字符串数组，最多 200 条，路径必须真实存在 |
| `findings` | `{page?, severity: warning\|error, message}` 数组，最多 500 条 |
| `completed_at` | 可选，带时区的 ISO 时间 |

`status = passed` 还额外要求：

1. `reviewed_assets` 非空，并且**覆盖 `report.json` 里 `artifacts.contact_sheets` 与 `artifacts.high_risk_previews` 的全部路径**（一个都不能漏）；
2. `findings` 里没有 `error`。

所以审稿顺序是：先 `read_image` 看完全部联系表，再看封面、目录、数据页、复杂对比页、结束页，再看 `artifacts.high_risk_previews`（机器结构层或自动视觉层报过告警的那几页，由 `ppt_image` 从逐页图里复制出来）。看得不完整就不要写 `passed`——写 `not_available` 比写假通过诚实，而且 finalize 接受它。

---

## 5. 降级路径

### 5.1 图源不可用 → 退回零外链的纯矢量设计

**信号**：Openverse 与 Wikimedia Commons 都没给出可用候选；候选的许可信息全部缺失到无法记账；下载失败或返回的不是图片格式。

**动作**：

1. 停掉所有图片相关计划。把配图意图从载荷里撤掉（image 项改成 `point` / `data`，或把该页的视觉锚点换成 `typography` / `diagram` / `data`），**回框架 agent 改 `frame/` 载荷后重新调 `ppt_outline`**——已经生成的 `outline.json` 不许手改。
2. HTML 里删干净所有 `img`、`svg image`、CSS `url(...)` 引用。整套设计的视觉重量交给 CSS 渐变、几何形状、内联 SVG 图形、大字排版与数据排版。
3. 照常走 Gate E 之后的流程。零外链稿子在 `html_create` 阶段没有本地素材依赖，在 `ppt_create` 阶段没有 `PPT_CREATE_ASSET_MISSING` 风险。

这不是纸面预案：**一次实际交付中图源整体不可用，最终整套页面改成零外链的矢量设计后按常规质量门交付。** 关键判断是——宁可用矢量化设计保交付，也不要拿外链或来源不明的图凑数。

### 5.2 其余故障与降级动作

| 故障 | 判断信号 | 降级动作 |
| :--- | :--- | :--- |
| 回渲后端不可用 | `ppt_image` 结果里 render 层为 `not_available`，且各后端尝试都失败 | `model_visual` 只能记 `not_available`，总体状态停在 `unverified`；交付说明里明确写「未回渲验证」，附 HTML 预览图 |
| 字体不在批准注册表内 | `ppt_outline` 的字体替换告警；或 `html_create` 报缺少依赖 | 一律采用 `ppt_outline` 返回的最终字体，不自行指定 |
| Art Direction 校验失败 | `PPT_ART_DIRECTION_INVALID`（页数不匹配、`background_sequence` 与逐页 `background_role` 不一致、字体族不在注册表等） | 按返回的 issues 逐条改，重新调 `ppt_outline`；不复用半成品 |
| 大纲有 chart 但数据未落实 | `html_create` 报「大纲里仍有数据待补的 chart」 | 回框架 agent 补 `data_ref` 或同页 `data` 项，或把这页降级为 `point` |
| 材料不可得 | 材料 agent 交不出带来源的事实 | 降级为「只用用户提供的材料」，交付说明里标注来源缺失，绝不编造 |
| 需要重做某几页 | `html_create` 报输出已存在（同目录已有 `deck.html`） | 重新调 `ppt_outline` 拿新目录，或换 `artifact_title`；不要原地覆盖 |
| 搜图 agent 整体失败 | 候选池为空或全不可用 | 主 agent 自己补一次 `image_search`；仍不可用则走 5.1 |
| 材料 agent 整体失败 | `research/` 为空 | 主 agent 用 `web_search` / `browser_visit` 自采；时间不允许就退到用户材料 |

---

## 6. 一页版执行清单

主 agent 视角，照抄即可：

1. `ask_user_question` 补齐用途、受众、页数、语言、品牌。
2. `todo_write` 建清单，把三个子角色的任务与写入范围写清楚。
3. 派发材料 agent（写 `research/`）与搜图 agent（写 `images/`），两者并行，互不干涉。
4. 框架 agent 交 `frame/outline-payload.json` + `frame/art-direction.json`；主 agent 检查 8 个逐页字段是否齐全、页数与 `background_sequence` 是否对齐。
5. `ppt_fonts` 复核字体可用性 → `ppt_outline` 落盘，记下返回的 `artifact_dir`、`outline_path`、`design_plan_path`、`design_status`、`fonts`；`blocking_warnings` 非空就地返工。
6. 冻结素材到 `<root>/assets/images/`，写 `assets/source-manifest.json`（图源不可用就跳去零外链矢量方案）。
7. 写 HTML（逐页声明与 Art Direction 一致的构图、密度、背景；关键元素带上角色标记）→ `html_create`，带上 `design_plan_path`。
8. `ppt_create` 出 `deck.pptx`。
9. `ppt_image(backend=auto, refresh_quality=true)`。
10. `read_image` 看完全部联系表 + 封面/目录/数据页/复杂对比页/结束页 + 全部高风险页。
11. `write` 出 `visual-review.json`。
12. `ppt_create(..., finalize_visual_review=true)` 重算四层。
13. 只有总体状态为 `verified` 才对外声明通过；否则写明是哪一层没执行或失败。

---

## 7. 与实现对照

- 工具白名单：`src/schemas.ts` 的 `PPT_MODE_TOOL_NAMES`（本机 shell 项随平台在 `pwsh` / `bash` 之间切换）。
- 工件目录与来源记账结构：`src/artifacts.ts`。
- Art Direction 的字段与强校验：`src/art-direction.ts`。
- 四层状态与审稿文件契约：`src/quality.ts`。
- 回渲与联系表：`src/ppt-image.ts`。

本文提到的每个工具名，都可以在 `src/schemas.ts` 里逐字找到；本文没有描述表外工具能力。
