# DSH PPT · Turn Ideas into Editable Presentations

[中文](README.md) | **English**

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
  <strong>An AI Presentation Assistant for Professionals Powered by DeepSeek Harness</strong><br>
  <em>One-prompt generation · Native editable PPTX · Smart layout & palette · Free commercial images · Auto data charts · AI visual quality review</em>
</p>

<p align="center">

[Highlights](#highlights) · [Quick Start](#quick-start) · [Dependencies and requirements](#dependencies-and-requirements) · [Built-in theme library](#built-in-theme-library) · [Motion and fonts](#motion-and-fonts) · [Use Cases](#use-cases) · [Output Files](#output-files) · [Multi-Agent Workflow](#multi-agent-workflow) · [FAQ](#faq) · [License](#license)

</p>

## Highlights

Tired of spending hours searching for templates, aligning text boxes, and tweaking layouts? Most AI PPT tools either output flattened, uneditable image slides or rigid, generic templates.

**DSH PPT brings presentation making back to what matters — your content:**

- 📝 **Truly Native & Editable**: Generates standard `.pptx` files rather than whole-page images. Texts, shapes, and charts can be freely edited and adjusted in Microsoft PowerPoint, Apple Keynote, or WPS.
- 🎨 **Professional Aesthetics & Layouts**: No more cookie-cutter templates. AI intelligently tailors color schemes, typography hierarchies, and page compositions for a polished business look.
- 🎭 **12 Built-in Themes**: Each theme is a verified design recipe — colours drawn from IBM Carbon, Tailwind, Radix and Microsoft Fluent 2, fonts restricted to families this machine really renders, plus motion rhythm, ornament vocabulary and per-page layout notes. Pick one with `ppt_themes`; every palette is checked against WCAG AA (≥4.5:1) by machine.
- 📈 **Smart Charts & Curated Visuals**: Automatically searches for royalty-free commercial images and plots high-resolution data charts using Python, replacing walls of text with engaging visuals.
- 👁️ **AI Closed-Loop Visual Review**: The AI renders and visually inspects the generated PPT, proactively detecting and fixing text overflows, overlapping elements, or layout flaws.
- ✨ **Slide Motion**: Per-page entrance animation for text and shapes — 29 effects covering fades, fly-ins, zooms, spins, and wipes, with per-paragraph builds, click-step sequencing, and slide transitions.
- 🔤 **Any Installed Font**: The built-in font table is no longer the limit. The plugin enumerates every font on the machine (with PANOSE, pitch/family byte, charset, and embedding permission) and can install a font file into the current user's font directory in one step.
- 🔒 **Local Safety & Privacy**: Generation and rendering run securely in local sandboxes, keeping clear records of asset licensing and sources for confident business use.

## Quick Start

### 1. Prerequisites

- **DeepSeek Harness 0.2.0-rc.2 or newer** (this plugin declares the DSH peer range `^0.2.0-rc.2`; a mismatched runtime refuses to load it)
- Any of the following office software installed locally (for high-fidelity rendering and visual review):
  - **macOS**: Keynote or Microsoft PowerPoint
  - **Windows**: Microsoft PowerPoint
  - **Cross-platform / Linux**: LibreOffice

### 2. Installation

Run the following commands to install the plugin directly from npm:

```bash
# Install directly via DSH plugin manager
dsh plugin --profile web add @yejiming/dsh-ppt
dsh plugin --profile headless add @yejiming/dsh-ppt
```

### 3. Start Creating

#### Option 1: Web Interface (Recommended)
Launch the Web console, create a new session, choose **"PPT Mode"**, and describe your requirements to the AI just like speaking with a designer:
```bash
dsh --profile web
```

#### Option 2: Headless Command Line
Generate a full deck directly with a single command:
```bash
dsh --profile headless "Create a 6-page AI project quarterly review and roadmap deck for management"
```

## Dependencies and requirements

The plugin runs inside the DSH host process. Generation, re-rendering, and quality review all happen in the local working directory; only image search and `browser_visit` ever leave the machine. Everything below describes what the source actually does — none of it is a recommendation.

### Required

| Item | Requirement | Notes |
| :--- | :--- | :--- |
| DeepSeek Harness | `0.2.0-rc.2` or newer | The plugin declares 19 `@deepseek-ai/*` peer dependencies (18 at `^0.2.0-rc.2`, plus `@deepseek-ai/cordis` at `^4.0.1`). A mismatched runtime refuses to load the plugin, so neither PPT mode nor its 21 tools get mounted |
| Node.js | `>= 20.19.0` | Declared by `engines.node` in `package.json` |
| Operating system | Windows / macOS / Linux | `os` in `package.json` lists only `win32`, `darwin`, and `linux`. Any other platform fails at load time with `PPT_PLATFORM_UNSUPPORTED` |
| Chromium-based browser | Chrome, Chromium, or Edge installed locally | Used for HTML validation rendering, `browser_visit`, and HTML preview screenshots. `playwright install` is not needed: the plugin probes for a local executable and then launches it explicitly via `chromium.launch({ executablePath })` |
| `sharp` | Required, but declared optional (see below) | Image metadata validation, re-render normalisation, and contact-sheet compositing |

Runtime npm dependencies:

| Dependency | Loaded as | Used for |
| :--- | :--- | :--- |
| `pptxgenjs` | Static import | Writing the validated page structure into a native PPTX |
| `fflate` | Static import | Reading and writing the ZIP container of PPTX/OOXML, including XML rewrites for transitions and text animation |
| `fontkit` | Static import | Parsing font files, testing glyph coverage, producing font subsets |
| `zod` | Static import | Validating the outline, the Art Direction payload, and quality-review payloads |
| `schemastery` | Static import | Declaring the plugin's Config schema |
| `playwright-core` | Static import | Driving the local Chromium-based browser |
| `happy-dom` | Dynamic `import()` | Loaded on demand only while `html_create` validates HTML, so it stays out of host startup |

Every dependency except `happy-dom` is imported statically at module top level. That makes them **load-time hard dependencies**: if any one is missing, the whole plugin fails to load rather than degrading a single tool.

### Optional (capability add-ons)

**`sharp`: declared optional, actually required.** `package.json` lists it under `peerDependencies` and marks it `optional: true` in `peerDependenciesMeta`, yet `src/image-search.ts`, `src/python.ts`, `src/ppt-image.ts`, and `src/quality.ts` each do a top-level `import sharp from 'sharp'` with no try/catch. A missing `sharp` therefore fails the entire plugin to load — the declaration and the code disagree. **Treat it as a required dependency.**

**A PPT renderer.** Without one you still get a PPTX, but you lose every per-page re-render and all visual quality evidence; `ppt_image` reports `not_available`. Candidates and fallback order are in the next section.

**A Python analysis runtime.** Only `python_execute` needs it, for data charts and image processing. The executable defaults to `python` on Windows and `python3` on macOS and Linux, and can be overridden with the `pythonExecutable` option. `matplotlib`, `Pillow`, and `opencv-python` must already be installed: every execution runs an `import matplotlib, PIL, cv2` pre-flight first and fails with `PYTHON_DEPENDENCY_MISSING` otherwise. The environment is pinned to `MPLBACKEND=Agg`, so nothing opens a window.

### External programs and fallback order

| Program | Used for | Windows | macOS | Linux | Behaviour when missing |
| :--- | :--- | :--- | :--- | :--- | :--- |
| Chrome / Chromium / Edge | HTML validation rendering, `browser_visit`, HTML preview screenshots | `Google\Chrome\Application\chrome.exe`, `Chromium\Application\chrome.exe`, and `Microsoft\Edge\Application\msedge.exe` under `%ProgramFiles%`, `%ProgramFiles(x86)%`, and `%LOCALAPPDATA%` | `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`, `/Applications/Chromium.app/Contents/MacOS/Chromium`, `~/Applications/Google Chrome.app/...` | `/usr/bin/google-chrome`, `/usr/bin/google-chrome-stable`, `/usr/bin/chromium`, `/usr/bin/chromium-browser`, `/snap/bin/chromium` | Throws `BROWSER_NOT_READY`; the `browser` diagnostic reports `failed` |
| PowerShell | Driving PowerPoint COM to re-render on Windows | `powershell.exe`, `powershell`, `pwsh.exe`, `pwsh` | — | — | The Windows `powerpoint` backend is unavailable and LibreOffice takes over |
| Microsoft PowerPoint (desktop) | Re-rendering a PPTX to per-page PNGs | Launched through COM; probing does **not** require the `POWERPNT.EXE` path to exist | Launched through AppleScript, with output captured by `screencapture` | — | That backend is unavailable |
| Keynote | The preferred macOS re-render: exports each page as PNG | — | `/Applications/Keynote.app`, `~/Applications/Keynote.app` | — | macOS falls back to LibreOffice |
| `osascript` | Driving Keynote or PowerPoint on macOS | — | `/usr/bin/osascript` | — | Both the macOS `keynote` and `powerpoint` backends are unavailable |
| `screencapture` | macOS PowerPoint fallback: screenshots the slideshow window | — | `/usr/sbin/screencapture` | — | The macOS `powerpoint` backend is unavailable |
| LibreOffice (`soffice`) | First step of the cross-platform re-render: PPTX to PDF | `soffice.exe`, `soffice`, and `LibreOffice\program\soffice.exe` under `%ProgramFiles%`, `%ProgramFiles(x86)%`, and `%LOCALAPPDATA%` | `soffice`, `/Applications/LibreOffice.app/Contents/MacOS/soffice` | `soffice`, `libreoffice`, `/usr/bin/soffice`, `/usr/bin/libreoffice` | The LibreOffice backend is unavailable |
| Poppler (`pdftoppm`) | Second step of the re-render: PDF to PNG | `pdftoppm.exe`, `pdftoppm` | `pdftoppm` | `pdftoppm` | **The whole LibreOffice backend is unavailable**: `soffice` and `pdftoppm` must both be present |
| Python | Data charts and image processing | `python` | `python3` | `python3` | `python_execute` throws `PYTHON_DEPENDENCY_MISSING` |

Browser probing order: the `browserExecutable` option → the browser path `playwright-core` records → the system candidates listed above.

Re-render order with `backend=auto`:

- **macOS**: `keynote` → `libreoffice` → `powerpoint` (screen-capture fallback, which also warns that animations and multi-display setups can affect the captured frame)
- **Windows**: `powerpoint` (PowerPoint COM, driven by PowerShell) → `libreoffice`
- **Linux**: `libreoffice` only

When no candidate is available the plugin does **not** throw: `ppt_image` returns `not_available`, the quality report sets `render_status: not_available` and appends a `RENDERER_NOT_AVAILABLE` warning, and the automatic visual layer becomes `not_available` too. In other words, a deck is still produced without a renderer — just with no per-page rendering evidence.

### Network and privacy

There are exactly two outbound paths:

| Operation | Target | Credentials | Behaviour on failure |
| :--- | :--- | :--- | :--- |
| `image_search` | `api.openverse.org/v1/images/` first, then `commons.wikimedia.org/w/api.php` when results are short | None; no API key | Does not throw: it returns whatever it found along with warnings such as `degraded:unavailable`, plus three fallback routes (`read_image` to reuse a frozen asset, `browser_visit` to capture from a public page, or an injected custom provider). `IMAGE_SEARCH_FAILED` is thrown only when `strictOnUnavailable` is explicitly enabled |
| `browser_visit` | Any public http(s) page the model names | None | Throws `BROWSER_URL_BLOCKED` (URL judged unreachable) or `BROWSER_NOT_READY` (no usable browser) |

Request limits for `image_search`: a 15-second timeout, `redirect: 'error'` (redirects are never followed), a 2 MiB per-response cap, and a 10-minute per-query result cache. Degraded results are not cached, so the next call genuinely retries the providers.

`browser_visit` and the candidate image URLs of `image_search` share one SSRF guard: only `http` and `https` are allowed; URLs carrying a username or password are rejected; only ports 80 and 443 are allowed; `localhost`, `*.localhost`, `*.local`, and `*.internal` are rejected; and every resolved address is checked, so loopback, private ranges, link-local ranges (including the cloud metadata address), CGNAT, multicast, and reserved space are all refused. Every request inside a browser context passes through the same interception route and is aborted on a hit.

Nothing on the HTML side links out: `html_create` rejects non-local resource references and any CSS containing remote, data, or file resources, and `dsh-tool-web` has `fetch` disabled in PPT mode so the tool only searches.

Every other capability — PPTX generation, font parsing and subsetting, re-rendering, automatic quality review, contact-sheet compositing — completes inside the local working directory.

### Platform matrix

| Capability | Windows | macOS | Linux | Behaviour when unmet |
| :--- | :--- | :--- | :--- | :--- |
| HTML rendering and browser access | Any of Chrome / Chromium / Edge | Same | Same | Throws `BROWSER_NOT_READY` |
| PPTX re-render (`auto`) | PowerPoint COM → LibreOffice | Keynote → LibreOffice → PowerPoint screen capture | LibreOffice only | Does not throw; `render_status=not_available` plus `RENDERER_NOT_AVAILABLE` |
| LibreOffice backend prerequisites | `soffice` and `pdftoppm` both present | Same | Same | That backend is unavailable, which on Linux means re-rendering is gone entirely |
| Python charts and image processing | `python` plus matplotlib / Pillow / opencv-python | `python3` plus the same | `python3` plus the same | Throws `PYTHON_DEPENDENCY_MISSING` |
| System font directories | `%SystemRoot%\Fonts`, `%LOCALAPPDATA%\Microsoft\Windows\Fonts` | `/System/Library/Fonts`, `/System/Library/Fonts/Supplemental`, `/System/Library/AssetsV2/com_apple_MobileAsset_Font7`, `...Font8`, `/Library/Fonts`, `~/Library/Fonts` | `/usr/share/fonts`, `/usr/local/share/fonts`, `~/.local/share/fonts`, `~/.fonts` | The `fontDirs` option **appends** to this list; declaring an uninstalled approved primary font in the HTML throws `PPT_DEPENDENCY_MISSING` |
| Font installation (custom fonts) | Copied into `%LOCALAPPDATA%\Microsoft\Windows\Fonts`, then registered under HKCU with `reg.exe` | Copied into `~/Library/Fonts` | Copied into `~/.local/share/fonts`, then a best-effort `fc-cache -f` | Affects only the availability of custom fonts, never deck generation; on Windows a missing `reg.exe` fails that step |

### Pre-flight checks

**Windows (PowerShell)**

```powershell
node -v
dsh --version

# Chromium-based browser: any one of these is enough
Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
Test-Path "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
Test-Path "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"

# Re-render backends: PowerShell plus PowerPoint, or LibreOffice plus Poppler
Get-Command powershell.exe
Test-Path "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE"
Get-Command soffice.exe, pdftoppm.exe -ErrorAction SilentlyContinue

# Python analysis dependencies
python -c "import matplotlib, PIL, cv2; print('python deps ok')"

# sharp (run this from the profile directory the plugin is installed into)
node --input-type=module -e "import('sharp').then(() => console.log('sharp ok')).catch(() => { console.error('sharp missing'); process.exit(1) })"
```

**macOS**

```bash
node -v
dsh --version

# Browser: any one of these is enough
ls -d "/Applications/Google Chrome.app" /Applications/Chromium.app 2>/dev/null

# Re-render backends: Keynote / PowerPoint, or LibreOffice plus Poppler
ls -d /Applications/Keynote.app "/Applications/Microsoft PowerPoint.app" 2>/dev/null
which osascript screencapture pdftoppm soffice

# Python analysis dependencies
python3 -c "import matplotlib, PIL, cv2; print('python deps ok')"

# sharp
node --input-type=module -e "import('sharp').then(() => console.log('sharp ok')).catch(() => { console.error('sharp missing'); process.exit(1) })"
```

**Linux**

```bash
node -v
dsh --version

# Browser: any one of these is enough
which google-chrome google-chrome-stable chromium chromium-browser

# Re-render backend: soffice and pdftoppm must both be present
which soffice libreoffice pdftoppm

# Python analysis dependencies
python3 -c "import matplotlib, PIL, cv2; print('python deps ok')"

# sharp
node --input-type=module -e "import('sharp').then(() => console.log('sharp ok')).catch(() => { console.error('sharp missing'); process.exit(1) })"
```

Once all of that is green, the PPT runtime diagnostics (`platform`, `tools`, `browser`, `python`, `fonts`, `renderer`, `attachments`, and `vision_model`) should all read `ready` — `vision_model` additionally requires the active model to declare image input. Anything that comes back `failed` is surfaced before work starts.

## Built-in theme library

The plugin ships 12 design themes. A theme is not a shell template — it is a **verifiable design recipe**: colours, four type roles, a per-page layout bias, rhythm limits and the ornament vocabulary that theme permits. It gives the Art Direction pass a vetted starting point rather than doing the design for the model: `concept`, `audience_effect` and every page `job` and `takeaway` still come from the actual content.

| Theme id | Name | Scenes | Base / inverted ground | Accent |
| :--- | :--- | :--- | :--- | :--- |
| `carbon-blueprint` | Carbon blueprint | technical proposals, architecture review, RFC | `#FFFFFF` / `#161616` | `#0043CE` |
| `slate-review` | Slate review | quarterly review, business analysis, OKR | `#F8FAFC` / `#0F172A` | `#1D4ED8` |
| `midnight-raise` | Midnight raise | fundraising, strategy, board | `#0B1220` / `#FFFFFF` | `#78A9FF` |
| `paper-ink` | Paper and ink | academic report, whitepaper, long-form | `#F7F5F0` / `#111318` | `#9F1853` |
| `editorial-serif` | Scholarly editorial | defence, research report, policy | `#F7F3F2` / `#171414` | `#8A3800` |
| `telemetry-teal` | Telemetry teal | ops review, dashboards, performance | `#FFFFFF` / `#081A1C` | `#005D5D` |
| `graphite-minimal` | Graphite minimal | specs, internal alignment, engineering docs | `#FCFCFC` / `#202020` | `#3A5BC7` |
| `indigo-launch` | Indigo launch | product launch, release keynote | `#1E1B4B` / `#EEF2FF` | `#A5B4FC` |
| `amber-academy` | Amber academy | training, workshop, onboarding | `#FFFBEB` / `#451A03` | `#92400E` |
| `crimson-brief` | Crimson brief | marketing proposal, one-pager decision | `#FFFFFF` / `#4C0519` | `#BE123C` |
| `fluent-azure` | Fluent azure | enterprise report, client proposal, tender | `#FFFFFF` / `#0A2E4A` | `#0F6CBD` |
| `moss-annual` | Moss annual | annual report, ESG disclosure, long-horizon plan | `#F0FDF4` / `#14532D` | `#15803D` |

Colours are not improvised: every palette traces to a published design system (IBM Carbon, Tailwind, Radix, Microsoft Fluent 2, Open Color) and the code records that origin in `palette_source`. Each theme also carries an **inverted accent**, because one accent cannot clear 4.5:1 on both a light and a dark ground — that is a property of colour, not a preference — so the per-page plan says which one that page uses.

### Usage

```text
ppt_themes()                                              # list every theme with scene tags
ppt_themes(scene="fundraising")                           # filter by scene
ppt_themes(theme_id="carbon-blueprint")                   # read the full recipe
ppt_themes(theme_id="carbon-blueprint", page_types=[...])  # per-page visual plan
ppt_outline(..., theme_id="carbon-blueprint", art_direction={...})
```

Passing `theme_id` to `ppt_outline` checks conformance and reports every drift instead of passing silently: `THEME_PALETTE_DRIFT`, `THEME_ACCENT_REPLACED`, `THEME_TYPOGRAPHY_REPLACED`, `THEME_ACCENT_DRIFT`, `THEME_FONT_DRIFT`, and `THEME_PLAN_MISSING` when a theme was named but no `art_direction` was supplied.

### Palette sources and brand colours

To follow a client brand instead, read its public style guide with `browser_visit`, then build the plan on those values under the same discipline: body text and accent at 4.5:1 or better against every ground, and a separate accent for the inverted group. `ppt_themes` returns those rules with every response.

### Design constraints

- Canvas 1280×720, margins 72px horizontal / 64px vertical, content area 1136×592, 12 columns with 24px gutters, radii limited to 4 / 8 / 12px
- Type scale: cover title 72px, page title 44px, section title 32px, lead 26px, body 24px, caption 16px, metric 88px, code 17px
- Ornament is a closed set: 1px and 2px borders, radii, tonal steps, opacity, z-index, type contrast, grid alignment, and solid or `linear-gradient` backgrounds
- `box-shadow`, `transform`, `filter`, `clip-path` and `content` are rejected outright, so depth comes from borders and tonal steps and the HTML preview cannot diverge from the PPTX

## Motion and fonts

### Slide motion

`ppt_create` takes an optional `effects` argument that declares, per page, what plays. Transitions and entrance animations are injected after pptxgenjs has written the package but before the atomic commit, and only `ppt/slides/slideN.xml` is rewritten — every other part stays byte-identical. A page that declares nothing is left completely untouched.

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

- `target` is an element id from the deck IR, which is also the shape name in the PPTX.
- There are 29 entrance effects: `appear`, `flash-once`, `fade`, `dissolve`, `wedge`, `wipe`, `blinds`, `checkerboard`, `random-bars`, `box`, `circle`, `diamond`, `plus`, `split`, `strips`, `wheel`, `zoom`, `fly-in`, `crawl`, `peek`, `stretch`, `swivel`, `spiral`, `bounce`, `credits`, `float-in`, `grow-turn`, `rise-up`, `unfold`.
- `zoom`, `fly-in`, `crawl`, `peek`, `stretch`, `swivel`, `spiral`, `bounce`, `credits`, `float-in`, `grow-turn`, `rise-up`, and `unfold` are **motion effects**: the text genuinely travels or scales rather than merely appearing.
- `direction` applies only to effects that render one: `wipe`, `fly-in`, `crawl`, `peek`, `blinds`, `checkerboard`, `random-bars`, `box`, `circle`, `diamond`, `plus`, `stretch`, `swivel`.
- `start` is `on-click` (the default), `with-previous`, or `after-previous`. The first animation of a page always opens a click step.
- A page carries at most 24 animations, one per element and build unit; `by_paragraph` makes an effect play paragraph by paragraph.

The generator follows the timelines PowerPoint actually writes rather than the prose of the specification: the `presetID` numbering space is partitioned by `presetClass`, `presetSubtype` is a direction bitmask (`1` top, `2` right, `4` bottom, `8` left, `16` in, `32` out), and what really drives rendering is the `p:animEffect` filter plus the from/to values of `p:anim`.

### Fonts: choose freely, install quickly

`ppt_fonts` has two scopes:

- `scope=registry` (the default): the plugin's curated table of recommended families with deterministic per-role recommendations.
- `scope=installed`: **every font face installed on this machine**, each with its PANOSE, pitch/family byte, charset, weight, glyph count, Latin and CJK coverage, and the font's own embedding permission bits (`fs_type` / `embeddable`).

`ppt_outline`'s `title_font` / `body_font` are consequently no longer restricted to the built-in table: **any family installed on this machine can be named directly**. The plugin scans the system font directories — reading every font file — only when a named family is missing from the table, so the common path costs nothing extra.

To install a font file sitting on disk into the current user's font directory, so the deck can then reference it:

```jsonc
{ "scope": "installed", "install_path": "assets/fonts/BrandSans.ttf", "dry_run": true }
```

- Installation is **per-user** and never needs administrator rights: Windows writes to `%LOCALAPPDATA%\Microsoft\Windows\Fonts` and registers the face under `HKCU`; macOS writes to `~/Library/Fonts`; Linux writes to `~/.local/share/fonts` and refreshes `fc-cache` on a best-effort basis.
- Run with `dry_run: true` first to see the target path and family name without writing anything.
- The result carries an `uninstall_hint` for manual removal.

On embedding a font **inside the PPTX file itself**: PowerPoint's native font part (`ppt/fonts/*.fntdata`) is an EOT container holding MicroType Express compressed TrueType data, and this plugin does not generate that format. When fidelity has to survive a machine change, install the font on the target machine with `ppt_fonts`, or ship the font file alongside the artifact and install it there.

## Use Cases

| Category | Example Prompt |
| :--- | :--- |
| 📊 **Business Review / Debrief** | *"Create a 6-page Q3 e-commerce review deck highlighting GMV growth, funnel conversion, and next quarter strategy."* |
| 💼 **Pitch Deck / Proposal** | *"Draft an 8-page enterprise private AI knowledge base proposal deck with a sleek business-tech aesthetic."* |
| 📢 **Product Launch** | *"Design a product launch deck for a new mobile app, emphasizing key selling points and UX highlights."* |
| 🎓 **Training / Team Guidelines** | *"Create a 5-page agile development & team collaboration handbook deck with a clean, lively style."* |

## Output Files

Each completed task creates a standalone folder inside `ppt-output/`:

- 📄 **`deck.pptx`**: The final native PPT file. Double-click to open in PowerPoint, Keynote, or WPS for presentation or direct editing.
- 🖼️ **`preview/`**: High-resolution image previews of each slide, convenient for quick mobile browsing and sharing.
- 📁 **`assets/`**：Image assets and charts used in the deck, complete with attribution and license tracking.

## Multi-Agent Workflow

A deck can also be produced by several agents working in parallel: one scouting images, one gathering source material, one authoring the frame (outline + Art Direction), and a lead agent that integrates, verifies, and delivers the PPTX. The plugin supplies a fixed 21-tool surface and a strictly validated artifact contract; the split itself is arranged at the session layer.

- 🔍 **Image scout**: writes image candidates into a candidate pool only, never into the artifact directory. Searching and freezing are separate steps — freezing belongs to the lead agent.
- 📚 **Research agent**: writes sourced notes and structured data points only.
- 🧭 **Frame agent**: produces the outline payload and Art Direction only; it creates no deliverable files.
- 🎯 **Lead agent**: exclusively owns `ppt_outline` / `html_create` / `ppt_create` / `ppt_image`, and follows the gate order `outline → html → pptx → re-render → per-page review → finalize`.

Roles, inputs and outputs, write boundaries (who writes `outline.json`, who writes `deck.html`, who only produces a candidate list), dependency order, and failure fallbacks — including the real case of "image sources unavailable → fall back to a zero-external-link vector design" — are documented in **[Multi-Agent Workflow](docs/agent-workflow.md)** (the document is written in Chinese).

## FAQ

<details>
<summary><b>Q: Will the generated PPT have missing fonts or broken layouts in Office / WPS?</b></summary>
No. The system incorporates multi-platform fallback font strategies (prioritizing common fonts like Segoe UI, Helvetica, PingFang, and Source Han), ensuring consistent and aesthetic typography across devices.
</details>

<details>
<summary><b>Q: Can I export the slides to PDF or present directly?</b></summary>
Yes! The output is a standard <code>.pptx</code> file. You can enter presenter mode in PowerPoint, Keynote, or WPS, or export directly to PDF or speaker notes.
</details>

<details>
<summary><b>Q: How can I adjust slides if I want changes?</b></summary>
You have two flexible options:
1. **Chat with AI**: Ask the AI directly in the conversation (e.g., "Change slide 3 to a bar chart comparison", "Switch theme color to navy blue");
2. **Edit Locally**: Open <code>deck.pptx</code> directly and modify text, swap images, or tweak layouts like any standard slide deck.
</details>

## License

This project is licensed under the [MIT License](LICENSE).
