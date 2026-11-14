import { createHash, randomUUID } from 'node:crypto'
import { access, copyFile, mkdir, readFile, readdir, rm, stat } from 'node:fs/promises'
import { basename, delimiter, dirname, join } from 'node:path'
import { pathToFileURL } from 'node:url'
import type { SandboxProvider } from '@deepseek-ai/dsh-sandbox'
import type { SubprocessRuntime } from '@deepseek-ai/dsh-subprocess'
import { unzipSync } from 'fflate'
import sharp from 'sharp'
import { atomicWriteFile, atomicWriteJson, atomicWriteText } from './atomic.ts'
import { PptError, throwIfAborted } from './errors.ts'
import { DEFAULT_LIMITS, boundedInteger } from './limits.ts'
import {
  appleScriptCandidates, keynoteCandidates, libreOfficeCandidates, pdfToPpmCandidates,
  powerPointCandidates, powerShellCandidates, pptImageBackendOrder, screenCaptureCandidates, type PptImageBackend,
  systemFontDirectories,
} from './platform.ts'
import { inspectPptxPackage } from './pptx.ts'
import { resolveWorkspacePath, workspaceRelative } from './paths.ts'
import { safeErrorMessage } from './security.ts'
import type { SessionOwner } from './session-resources.ts'
import { SessionResourceRegistry } from './session-resources.ts'
import { runCollected } from './subprocess.ts'

export type PptImageStatus = 'passed' | 'failed' | 'not_available'
export type PptImageCaptureMethod = 'native-export' | 'pdf-raster' | 'screen-capture'

export interface PptImageAttempt {
  backend: PptImageBackend
  capture_method?: PptImageCaptureMethod
  status: PptImageStatus
  message: string
}

export interface PptImageResult {
  status: PptImageStatus
  backend?: PptImageBackend
  backend_version?: string
  capture_method?: PptImageCaptureMethod
  page_count: number
  image_paths: string[]
  contact_sheet_paths: string[]
  manifest_path?: string
  cached: boolean
  attempts: PptImageAttempt[]
  warnings: string[]
}

export interface PptImageRenderOptions {
  backend?: 'auto' | PptImageBackend
  force?: boolean
  outputDirectory?: string
  /** Host tool policy approved one unconfined native-app automation call. */
  nativeAutomationApproved?: boolean
  /** One-based macOS display selected by the PowerPoint screen-capture fallback. */
  screenIndex?: number
}

interface AvailableBackend {
  backend: PptImageBackend
  captureMethod: PptImageCaptureMethod
  version: string
  executables: Record<string, string>
}

interface RenderManifest {
  version: 2
  pptx_path: string
  pptx_sha256: string
  backend: PptImageBackend
  backend_version: string
  capture_method: PptImageCaptureMethod
  page_count: number
  image_paths: string[]
  contact_sheet_paths: string[]
  warnings: string[]
  generated_at: string
}

const KEYNOTE_SCRIPT = `on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set inputFile to POSIX file inputPath as alias
  set outputFile to POSIX file outputPath
  tell application "Keynote"
    set sourceDocument to open inputFile
    export sourceDocument to outputFile as slide images with properties {image format:PNG}
    close sourceDocument saving no
  end tell
end run
`

const POWERPOINT_SCRIPT = `param(
  [Parameter(Mandatory=$true)][string]$InputPptx,
  [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference = "Stop"
$application = $null
$presentation = $null
try {
  $application = New-Object -ComObject PowerPoint.Application
  $presentation = $application.Presentations.Open($InputPptx, $true, $true, $false)
  $presentation.Export($OutputDir, "PNG", 1280, 720)
  Write-Output ("PowerPoint " + $application.Version)
} finally {
  if ($presentation -ne $null) { $presentation.Close() }
  if ($application -ne $null) { $application.Quit() }
  if ($presentation -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
  if ($application -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application) }
}
`

const POWERPOINT_MAC_SCREEN_SCRIPT = `on run argv
  set inputPath to item 1 of argv
  set outputDirectory to item 2 of argv
  set pageCount to (item 3 of argv) as integer
  set captureBinary to item 4 of argv
  set screenIndex to (item 5 of argv) as integer
  set inputFile to POSIX file inputPath as alias
  set sourcePresentation to missing value
  set showView to missing value
  tell application "Microsoft PowerPoint"
    try
      activate
      open inputFile
      set sourcePresentation to active presentation
      set settings to slide show settings of sourcePresentation
      try
        set advance mode of settings to slide show advance manual advance
      end try
      try
        set range type of settings to slide show range show all
      end try
      try
        set loop until stopped of settings to false
      end try
      try
        set show with presenter of settings to false
      end try
      set showWindow to run slide show settings
      set showView to slideshow view of showWindow
      delay 1
      repeat with pageNumber from 1 to pageCount
        if pageNumber > 1 then
          go to next slide showView
          delay 1
        end if
        set targetPath to outputDirectory & "/page-" & my zeroPad(pageNumber) & ".png"
        do shell script quoted form of captureBinary & " -x -D " & screenIndex & " " & quoted form of targetPath
      end repeat
      exit slide show showView
      close sourcePresentation saving no
    on error errorMessage number errorNumber
      try
        if showView is not missing value then exit slide show showView
      end try
      try
        if sourcePresentation is not missing value then close sourcePresentation saving no
      end try
      error errorMessage number errorNumber
    end try
  end tell
end run

on zeroPad(pageNumber)
  if pageNumber < 10 then return "00" & pageNumber
  if pageNumber < 100 then return "0" & pageNumber
  return pageNumber as text
end zeroPad
`

function xmlEscape(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

export function fontconfigDocument(fontDirs: readonly string[], cacheDir: string): string {
  const directories = [...new Set(fontDirs.map(path => path.replaceAll('\\', '/')))]
  return `<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n${directories.map(path => `  <dir>${xmlEscape(path)}</dir>`).join('\n')}\n  <cachedir>${xmlEscape(cacheDir.replaceAll('\\', '/'))}</cachedir>\n  <config><rescan><int>30</int></rescan></config>\n</fontconfig>\n`
}

export function pptxPageCount(data: Uint8Array): number {
  let files: Record<string, Uint8Array>
  try { files = unzipSync(data) } catch (error) {
    throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX is not a readable ZIP package', { cause: error })
  }
  const pages = Object.keys(files).filter(name => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).length
  boundedInteger(pages, 'PPTX page count', 1, DEFAULT_LIMITS.maxSlides)
  inspectPptxPackage(data, pages)
  return pages
}

async function firstExecutable(subprocess: SubprocessRuntime, candidates: readonly string[], signal?: AbortSignal): Promise<string | undefined> {
  for (const candidate of candidates) {
    try { return await subprocess.resolveExecutable(candidate, undefined, signal) } catch { /* Try the next path. */ }
  }
  return undefined
}

async function firstExisting(candidates: readonly string[]): Promise<string | undefined> {
  for (const candidate of candidates) {
    try { await access(candidate); return candidate } catch { /* Try the next path. */ }
  }
  return undefined
}

async function fingerprint(path: string | undefined, fallback: string): Promise<string> {
  if (path === undefined) return fallback
  try {
    const info = await stat(path)
    return `${basename(path)}:${Math.round(info.mtimeMs)}:${info.size}`
  } catch {
    return fallback
  }
}

async function filesRecursively(root: string): Promise<string[]> {
  const output: string[] = []
  for (const item of await readdir(root, { withFileTypes: true })) {
    const path = join(root, item.name)
    if (item.isDirectory()) output.push(...await filesRecursively(path))
    else if (item.isFile()) output.push(path)
  }
  return output
}

function naturalPageNumber(path: string): number {
  const numbers = basename(path).match(/\d+/gu)
  return numbers === null ? Number.MAX_SAFE_INTEGER : Number(numbers.at(-1))
}

function processFailure(name: string, exitCode: number | null, stdout: string, stderr: string): string {
  const diagnostic = safeErrorMessage(stderr.trim() || stdout.trim(), 1_000)
  return `${name} exited with ${exitCode ?? 'no exit code'}${diagnostic.length === 0 ? '' : `: ${diagnostic}`}`
}

async function createContactSheets(workspace: string, previews: string[], targetDir: string, signal?: AbortSignal): Promise<string[]> {
  const results: string[] = []
  await mkdir(targetDir, { recursive: true })
  for (let start = 0; start < previews.length; start += 4) {
    throwIfAborted(signal)
    const paths = previews.slice(start, start + 4)
    const composites = await Promise.all(paths.map(async (path, index) => ({
      input: await sharp(join(workspace, path)).resize(620, 349, { fit: 'contain', background: '#FFFFFF' }).png().toBuffer(),
      left: index % 2 * 640 + 10, top: Math.floor(index / 2) * 360 + 5,
    })))
    const target = join(targetDir, `contact-sheet-${String(start / 4 + 1).padStart(3, '0')}.png`)
    const data = await sharp({ create: { width: 1280, height: 720, channels: 4, background: '#E5E7EB' } }).composite(composites).png().toBuffer()
    await atomicWriteFile(target, data, { overwrite: true, signal })
    results.push(workspaceRelative(workspace, target))
  }
  return results
}

export class PptImageRuntime {
  constructor(
    private readonly subprocess: SubprocessRuntime | undefined,
    private readonly sandbox: SandboxProvider | undefined,
    private readonly resources: SessionResourceRegistry,
    private readonly executables: {
      soffice?: readonly string[]
      pdftoppm?: readonly string[]
      osascript?: readonly string[]
      powershell?: readonly string[]
      keynote?: readonly string[]
      powerpoint?: readonly string[]
      screencapture?: readonly string[]
    } = {},
    private readonly fontDirs: readonly string[] = [],
    private readonly platform: NodeJS.Platform = process.platform,
  ) {}

  async render(
    owner: SessionOwner,
    workspace: string,
    pptxPathInput: string,
    options: PptImageRenderOptions = {},
    signal?: AbortSignal,
  ): Promise<PptImageResult> {
    throwIfAborted(signal)
    const pptxPath = await resolveWorkspacePath(workspace, pptxPathInput, { mustExist: true, kind: 'file' })
    if (!pptxPath.toLowerCase().endsWith('.pptx')) throw new PptError('PPT_PATH_INVALID', 'ppt_image requires a .pptx file')
    const bytes = new Uint8Array(await readFile(pptxPath))
    const pageCount = pptxPageCount(bytes)
    const hash = createHash('sha256').update(bytes).digest('hex')
    const artifactRoot = dirname(pptxPath)
    const outputDirectory = await resolveWorkspacePath(workspace, options.outputDirectory ?? join(workspaceRelative(workspace, artifactRoot), 'preview', 'pptx'))
    if (dirname(dirname(outputDirectory)) !== artifactRoot && dirname(outputDirectory) !== artifactRoot) {
      throw new PptError('PPT_PATH_INVALID', 'ppt_image output directory must remain inside the PPTX artifact directory')
    }
    await mkdir(outputDirectory, { recursive: true })
    const manifestPath = join(outputDirectory, 'render-manifest.json')
    const attempts: PptImageAttempt[] = []
    if (this.subprocess === undefined || this.sandbox === undefined) {
      return { status: 'not_available', page_count: pageCount, image_paths: [], contact_sheet_paths: [], cached: false, attempts: [{ backend: 'libreoffice', status: 'not_available', message: 'DSH subprocess or sandbox service is unavailable' }], warnings: [] }
    }

    const requested = options.backend ?? 'auto'
    for (const backend of pptImageBackendOrder(this.platform, requested)) {
      throwIfAborted(signal)
      const available = await this.discover(backend, signal)
      if (available === undefined) {
        attempts.push({ backend, status: 'not_available', message: `${backend} rendering dependencies were not found` })
        continue
      }
      if (options.force !== true) {
        const cached = await this.cachedResult(workspace, manifestPath, hash, available, attempts)
        if (cached !== undefined) return cached
      }
      const temporary = join(artifactRoot, `.ppt-image-${randomUUID()}`)
      const rawDirectory = join(temporary, 'raw')
      const normalizedDirectory = join(temporary, 'normalized')
      await Promise.all([mkdir(rawDirectory, { recursive: true }), mkdir(normalizedDirectory, { recursive: true })])
      this.resources.open(owner, workspace)
      const untrack = this.resources.trackTemporaryPath(owner, temporary)
      try {
        await this.renderBackend(
          available, workspace, pptxPath, temporary, rawDirectory,
          options.nativeAutomationApproved === true, options.screenIndex, signal,
        )
        const rawImages = (await filesRecursively(rawDirectory))
          .filter(path => /\.png$/iu.test(path))
          .sort((left, right) => naturalPageNumber(left) - naturalPageNumber(right) || left.localeCompare(right))
        if (rawImages.length !== pageCount) throw new PptError('PPT_RENDER_FAILED', `${backend} rendered ${rawImages.length} pages, expected ${pageCount}`)
        const normalized: string[] = []
        for (let index = 0; index < rawImages.length; index += 1) {
          throwIfAborted(signal)
          const source = rawImages[index]!
          const metadata = await sharp(source, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).metadata()
          if (metadata.width === undefined || metadata.height === undefined) throw new PptError('PPT_RENDER_FAILED', `${basename(source)} has no decodable dimensions`)
          const ratio = metadata.width / metadata.height
          const target = join(normalizedDirectory, `page-${String(index + 1).padStart(3, '0')}.png`)
          let pipeline = sharp(source, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).flatten({ background: '#FFFFFF' })
          if (available.captureMethod === 'screen-capture') {
            const targetRatio = 16 / 9
            if (ratio > targetRatio) {
              const width = Math.max(1, Math.floor(metadata.height * targetRatio))
              pipeline = pipeline.extract({ left: Math.floor((metadata.width - width) / 2), top: 0, width, height: metadata.height })
            } else {
              const height = Math.max(1, Math.floor(metadata.width / targetRatio))
              pipeline = pipeline.extract({ left: 0, top: Math.floor((metadata.height - height) / 2), width: metadata.width, height })
            }
          } else if (Math.abs(ratio / (16 / 9) - 1) > 0.03) {
            throw new PptError('PPT_RENDER_FAILED', `${basename(source)} is not a supported 16:9 slide image`)
          }
          await pipeline.resize(1280, 720, { fit: 'fill' }).toColourspace('srgb').png().toFile(target)
          normalized.push(target)
        }
        for (const name of await readdir(outputDirectory)) {
          if (/^page-\d+\.png$/u.test(name)) await rm(join(outputDirectory, name), { force: true })
        }
        const imagePaths: string[] = []
        for (let index = 0; index < normalized.length; index += 1) {
          const target = join(outputDirectory, `page-${String(index + 1).padStart(3, '0')}.png`)
          await atomicWriteFile(target, await readFile(normalized[index]!), { overwrite: true, signal })
          imagePaths.push(workspaceRelative(workspace, target))
        }
        const contactRoot = dirname(outputDirectory)
        for (const name of await readdir(contactRoot)) {
          if (/^contact-sheet-\d+\.png$/u.test(name)) await rm(join(contactRoot, name), { force: true })
        }
        const contactSheets = await createContactSheets(workspace, imagePaths, contactRoot, signal)
        const warnings = available.captureMethod === 'screen-capture'
          ? ['PowerPoint screen capture is a last-resort renderer; slide animations and multi-display configuration can affect the captured frame.']
          : []
        const manifest: RenderManifest = {
          version: 2, pptx_path: workspaceRelative(workspace, pptxPath), pptx_sha256: hash,
          backend, backend_version: available.version, capture_method: available.captureMethod, page_count: pageCount,
          image_paths: imagePaths, contact_sheet_paths: contactSheets, warnings, generated_at: new Date().toISOString(),
        }
        await atomicWriteJson(manifestPath, manifest, { overwrite: true, signal })
        attempts.push({ backend, capture_method: available.captureMethod, status: 'passed', message: `${backend} rendered ${pageCount} pages` })
        return {
          status: 'passed', backend, backend_version: available.version, capture_method: available.captureMethod, page_count: pageCount,
          image_paths: imagePaths, contact_sheet_paths: contactSheets,
          manifest_path: workspaceRelative(workspace, manifestPath), cached: false, attempts, warnings,
        }
      } catch (error) {
        if (signal?.aborted) throwIfAborted(signal)
        attempts.push({ backend, capture_method: available.captureMethod, status: 'failed', message: safeErrorMessage(error) })
      } finally {
        untrack()
        await rm(temporary, { recursive: true, force: true })
      }
    }
    return {
      status: attempts.some(attempt => attempt.status === 'failed') ? 'failed' : 'not_available',
      page_count: pageCount, image_paths: [], contact_sheet_paths: [], cached: false, attempts, warnings: [],
    }
  }

  private async discover(backend: PptImageBackend, signal?: AbortSignal): Promise<AvailableBackend | undefined> {
    if (backend === 'keynote') {
      if (this.platform !== 'darwin') return undefined
      const [osascript, keynote] = await Promise.all([
        firstExecutable(this.subprocess!, this.executables.osascript ?? appleScriptCandidates(this.platform), signal),
        firstExisting(this.executables.keynote ?? keynoteCandidates(this.platform)),
      ])
      if (osascript === undefined || keynote === undefined) return undefined
      return { backend, captureMethod: 'native-export', version: await fingerprint(keynote, 'Keynote'), executables: { osascript, keynote } }
    }
    if (backend === 'powerpoint') {
      if (this.platform === 'darwin') {
        const [osascript, powerpoint, screencapture] = await Promise.all([
          firstExecutable(this.subprocess!, this.executables.osascript ?? appleScriptCandidates(this.platform), signal),
          firstExisting(this.executables.powerpoint ?? powerPointCandidates(this.platform)),
          firstExecutable(this.subprocess!, this.executables.screencapture ?? screenCaptureCandidates(this.platform), signal),
        ])
        if (osascript === undefined || powerpoint === undefined || screencapture === undefined) return undefined
        return {
          backend, captureMethod: 'screen-capture', version: await fingerprint(powerpoint, 'PowerPoint-screen-capture'),
          executables: { osascript, powerpoint, screencapture },
        }
      }
      if (this.platform !== 'win32') return undefined
      const powershell = await firstExecutable(this.subprocess!, this.executables.powershell ?? powerShellCandidates(this.platform), signal)
      if (powershell === undefined) return undefined
      const powerpoint = await firstExisting(this.executables.powerpoint ?? powerPointCandidates(this.platform))
      return { backend, captureMethod: 'native-export', version: await fingerprint(powerpoint, 'PowerPoint-COM'), executables: { powershell, ...(powerpoint === undefined ? {} : { powerpoint }) } }
    }
    const [soffice, pdftoppm] = await Promise.all([
      firstExecutable(this.subprocess!, this.executables.soffice ?? libreOfficeCandidates(this.platform), signal),
      firstExecutable(this.subprocess!, this.executables.pdftoppm ?? pdfToPpmCandidates(this.platform), signal),
    ])
    if (soffice === undefined || pdftoppm === undefined) return undefined
    return { backend, captureMethod: 'pdf-raster', version: await fingerprint(soffice, 'LibreOffice'), executables: { soffice, pdftoppm } }
  }

  private async cachedResult(
    workspace: string,
    manifestPath: string,
    hash: string,
    available: AvailableBackend,
    attempts: PptImageAttempt[],
  ): Promise<PptImageResult | undefined> {
    try {
      const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as RenderManifest
      if (manifest.version !== 2 || manifest.pptx_sha256 !== hash || manifest.backend !== available.backend || manifest.backend_version !== available.version || manifest.capture_method !== available.captureMethod) return undefined
      for (const path of [...manifest.image_paths, ...manifest.contact_sheet_paths]) await resolveWorkspacePath(workspace, path, { mustExist: true, kind: 'file' })
      attempts.push({ backend: available.backend, capture_method: available.captureMethod, status: 'passed', message: 'reused complete render cache' })
      return {
        status: 'passed', backend: manifest.backend, backend_version: manifest.backend_version, capture_method: manifest.capture_method,
        page_count: manifest.page_count, image_paths: manifest.image_paths,
        contact_sheet_paths: manifest.contact_sheet_paths, manifest_path: workspaceRelative(workspace, manifestPath),
        cached: true, attempts, warnings: manifest.warnings,
      }
    } catch {
      return undefined
    }
  }

  private async renderBackend(
    available: AvailableBackend,
    workspace: string,
    pptxPath: string,
    temporary: string,
    rawDirectory: string,
    nativeAutomationApproved: boolean,
    screenIndexInput?: number,
    signal?: AbortSignal,
  ): Promise<void> {
    if (available.backend === 'keynote') {
      const script = join(temporary, 'render.applescript')
      await atomicWriteText(script, KEYNOTE_SCRIPT, { signal })
      const output = join(rawDirectory, 'export')
      const argv = [available.executables.osascript!, script, pptxPath, output]
      const command = nativeAutomationApproved ? argv : (await this.sandbox!.confine(argv, { mode: 'workspace-write', workspaceRoot: workspace })).argv
      const result = await runCollected(this.subprocess!, command, { cwd: dirname(pptxPath), signal, timeoutMs: 90_000, maxOutputBytes: 32_768 })
      if (result.exitCode !== 0) throw new PptError('PPT_RENDER_FAILED', processFailure('Keynote export', result.exitCode, result.stdout, result.stderr))
      return
    }
    if (available.backend === 'powerpoint') {
      if (this.platform === 'darwin') {
        const screenIndex = screenIndexInput ?? 1
        boundedInteger(screenIndex, 'screen_index', 1, 16)
        const script = join(temporary, 'render-powerpoint-screen.applescript')
        const sourceCopy = join(temporary, 'source.pptx')
        await copyFile(pptxPath, sourceCopy)
        await atomicWriteText(script, POWERPOINT_MAC_SCREEN_SCRIPT, { signal })
        const argv = [
          available.executables.osascript!, script, sourceCopy, rawDirectory,
          String(pptxPageCount(new Uint8Array(await readFile(sourceCopy)))),
          available.executables.screencapture!, String(screenIndex),
        ]
        const command = nativeAutomationApproved ? argv : (await this.sandbox!.confine(argv, { mode: 'workspace-write', workspaceRoot: workspace })).argv
        const result = await runCollected(this.subprocess!, command, { cwd: dirname(pptxPath), signal, timeoutMs: 180_000, maxOutputBytes: 32_768 })
        if (result.exitCode !== 0) throw new PptError('PPT_RENDER_FAILED', processFailure('PowerPoint screen capture', result.exitCode, result.stdout, result.stderr))
        return
      }
      const script = join(temporary, 'render.ps1')
      const output = join(rawDirectory, 'export')
      await mkdir(output, { recursive: true })
      await atomicWriteText(script, POWERPOINT_SCRIPT, { signal })
      const argv = [
        available.executables.powershell!, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', script, '-InputPptx', pptxPath, '-OutputDir', output,
      ]
      const command = nativeAutomationApproved ? argv : (await this.sandbox!.confine(argv, { mode: 'workspace-write', workspaceRoot: workspace })).argv
      const result = await runCollected(this.subprocess!, command, { cwd: dirname(pptxPath), signal, timeoutMs: 90_000, maxOutputBytes: 32_768 })
      if (result.exitCode !== 0) throw new PptError('PPT_RENDER_FAILED', processFailure('PowerPoint export', result.exitCode, result.stdout, result.stderr))
      return
    }
    const profile = join(temporary, 'profile')
    const fontCache = join(temporary, 'font-cache')
    const fontconfig = join(temporary, 'fonts.conf')
    const pdfDirectory = join(temporary, 'pdf')
    await Promise.all([mkdir(profile, { recursive: true }), mkdir(fontCache, { recursive: true }), mkdir(pdfDirectory, { recursive: true })])
    const renderFontDirs = [...systemFontDirectories(this.platform), ...this.fontDirs]
    await atomicWriteText(fontconfig, fontconfigDocument(renderFontDirs, fontCache), { signal })
    const renderEnv = {
      FONTCONFIG_FILE: fontconfig,
      FONTCONFIG_PATH: temporary,
      XDG_CACHE_HOME: fontCache,
      SAL_PRIVATE_FONTPATH: renderFontDirs.join(delimiter),
    }
    const convert = await this.sandbox!.confine([
      available.executables.soffice!, '--headless', '--nologo', '--nodefault', '--nolockcheck', '--norestore',
      `-env:UserInstallation=${pathToFileURL(profile).href}`, '--convert-to', 'pdf', '--outdir', pdfDirectory, pptxPath,
    ], { mode: 'workspace-write', workspaceRoot: workspace })
    if (convert.enforcement !== 'full') throw new PptError('PPT_RENDER_FAILED', 'LibreOffice sandbox enforcement is partial')
    const converted = await runCollected(this.subprocess!, convert.argv, { cwd: dirname(pptxPath), signal, timeoutMs: 60_000, maxOutputBytes: 32_768, env: renderEnv })
    if (converted.exitCode !== 0) throw new PptError('PPT_RENDER_FAILED', processFailure('LibreOffice', converted.exitCode, converted.stdout, converted.stderr))
    const pdf = join(pdfDirectory, `${basename(pptxPath, '.pptx')}.pdf`)
    await access(pdf)
    const raster = await this.sandbox!.confine([
      available.executables.pdftoppm!, '-png', '-r', '96', '-scale-to-x', '1280', '-scale-to-y', '720', '-cropbox', pdf, join(rawDirectory, 'page'),
    ], { mode: 'workspace-write', workspaceRoot: workspace })
    const rasterized = await runCollected(this.subprocess!, raster.argv, { cwd: dirname(pptxPath), signal, timeoutMs: 60_000, maxOutputBytes: 32_768 })
    if (rasterized.exitCode !== 0) throw new PptError('PPT_RENDER_FAILED', processFailure('pdftoppm', rasterized.exitCode, rasterized.stdout, rasterized.stderr))
  }
}
