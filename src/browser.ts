import { mkdir } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import type { Browser, BrowserContext, Page } from 'playwright-core'
import { chromium } from 'playwright-core'
import { discoverBrowserExecutable } from './browser-discovery.ts'
import { validatePublicHttpUrl } from './browser-security.ts'
import { DEFAULT_LIMITS } from './limits.ts'
import { PptError, throwIfAborted } from './errors.ts'
import { isLocalFilesystemPath, isPathInside, resolveWorkspacePath, workspaceRelative } from './paths.ts'
import type { SessionOwner } from './session-resources.ts'
import { SessionResourceRegistry } from './session-resources.ts'
import type { DeckIR, ElementIR } from './ir.ts'

export interface BrowserElementRef {
  ref: string
  tag: string
  text: string
  href?: string
  clickable: boolean
}

export interface BrowserPageResult {
  url: string
  title: string
  text: string
  page_version: number
  content_is_untrusted: true
  elements?: BrowserElementRef[]
}

export interface HtmlPreviewResult {
  previews: string[]
  fonts: string[]
  warnings: string[]
  designPages: Array<{
    page: number
    anchorAreaRatio?: number
    frameCount: number
    occupancy: number[]
    roleStyles: Array<{ role: string; fontFamily: string; fontWeight: number }>
  }>
}

interface BrowserState {
  owner: SessionOwner
  workspace: string
  context: BrowserContext
  page: Page
  version: number
  observedMutation: number
  refs: Map<string, { version: number; selector: string; clickable: boolean }>
  refSequence: number
  scrollCount: number
}

function keyOf(owner: SessionOwner): string {
  return `${owner.agentId}\0${owner.sessionId}`
}

/** A leaf rectangle expressed relative to the slide origin. */
export interface SlideLeafBox {
  left: number
  top: number
  right: number
  bottom: number
}

export interface SlideSize {
  width: number
  height: number
}

const SLIDE_SIZE: SlideSize = { width: 1280, height: 720 }
/** Only a leaf covering the slide on all four sides counts as a full-bleed layer. */
const FULL_BLEED_TOLERANCE_PX = 1

function coversWholeSlide(box: SlideLeafBox, slide: SlideSize): boolean {
  return box.left <= FULL_BLEED_TOLERANCE_PX && box.top <= FULL_BLEED_TOLERANCE_PX
    && box.right >= slide.width - FULL_BLEED_TOLERANCE_PX && box.bottom >= slide.height - FULL_BLEED_TOLERANCE_PX
}

/**
 * 4x4 occupancy silhouette of one slide, used to detect adjacent pages with the same layout.
 * Full-bleed layers are skipped: a page background leaf is identical on every page and would
 * light all sixteen cells everywhere, so every adjacent pair would look repeated.
 */
export function occupancySilhouette(leafBoxes: readonly SlideLeafBox[], slide: SlideSize = SLIDE_SIZE): number[] {
  const occupancy = Array.from({ length: 16 }, () => 0)
  const cellWidth = slide.width / 4
  const cellHeight = slide.height / 4
  for (const box of leafBoxes) {
    if (coversWholeSlide(box, slide)) continue
    for (let row = 0; row < 4; row += 1) for (let column = 0; column < 4; column += 1) {
      const left = column * cellWidth
      const top = row * cellHeight
      if (box.right > left && box.left < left + cellWidth && box.bottom > top && box.top < top + cellHeight) occupancy[row * 4 + column] = 1
    }
  }
  return occupancy
}

export class BrowserRuntime {
  private browser?: Browser
  private launching?: Promise<Browser>
  private readonly states = new Map<string, BrowserState>()

  constructor(
    private readonly resources: SessionResourceRegistry,
    private readonly configuredExecutable?: string,
    private readonly outputRoot = 'ppt-output',
  ) {}

  async visit(owner: SessionOwner, workspace: string, input: string, signal?: AbortSignal): Promise<BrowserPageResult> {
    await this.abortOwnerIfRequested(owner, signal)
    const state = await this.requireState(owner, workspace, signal)
    const url = await this.resolveVisitUrl(workspace, input)
    const response = await this.cancellable(owner, state.page.goto(url.href, { waitUntil: 'domcontentloaded', timeout: 30_000 }), signal)
    throwIfAborted(signal)
    if (response !== null) {
      let redirects = 0
      let request = response.request().redirectedFrom()
      while (request !== null) { redirects += 1; request = request.redirectedFrom() }
      if (redirects > DEFAULT_LIMITS.maxRedirects) throw new PptError('BROWSER_LIMIT_EXCEEDED', 'redirect limit exceeded')
      const length = Number(response.headers()['content-length'] ?? 0)
      if (Number.isFinite(length) && length > DEFAULT_LIMITS.maxResponseBytes) {
        throw new PptError('BROWSER_LIMIT_EXCEEDED', `response exceeds ${DEFAULT_LIMITS.maxResponseBytes} bytes`)
      }
    }
    await this.validateCurrentUrl(state)
    await this.bumpVersion(state)
    return this.pageResult(state)
  }

  async find(owner: SessionOwner, query: string, signal?: AbortSignal): Promise<BrowserPageResult> {
    await this.abortOwnerIfRequested(owner, signal)
    const state = this.requireExisting(owner)
    if (query.trim().length === 0 || query.length > 200) throw new PptError('BROWSER_LIMIT_EXCEEDED', 'find query must contain 1..200 characters')
    await this.refreshMutationVersion(state)
    const refs = await state.page.evaluate(({ query, version, start }) => {
      const normalized = query.toLocaleLowerCase()
      const candidates = [...document.querySelectorAll('a,button,[role="button"],[role="link"],summary,h1,h2,h3,p,li')]
      const visible = (element: Element) => {
        const style = getComputedStyle(element)
        const box = element.getBoundingClientRect()
        return style.visibility !== 'hidden' && style.display !== 'none' && box.width > 0 && box.height > 0
      }
      const output: Array<{ ref: string; tag: string; text: string; href?: string; clickable: boolean }> = []
      let sequence = start
      for (const element of candidates) {
        const text = (element.textContent ?? '').replace(/\s+/g, ' ').trim()
        if (!visible(element) || !text.toLocaleLowerCase().includes(normalized)) continue
        const ref = `v${version}-e${sequence++}`
        element.setAttribute('data-dsh-ppt-ref', ref)
        const anchor = element instanceof HTMLAnchorElement ? element : undefined
        output.push({
          ref,
          tag: element.tagName.toLocaleLowerCase(),
          text: text.slice(0, 240),
          ...(anchor?.href === undefined ? {} : { href: anchor.href }),
          clickable: anchor !== undefined || element instanceof HTMLButtonElement || element.getAttribute('role') === 'button' || element.getAttribute('role') === 'link' || element.tagName === 'SUMMARY',
        })
        if (output.length >= 20) break
      }
      return { output, next: sequence, mutation: Number((window as unknown as { __dshPptMutationVersion?: number }).__dshPptMutationVersion ?? 0) }
    }, { query: query.trim(), version: state.version, start: state.refSequence })
    state.refSequence = refs.next
    state.observedMutation = refs.mutation
    state.refs.clear()
    for (const item of refs.output) state.refs.set(item.ref, { version: state.version, selector: `[data-dsh-ppt-ref="${item.ref}"]`, clickable: item.clickable })
    return this.pageResult(state, refs.output)
  }

  async click(owner: SessionOwner, ref: string, signal?: AbortSignal): Promise<BrowserPageResult> {
    await this.abortOwnerIfRequested(owner, signal)
    const state = this.requireExisting(owner)
    await this.refreshMutationVersion(state)
    const target = state.refs.get(ref)
    if (target === undefined || target.version !== state.version) throw new PptError('BROWSER_REF_STALE', `element reference is stale: ${ref}`)
    if (!target.clickable) throw new PptError('BROWSER_REF_STALE', `element is not clickable: ${ref}`)
    const locator = state.page.locator(target.selector)
    if (await locator.count() !== 1 || !(await locator.isVisible())) throw new PptError('BROWSER_REF_STALE', `element reference no longer resolves: ${ref}`)
    await this.cancellable(owner, locator.click({ timeout: 10_000 }), signal)
    await state.page.waitForTimeout(250)
    throwIfAborted(signal)
    await this.validateCurrentUrl(state)
    await this.bumpVersion(state)
    return this.pageResult(state)
  }

  async scroll(owner: SessionOwner, direction: 'up' | 'down', amount = 640, signal?: AbortSignal): Promise<BrowserPageResult> {
    await this.abortOwnerIfRequested(owner, signal)
    const state = this.requireExisting(owner)
    if (!Number.isInteger(amount) || amount < 100 || amount > 2_000) throw new PptError('BROWSER_LIMIT_EXCEEDED', 'scroll amount must be an integer between 100 and 2000')
    state.scrollCount += 1
    if (state.scrollCount > 100) throw new PptError('BROWSER_LIMIT_EXCEEDED', 'scroll call limit exceeded')
    await state.page.evaluate(delta => window.scrollBy({ top: delta, behavior: 'instant' }), direction === 'down' ? amount : -amount)
    await state.page.waitForTimeout(150)
    await this.bumpVersion(state)
    return this.pageResult(state)
  }

  async renderHtmlPreview(
    owner: SessionOwner,
    workspace: string,
    htmlPath: string,
    previewDirectory: string,
    pageCount: number,
    allowedFonts: readonly string[],
    signal?: AbortSignal,
  ): Promise<HtmlPreviewResult> {
    await this.visit(owner, workspace, htmlPath, signal)
    const state = this.requireExisting(owner)
    const preview = await resolveWorkspacePath(workspace, previewDirectory)
    await mkdir(preview, { recursive: true })
    await state.page.addStyleTag({ content: '*,*::before,*::after{animation:none!important;transition:none!important;caret-color:transparent!important}' })
    await state.page.evaluate(async () => { await document.fonts.ready })
    const inspection = await state.page.evaluate(({ count, fonts }) => {
      const slides = [...document.querySelectorAll<HTMLElement>('.ppt-slide[data-page]')]
      const errors: string[] = []
      const warnings: string[] = []
      const usedFonts = new Set<string>()
      if (slides.length !== count) errors.push(`expected ${count} slides but found ${slides.length}`)
      const designPages: Array<{ page: number; anchorAreaRatio?: number; frameCount: number; leafBoxes: SlideLeafBox[]; roleStyles: Array<{ role: string; fontFamily: string; fontWeight: number }> }> = []
      slides.forEach((slide, slideIndex) => {
        const slideBox = slide.getBoundingClientRect()
        const leafBoxes: SlideLeafBox[] = []
        if (Math.abs(slideBox.width - 1280) > 0.5 || Math.abs(slideBox.height - 720) > 0.5) {
          errors.push(`page ${slideIndex + 1}: slide box is ${slideBox.width}x${slideBox.height}, expected 1280x720`)
        }
        for (const leaf of slide.querySelectorAll<HTMLElement>('[data-ppt-id][data-ppt-kind]')) {
          const box = leaf.getBoundingClientRect()
          leafBoxes.push({
            left: box.left - slideBox.left, top: box.top - slideBox.top,
            right: box.right - slideBox.left, bottom: box.bottom - slideBox.top,
          })
          if (box.left < slideBox.left - 1 || box.top < slideBox.top - 1 || box.right > slideBox.right + 1 || box.bottom > slideBox.bottom + 1) {
            errors.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} exceeds slide bounds`)
          }
          const style = getComputedStyle(leaf)
          const family = style.fontFamily.split(',')[0]!.trim().replace(/^['"]|['"]$/g, '')
          if (family.length > 0) {
            usedFonts.add(family)
            if (!fonts.includes(family)) errors.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} uses unauthorized font ${family}`)
          }
          if (leaf.dataset.pptKind === 'text' && (leaf.scrollWidth > leaf.clientWidth + 1 || leaf.scrollHeight > leaf.clientHeight + 1)) {
            errors.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} has text overflow`)
          }
          if (box.width < 1 || box.height < 1) warnings.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} has near-zero bounds`)
        }
        const anchor = slide.querySelector<HTMLElement>('[data-art-role="visual-anchor"]')?.getBoundingClientRect()
        designPages.push({
          page: slideIndex + 1,
          ...(anchor === undefined ? {} : { anchorAreaRatio: Math.max(0, anchor.width * anchor.height) / Math.max(1, slideBox.width * slideBox.height) }),
          frameCount: slide.querySelectorAll('[data-art-role="frame"]').length,
          leafBoxes,
          roleStyles: [...slide.querySelectorAll<HTMLElement>('[data-art-role]')].map(element => {
            const style = getComputedStyle(element)
            return {
              role: element.dataset.artRole ?? '',
              fontFamily: style.fontFamily.split(',')[0]!.trim().replace(/^['"]|['"]$/g, ''),
              fontWeight: Number.parseInt(style.fontWeight, 10) || 400,
            }
          }),
        })
      })
      return { errors, warnings, fonts: [...usedFonts].sort(), designPages }
    }, { count: pageCount, fonts: [...allowedFonts] })
    if (inspection.errors.length > 0) {
      throw new PptError('HTML_CREATE_VALIDATION_FAILED', 'HTML browser validation failed', { details: { issues: inspection.errors } })
    }
    const designPages: HtmlPreviewResult['designPages'] = inspection.designPages.map(page => ({
      page: page.page,
      ...(page.anchorAreaRatio === undefined ? {} : { anchorAreaRatio: page.anchorAreaRatio }),
      frameCount: page.frameCount,
      occupancy: occupancySilhouette(page.leafBoxes),
      roleStyles: page.roleStyles,
    }))
    const previews: string[] = []
    for (let index = 0; index < pageCount; index += 1) {
      throwIfAborted(signal)
      const target = join(preview, `page-${String(index + 1).padStart(3, '0')}.png`)
      await this.cancellable(owner, state.page.locator('.ppt-slide[data-page]').nth(index).screenshot({ path: target, type: 'png', animations: 'disabled' }), signal)
      previews.push(workspaceRelative(workspace, target))
    }
    return { previews, fonts: inspection.fonts, warnings: inspection.warnings, designPages }
  }

  async extractDeckIr(owner: SessionOwner, workspace: string, htmlPath: string, pageCount: number, signal?: AbortSignal): Promise<DeckIR> {
    await this.visit(owner, workspace, htmlPath, signal)
    const state = this.requireExisting(owner)
    const slides = await state.page.evaluate((expected) => {
      const px = (value: string, fallback = 0) => {
        const parsed = Number.parseFloat(value)
        return Number.isFinite(parsed) ? parsed : fallback
      }
      const textRuns = (element: HTMLElement) => {
        const runs: Array<{ text: string; fontFamily: string; fontSizePx: number; fontWeight: number; fontStyle: string; color: string; textDecoration: string }> = []
        const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
        for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
          const value = node.textContent ?? ''
          if (value.length === 0) continue
          const parent = node.parentElement ?? element
          const style = getComputedStyle(parent)
          runs.push({
            text: value, fontFamily: style.fontFamily.split(',')[0]!.trim().replace(/^['"]|['"]$/g, ''),
            fontSizePx: px(style.fontSize), fontWeight: Number.parseInt(style.fontWeight, 10) || 400,
            fontStyle: style.fontStyle, color: style.color, textDecoration: style.textDecorationLine,
          })
        }
        return runs
      }
      const pages = [...document.querySelectorAll<HTMLElement>('.ppt-slide[data-page]')]
      if (pages.length !== expected) throw new Error(`expected ${expected} slides, found ${pages.length}`)
      return pages.map((slide, slideIndex) => {
        const slideBox = slide.getBoundingClientRect()
        const elements = [...slide.querySelectorAll<HTMLElement>('[data-ppt-id][data-ppt-kind]')].map((element, domOrder) => {
          const box = element.getBoundingClientRect()
          const style = getComputedStyle(element)
          const kind = element.dataset.pptKind as ElementIR['kind']
          const item: ElementIR = {
            id: element.dataset.pptId!, kind, z: Number.parseInt(element.dataset.pptZ ?? '0', 10), domOrder,
            box: { x: box.left - slideBox.left, y: box.top - slideBox.top, w: box.width, h: box.height },
            style: {
              fontFamily: style.fontFamily.split(',')[0]!.trim().replace(/^['"]|['"]$/g, ''),
              fontSizePx: px(style.fontSize), fontWeight: Number.parseInt(style.fontWeight, 10) || 400,
              fontStyle: style.fontStyle, color: style.color, backgroundColor: style.backgroundColor,
              backgroundImage: style.backgroundImage, borderColor: style.borderTopColor,
              borderWidthPx: px(style.borderTopWidth), borderStyle: style.borderTopStyle,
              borderRadius: style.borderRadius, textAlign: style.textAlign, verticalAlign: style.verticalAlign,
              lineHeightPx: style.lineHeight === 'normal' ? px(style.fontSize) * 1.2 : px(style.lineHeight),
              opacity: px(style.opacity, 1), objectFit: style.objectFit, objectPosition: style.objectPosition,
            },
          }
          if (kind === 'text') {
            item.text = element.innerText
            item.runs = textRuns(element)
            if (style.backgroundImage !== 'none') item.unsupportedReason = 'gradient or background image on text'
          } else if (kind === 'image') {
            const image = element as HTMLImageElement
            item.imagePath = image.currentSrc || image.src
            if (!image.complete || image.naturalWidth === 0 || image.naturalHeight === 0) item.unsupportedReason = 'image failed to load'
          } else if (kind === 'svg') {
            item.svg = element.outerHTML
            if (element.querySelector('filter,mask,pattern,foreignObject,script') !== null) item.unsupportedReason = 'complex SVG feature'
          } else if (kind === 'table') {
            const table = element as HTMLTableElement
            item.table = [...table.rows].map(row => [...row.cells].map(cell => cell.innerText))
            if ([...table.rows].some(row => [...row.cells].some(cell => cell.rowSpan !== 1 || cell.colSpan !== 1))) item.unsupportedReason = 'merged table cells'
          } else if (kind === 'shape' && style.backgroundImage !== 'none') {
            item.unsupportedReason = 'gradient or background image on shape'
          }
          return item
        }).sort((a, b) => a.z - b.z || a.domOrder - b.domOrder)
        return { page: slideIndex + 1, elements, speakerNotes: [] }
      })
    }, pageCount)
    return { widthPx: 1280, heightPx: 720, slides }
  }

  async rasterizeElement(
    owner: SessionOwner,
    workspace: string,
    htmlPath: string,
    elementId: string,
    targetInput: string,
    signal?: AbortSignal,
  ): Promise<string> {
    await this.visit(owner, workspace, htmlPath, signal)
    const state = this.requireExisting(owner)
    const target = await resolveWorkspacePath(workspace, targetInput, { createParent: true })
    const locator = state.page.locator(`[data-ppt-id="${elementId}"]`)
    if (await locator.count() !== 1) throw new PptError('PPT_CREATE_UNSUPPORTED_ELEMENT', `cannot rasterize missing or duplicate element: ${elementId}`)
    await this.cancellable(owner, locator.screenshot({ path: target, type: 'png', animations: 'disabled' }), signal)
    return workspaceRelative(workspace, target)
  }

  async dispose(): Promise<void> {
    const states = [...this.states.values()]
    this.states.clear()
    await Promise.allSettled(states.map(state => state.context.close()))
    const browser = this.browser
    this.browser = undefined
    if (browser !== undefined) await browser.close()
  }

  private async launch(signal?: AbortSignal): Promise<Browser> {
    if (this.browser !== undefined) return this.browser
    if (this.launching !== undefined) return this.launching
    this.launching = (async () => {
      const found = await discoverBrowserExecutable(this.configuredExecutable)
      if (found.executable === undefined) throw new PptError('BROWSER_NOT_READY', 'no compatible Chromium/Chrome executable found')
      throwIfAborted(signal)
      const browser = await chromium.launch({ executablePath: found.executable, headless: true, args: ['--disable-dev-shm-usage'] })
      this.browser = browser
      return browser
    })().finally(() => { this.launching = undefined })
    return this.launching
  }

  private async requireState(owner: SessionOwner, workspace: string, signal?: AbortSignal): Promise<BrowserState> {
    const current = this.states.get(keyOf(owner))
    if (current !== undefined) {
      if (current.workspace !== workspace) throw new PptError('BROWSER_URL_BLOCKED', 'browser workspace cannot change within a session')
      return current
    }
    this.resources.open(owner, workspace)
    const browser = await this.launch(signal)
    const context = await browser.newContext({ acceptDownloads: false, serviceWorkers: 'block', viewport: { width: 1280, height: 720 } })
    await context.addInitScript(() => {
      const target = window as unknown as { __dshPptMutationVersion?: number }
      target.__dshPptMutationVersion = 0
      new MutationObserver((records) => {
        if (records.some(record => record.type !== 'attributes' || record.attributeName !== 'data-dsh-ppt-ref')) {
          target.__dshPptMutationVersion = (target.__dshPptMutationVersion ?? 0) + 1
        }
      })
        .observe(document, { childList: true, subtree: true, attributes: true, characterData: true })
    })
    await context.route('**/*', async (route) => {
      const requestUrl = new URL(route.request().url())
      try {
        if (requestUrl.protocol === 'file:') {
          const path = fileURLToPath(requestUrl)
          const outputRoot = await resolveWorkspacePath(workspace, this.outputRoot)
          if (!isPathInside(outputRoot, path)) throw new PptError('BROWSER_URL_BLOCKED', 'local browser path is outside plugin output')
        } else {
          await validatePublicHttpUrl(requestUrl.href)
        }
        await route.continue()
      } catch {
        await route.abort('blockedbyclient')
      }
    })
    const page = await context.newPage()
    page.on('dialog', dialog => { void dialog.dismiss() })
    page.on('download', download => { void download.cancel() })
    page.on('popup', popup => { void popup.close() })
    const state: BrowserState = {
      owner: { ...owner }, workspace, context, page, version: 0, observedMutation: 0,
      refs: new Map(), refSequence: 1, scrollCount: 0,
    }
    this.states.set(keyOf(owner), state)
    this.resources.track(owner, {
      label: 'browser-context',
      dispose: async () => {
        this.states.delete(keyOf(owner))
        await context.close().catch(() => undefined)
      },
    })
    return state
  }

  private requireExisting(owner: SessionOwner): BrowserState {
    const state = this.states.get(keyOf(owner))
    if (state === undefined) throw new PptError('BROWSER_NOT_READY', 'call browser_visit before using this browser tool')
    return state
  }

  private async resolveVisitUrl(workspace: string, input: string): Promise<URL> {
    if (isLocalFilesystemPath(input)) return this.resolveLocalFileUrl(workspace, input)
    try {
      const parsed = new URL(input)
      if (parsed.protocol === 'file:') return await this.resolveLocalFileUrl(workspace, fileURLToPath(parsed))
      return await validatePublicHttpUrl(parsed.href)
    } catch (error) {
      if (error instanceof PptError) throw error
      return this.resolveLocalFileUrl(workspace, input)
    }
  }

  private async resolveLocalFileUrl(workspace: string, input: string): Promise<URL> {
    const path = await resolveWorkspacePath(workspace, input, { mustExist: true, kind: 'file' })
    const outputRoot = await resolveWorkspacePath(workspace, this.outputRoot)
    if (!isPathInside(outputRoot, path)) throw new PptError('BROWSER_URL_BLOCKED', 'local HTML is outside plugin output')
    return pathToFileURL(path)
  }

  private async validateCurrentUrl(state: BrowserState): Promise<void> {
    const current = new URL(state.page.url())
    if (current.protocol === 'file:') {
      const outputRoot = await resolveWorkspacePath(state.workspace, this.outputRoot)
      if (!isPathInside(outputRoot, fileURLToPath(current))) throw new PptError('BROWSER_URL_BLOCKED', 'navigation left plugin output')
    } else {
      await validatePublicHttpUrl(current.href)
    }
  }

  private async bumpVersion(state: BrowserState): Promise<void> {
    state.version += 1
    state.refs.clear()
    state.observedMutation = await state.page.evaluate(() => Number((window as unknown as { __dshPptMutationVersion?: number }).__dshPptMutationVersion ?? 0))
  }

  private async refreshMutationVersion(state: BrowserState): Promise<void> {
    const mutation = await state.page.evaluate(() => Number((window as unknown as { __dshPptMutationVersion?: number }).__dshPptMutationVersion ?? 0))
    if (mutation !== state.observedMutation) {
      state.version += 1
      state.refs.clear()
      state.observedMutation = mutation
    }
  }

  private async pageResult(state: BrowserState, elements?: BrowserElementRef[]): Promise<BrowserPageResult> {
    const [title, text] = await Promise.all([
      state.page.title(),
      state.page.locator('body').innerText({ timeout: 5_000 }).catch(() => ''),
    ])
    return {
      url: state.page.url(),
      title: title.slice(0, 500),
      text: text.replace(/\u0000/g, '').slice(0, DEFAULT_LIMITS.maxBrowserTextChars),
      page_version: state.version,
      content_is_untrusted: true,
      ...(elements === undefined ? {} : { elements }),
    }
  }

  private async cancellable<T>(owner: SessionOwner, operation: Promise<T>, signal?: AbortSignal): Promise<T> {
    throwIfAborted(signal)
    if (signal === undefined) return operation
    return new Promise<T>((resolve, reject) => {
      let settled = false
      const finish = (callback: () => void) => {
        if (settled) return
        settled = true
        signal.removeEventListener('abort', abort)
        callback()
      }
      const abort = () => {
        if (settled) return
        settled = true
        signal.removeEventListener('abort', abort)
        void this.resources.release(owner).then(
          () => reject(new PptError('PPT_ABORTED', 'browser operation aborted and its isolated page was released')),
          error => reject(new PptError('PPT_ABORTED', 'browser operation aborted but page cleanup failed', { cause: error })),
        )
      }
      signal.addEventListener('abort', abort, { once: true })
      operation.then(
        value => finish(() => resolve(value)),
        error => finish(() => reject(error)),
      )
    })
  }

  private async abortOwnerIfRequested(owner: SessionOwner, signal?: AbortSignal): Promise<void> {
    if (!signal?.aborted) return
    await this.resources.release(owner)
    throwIfAborted(signal)
  }
}
