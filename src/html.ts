import { randomUUID } from 'node:crypto'
import { access, readFile, rm } from 'node:fs/promises'
import { dirname, isAbsolute, join } from 'node:path'
import { atomicWriteJson, atomicWriteText } from './atomic.ts'
import { ART_ROLES, artDirectionFindings, type ArtDirection, type DesignFinding, validateArtDirection } from './art-direction.ts'
import type { BrowserRuntime } from './browser.ts'
import { PptError, throwIfAborted } from './errors.ts'
import { discoverRegisteredFonts, FONT_REGISTRY } from './fonts.ts'
import { validatePptOutline } from './outline.ts'
import { isPathInside, resolveWorkspacePath, workspaceRelative } from './paths.ts'
import type { SessionOwner } from './session-resources.ts'

const LEAF_KINDS = new Set(['text', 'image', 'shape', 'svg', 'table'])
/** `NodeFilter.SHOW_TEXT`, inlined so the walker call does not depend on the DOM library's globals. */
const SHOW_TEXT = 4
const BLOCKED_ELEMENTS = new Set(['SCRIPT', 'IFRAME', 'OBJECT', 'EMBED', 'FORM', 'INPUT', 'TEXTAREA', 'SELECT', 'VIDEO', 'AUDIO', 'CANVAS', 'FOREIGNOBJECT'])
const ALLOWED_CSS_PREFIXES = [
  'width', 'height', 'min-', 'max-', 'box-sizing', 'position', 'top', 'right', 'bottom', 'left', 'display',
  'flex', 'grid', 'gap', 'row-gap', 'column-gap', 'align-', 'justify-', 'place-', 'order', 'margin', 'padding',
  'font', 'line-height', 'letter-spacing', 'text-', 'white-space', 'word-break', 'overflow', 'color', 'background',
  'border', 'border-radius', 'opacity', 'z-index', 'object-fit', 'object-position', 'list-style', 'vertical-align',
]
const BLOCKED_CSS = /(?:^|[;{])\s*(?:transform|filter|animation|transition|clip-path|mask|mix-blend-mode|perspective|backdrop-filter|box-shadow)\s*:/imu
const REMOTE_RESOURCE = /(?:url\s*\(\s*['"]?\s*(?:https?:|data:|javascript:|file:)|@import\b)/iu

export interface HtmlCreateResult {
  html_path: string
  page_count: number
  preview_paths: string[]
  fonts: string[]
  external_resources: 'none'
  warnings: string[]
  unsupported_css: string[]
  design_status: 'directed' | 'legacy'
  design_findings: DesignFinding[]
  design_validation_path: string
}

function cssProperties(css: string): string[] {
  const result = new Set<string>()
  for (const match of css.matchAll(/(?:^|[;{])\s*([a-z-]+)\s*:/gimu)) result.add(match[1]!.toLowerCase())
  return [...result]
}

function localReference(value: string): boolean {
  if (value.trim().length === 0 || value.includes('\0') || value.includes('\\') || isAbsolute(value)) return false
  try { new URL(value); return false } catch { /* Relative reference. */ }
  return value.split('/').every(segment => segment !== '..' && segment !== '.')
}

export async function validateDeckHtmlSource(
  workspace: string,
  artifactRoot: string,
  html: string,
  outlineLength: number,
  designPlan?: ArtDirection,
  strictDesign = false,
): Promise<{ fonts: string[]; primaryFonts: string[]; unsupported: string[]; designFindings: DesignFinding[] }> {
  if (Buffer.byteLength(html) > 5 * 1024 * 1024) throw new PptError('PPT_RESOURCE_LIMIT', 'HTML source exceeds 5 MiB')
  // Loaded on demand: the DOM implementation is a large module graph and this is
  // the only consumer, so the plugin entry must not pull it during host startup.
  const { Window } = await import('happy-dom')
  const domWindow = new Window()
  domWindow.document.write(html)
  // happy-dom types its element queries by tag-name map and rejects the standard
  // `querySelectorAll<HTMLElement>` generics; the runtime document is a standard DOM.
  const document = domWindow.document as unknown as Document
  const issues: string[] = []
  const unsupported = new Set<string>()
  const allowedFonts = new Set(FONT_REGISTRY.map(font => font.name))
  const usedFonts = new Set<string>()
  const primaryFonts = new Set<string>()
  const designFindings = designPlan === undefined ? [] : artDirectionFindings(designPlan)
  for (const element of document.querySelectorAll('*')) {
    if (BLOCKED_ELEMENTS.has(element.tagName)) issues.push(`blocked element: ${element.tagName.toLowerCase()}`)
    for (const attribute of [...element.attributes]) {
      if (/^on/iu.test(attribute.name)) issues.push(`event handler attribute is blocked: ${attribute.name}`)
      if (['src', 'href', 'xlink:href'].includes(attribute.name) && attribute.value.trim().length > 0) {
        if (!localReference(attribute.value)) issues.push(`non-local resource is blocked: ${attribute.value.slice(0, 120)}`)
      }
    }
  }
  const css = [...document.querySelectorAll('style')].map(style => style.textContent ?? '').join('\n')
    + '\n' + [...document.querySelectorAll<HTMLElement>('[style]')].map(element => element.getAttribute('style') ?? '').join('\n')
  if (REMOTE_RESOURCE.test(css)) issues.push('CSS contains a remote, data, script, or file resource')
  if (BLOCKED_CSS.test(css)) issues.push('CSS contains an unsupported effects or animation property')
  for (const property of cssProperties(css)) {
    if (!ALLOWED_CSS_PREFIXES.some(prefix => prefix.endsWith('-') ? property.startsWith(prefix) : property === prefix || property.startsWith(`${prefix}-`))) unsupported.add(property)
  }
  if (unsupported.size > 0) issues.push(`unsupported CSS properties: ${[...unsupported].sort().join(', ')}`)
  for (const match of css.matchAll(/font-family\s*:\s*([^;}{]+)/gimu)) {
    const families = match[1]!.split(',').map(value => value.trim().replace(/^['"]|['"]$/g, '')).filter(Boolean)
    if (families[0] !== undefined) primaryFonts.add(families[0])
    for (const family of families) {
      usedFonts.add(family)
      if (!allowedFonts.has(family)) issues.push(`unauthorized font: ${family}`)
    }
  }

  const slides = [...document.querySelectorAll<HTMLElement>('.ppt-slide[data-page]')]
  if (slides.length !== outlineLength) issues.push(`expected ${outlineLength} .ppt-slide elements, found ${slides.length}`)
  const ids = new Set<string>()
  slides.forEach((slide, index) => {
    if (slide.dataset.page !== String(index + 1)) issues.push(`slide ${index + 1} has non-contiguous data-page`)
    const planned = designPlan?.slides[index]
    if (planned !== undefined) {
      if (slide.dataset.artComposition !== planned.composition) issues.push(`page ${index + 1} data-art-composition must be ${planned.composition}`)
      if (slide.dataset.artDensity !== planned.density) issues.push(`page ${index + 1} data-art-density must be ${planned.density}`)
      if (slide.dataset.artBackground !== planned.background_role) issues.push(`page ${index + 1} data-art-background must be ${planned.background_role}`)
      const roleElements = [...slide.querySelectorAll<HTMLElement>('[data-art-role]')]
      for (const element of roleElements) {
        if (!(ART_ROLES as readonly string[]).includes(element.dataset.artRole ?? '')) issues.push(`page ${index + 1} has invalid data-art-role: ${element.dataset.artRole ?? ''}`)
      }
      const anchors = roleElements.filter(element => element.dataset.artRole === 'visual-anchor')
      if (planned.visual_anchor.kind === 'none' && anchors.length !== 0) issues.push(`page ${index + 1} declares no visual anchor but HTML contains ${anchors.length}`)
      if (planned.visual_anchor.kind !== 'none' && anchors.length !== 1) issues.push(`page ${index + 1} requires exactly one visual-anchor, found ${anchors.length}`)
      const frames = roleElements.filter(element => element.dataset.artRole === 'frame').length
      if (planned.frame_policy === 'none' && frames > 0) issues.push(`page ${index + 1} frame_policy=none but HTML contains frames`)
      if (planned.frame_policy === 'single' && frames > 1) issues.push(`page ${index + 1} frame_policy=single but HTML contains ${frames} frames`)
      if (planned.frame_policy === 'grouped' && frames < 2) {
        designFindings.push({ code: 'ART_GROUPED_FRAMES_NOT_REALIZED', severity: 'warning', message: 'grouped frame policy is not visibly realized', page: index + 1 })
      }
    }
    const walker = document.createTreeWalker(slide, SHOW_TEXT)
    for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
      if ((node.textContent ?? '').trim().length > 0 && node.parentElement?.closest('[data-ppt-id][data-ppt-kind]') === null) {
        issues.push(`page ${index + 1} contains visible text outside a convertible leaf`)
        break
      }
    }
    for (const leaf of slide.querySelectorAll<HTMLElement>('[data-ppt-id], [data-ppt-kind]')) {
      const id = leaf.dataset.pptId
      const kind = leaf.dataset.pptKind
      if (id === undefined || !/^[A-Za-z][A-Za-z0-9_.-]{0,79}$/u.test(id)) issues.push(`page ${index + 1} has invalid or missing data-ppt-id`)
      else if (ids.has(id)) issues.push(`duplicate data-ppt-id: ${id}`)
      else ids.add(id)
      if (kind === undefined || !LEAF_KINDS.has(kind)) issues.push(`page ${index + 1} has invalid or missing data-ppt-kind`)
      if (leaf.querySelector('[data-ppt-id][data-ppt-kind]') !== null) issues.push(`${id ?? 'unknown'} is not a leaf`)
      if (kind === 'image' && leaf.tagName !== 'IMG') issues.push(`${id ?? 'unknown'} kind=image must be an img`)
      if (kind === 'svg' && leaf.tagName !== 'svg') issues.push(`${id ?? 'unknown'} kind=svg must be an svg`)
      if (kind === 'table' && leaf.tagName !== 'TABLE') issues.push(`${id ?? 'unknown'} kind=table must be a table`)
      const z = leaf.dataset.pptZ
      if (z !== undefined && (!/^-?\d+$/u.test(z) || Math.abs(Number(z)) > 10_000)) issues.push(`${id ?? 'unknown'} has invalid data-ppt-z`)
    }
  })
  if (designPlan !== undefined) {
    const plannedFonts = new Set(Object.values(designPlan.typography).map(role => role.family))
    for (const family of plannedFonts) {
      if (!usedFonts.has(family)) designFindings.push({ code: 'ART_FONT_ROLE_UNUSED', severity: 'warning', message: `planned font family ${family} is not declared in HTML CSS` })
    }
    if (strictDesign) {
      for (const finding of designFindings) if (finding.severity === 'warning') issues.push(`${finding.code}${finding.page === undefined ? '' : ` page ${finding.page}`}: ${finding.message}`)
    }
  }
  const assetReferences = new Set<string>()
  for (const element of document.querySelectorAll<HTMLElement>('img[src], svg image[href], svg image[xlink\\:href]')) {
    assetReferences.add(element.getAttribute('src') ?? element.getAttribute('href') ?? element.getAttribute('xlink:href') ?? '')
  }
  for (const match of css.matchAll(/url\s*\(\s*['"]?([^'")]+)['"]?\s*\)/gimu)) assetReferences.add(match[1]!.trim())
  for (const ref of assetReferences) {
    if (!localReference(ref)) continue
    try {
      const path = await resolveWorkspacePath(workspace, join(artifactRoot, ref), { mustExist: true, kind: 'file' })
      if (!isPathInside(artifactRoot, path)) issues.push(`asset leaves artifact directory: ${ref}`)
    } catch {
      issues.push(`missing or invalid local asset: ${ref}`)
    }
  }
  domWindow.close()
  if (issues.length > 0) throw new PptError('HTML_CREATE_VALIDATION_FAILED', 'HTML static validation failed', { details: { issues } })
  return { fonts: [...usedFonts].sort(), primaryFonts: [...primaryFonts].sort(), unsupported: [...unsupported].sort(), designFindings }
}

export async function createHtmlDeck(
  browser: BrowserRuntime,
  owner: SessionOwner,
  workspace: string,
  outlinePathInput: string,
  html: string,
  signal?: AbortSignal,
  designPlanPathInput?: string,
  strictDesign = false,
  fontDirs?: readonly string[],
): Promise<HtmlCreateResult> {
  throwIfAborted(signal)
  const outlinePath = await resolveWorkspacePath(workspace, outlinePathInput, { mustExist: true, kind: 'file' })
  const artifactRoot = dirname(outlinePath)
  const outline = validatePptOutline(JSON.parse(await readFile(outlinePath, 'utf8')))
  let designPlan: ArtDirection | undefined
  if (designPlanPathInput !== undefined) {
    const designPlanPath = await resolveWorkspacePath(workspace, designPlanPathInput, { mustExist: true, kind: 'file' })
    if (dirname(designPlanPath) !== artifactRoot) throw new PptError('HTML_CREATE_INPUT_INVALID', 'design plan must be in the same artifact directory as outline.json')
    designPlan = validateArtDirection(JSON.parse(await readFile(designPlanPath, 'utf8')), outline.length)
  }
  if (outline.some(slide => slide.content.some(item => item.kind === 'chart' && item.data_ref === undefined && !slide.content.some(other => other.kind === 'data')))) {
    throw new PptError('HTML_CREATE_INPUT_INVALID', 'outline still contains a chart with explicitly pending data')
  }
  const output = join(artifactRoot, 'deck.html')
  try { await access(output); throw new PptError('PPT_OUTPUT_EXISTS', `output already exists: ${workspaceRelative(workspace, output)}`) } catch (error) {
    if (error instanceof PptError) throw error
  }
  const validation = await validateDeckHtmlSource(workspace, artifactRoot, html, outline.length, designPlan, strictDesign)
  if (fontDirs !== undefined) {
    const discovered = await discoverRegisteredFonts(fontDirs)
    const available = new Set(discovered.map(font => font.name))
    const unavailable = validation.primaryFonts.filter(font => !available.has(font))
    if (unavailable.length > 0) {
      throw new PptError('PPT_DEPENDENCY_MISSING', 'HTML declares a primary font that is not installed in the approved registry', {
        details: { unavailable, available: [...available].sort(), scope: 'approved_registry' },
      })
    }
  }
  const temporary = join(artifactRoot, `.deck.${process.pid}.${randomUUID()}.html`)
  const previewDirectory = join(artifactRoot, 'preview')
  try {
    await atomicWriteText(temporary, html, { signal })
    const rendered = await browser.renderHtmlPreview(
      owner, workspace, temporary, previewDirectory, outline.length, FONT_REGISTRY.map(font => font.name), signal,
    )
    if (designPlan !== undefined) {
      for (const page of rendered.designPages) {
        const planned = designPlan.slides[page.page - 1]!
        const typographyRole = (role: string): keyof ArtDirection['typography'] | undefined => {
          if (role === 'title' || role === 'subtitle') return 'display'
          if (role === 'body' || role === 'supporting' || role === 'frame' || role === 'diagram') return 'body'
          if (role === 'metric') return 'latin'
          if (role === 'code') return 'code'
          if (role === 'visual-anchor') {
            if (planned.visual_anchor.kind === 'typography') return 'display'
            if (planned.visual_anchor.kind === 'code') return 'code'
            if (planned.visual_anchor.kind === 'data') return 'latin'
          }
          return undefined
        }
        for (const roleStyle of page.roleStyles) {
          const role = typographyRole(roleStyle.role)
          if (role === undefined) continue
          const expected = designPlan.typography[role]
          if (roleStyle.fontFamily !== expected.family || roleStyle.fontWeight !== expected.weight) {
            validation.designFindings.push({
              code: 'ART_TYPOGRAPHY_ROLE_MISMATCH', severity: 'warning', page: page.page,
              message: `${roleStyle.role} uses ${roleStyle.fontFamily} ${roleStyle.fontWeight}, expected ${expected.family} ${expected.weight}`,
            })
          }
        }
        const minimum = planned.visual_anchor.min_area_ratio
        if (minimum !== undefined && (page.anchorAreaRatio ?? 0) < minimum) {
          validation.designFindings.push({
            code: 'ART_VISUAL_ANCHOR_TOO_SMALL', severity: 'warning', page: page.page,
            message: `visual anchor covers ${((page.anchorAreaRatio ?? 0) * 100).toFixed(1)}% of the slide, below the planned ${(minimum * 100).toFixed(1)}%`,
          })
        }
        if (page.page > 1 && !planned.allow_intentional_repeat) {
          const previous = rendered.designPages[page.page - 2]!
          const intersection = page.occupancy.filter((value, index) => value === 1 && previous.occupancy[index] === 1).length
          const union = page.occupancy.filter((value, index) => value === 1 || previous.occupancy[index] === 1).length
          if (union > 0 && intersection / union > 0.88) {
            validation.designFindings.push({ code: 'ART_SILHOUETTE_REPEATED', severity: 'warning', page: page.page, message: 'adjacent slide occupancy silhouettes are highly similar' })
          }
        }
      }
      if (strictDesign && validation.designFindings.some(finding => finding.severity === 'warning')) {
        throw new PptError('HTML_CREATE_VALIDATION_FAILED', 'HTML design fidelity validation failed in strict mode', {
          details: { issues: validation.designFindings.map(finding => ({ code: finding.code, page: finding.page, message: finding.message })) },
        })
      }
    }
    await atomicWriteText(output, html, { signal })
    const designValidation = join(artifactRoot, 'design-validation.json')
    await atomicWriteJson(designValidation, {
      version: 1,
      mode: designPlan === undefined ? 'legacy' : 'directed',
      checks: designPlan === undefined ? [] : ['html-art-attributes', 'art-roles', 'visual-anchor-area', 'frame-policy', 'occupancy-silhouette', 'font-declarations', 'computed-typography-roles'],
      pages: rendered.designPages,
      findings: validation.designFindings,
    }, { signal })
    return {
      html_path: workspaceRelative(workspace, output), page_count: outline.length,
      preview_paths: rendered.previews, fonts: [...new Set([...validation.fonts, ...rendered.fonts])].sort(),
      external_resources: 'none',
      warnings: [
        ...(designPlan === undefined ? ['ART_DIRECTION_MISSING: HTML validated in legacy design mode'] : []),
        ...validation.designFindings.map(finding => `${finding.code}${finding.page === undefined ? '' : ` (page ${finding.page})`}: ${finding.message}`),
        ...rendered.warnings,
      ],
      unsupported_css: validation.unsupported,
      design_status: designPlan === undefined ? 'legacy' : 'directed',
      design_findings: validation.designFindings,
      design_validation_path: workspaceRelative(workspace, designValidation),
    }
  } finally {
    await rm(temporary, { force: true })
  }
}
