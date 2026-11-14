import { readFile, rm } from 'node:fs/promises'
import { dirname, join, posix } from 'node:path'
import { fileURLToPath } from 'node:url'
import { strFromU8, unzipSync } from 'fflate'
import PptxGenJS from 'pptxgenjs'
import { atomicWriteFile } from './atomic.ts'
import type { BrowserRuntime } from './browser.ts'
import { PptError, throwIfAborted } from './errors.ts'
import type { DeckIR, ElementIR, ElementStyleIR } from './ir.ts'
import { pxToInches, pxToPoints, SLIDE_HEIGHT_IN, SLIDE_WIDTH_IN } from './ir.ts'
import { validatePptOutline } from './outline.ts'
import { isLocalFilesystemPath, isPathInside, resolveWorkspacePath, workspaceRelative } from './paths.ts'
import type { SessionOwner } from './session-resources.ts'
import { validateDeckHtmlSource } from './html.ts'
import { rewritePptxTransitions, type SlideTransitionPlan } from './transitions.ts'
import { rewritePptxAnimations, type AnimationPlan } from './animation.ts'

export type PptFallbackMode = 'reject' | 'rasterize-element'

export interface RasterizedElementRecord {
  page: number
  element_id: string
  reason: string
  image_path: string
}

export interface PptCreateResult {
  pptx_path: string
  page_count: number
  native_element_count: number
  rasterized_elements: RasterizedElementRecord[]
  structural_status: 'passed'
}

export interface PptxPackageInspection {
  pageCount: number
  widthEmu: number
  heightEmu: number
  entries: string[]
}

function color(value: string): { hex: string; transparency: number } | undefined {
  if (value === 'transparent') return undefined
  const match = /rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)(?:\s*[,/]\s*(\d+(?:\.\d+)?))?\s*\)/iu.exec(value)
  if (match === null) return undefined
  const hex = [match[1], match[2], match[3]].map(channel => Math.max(0, Math.min(255, Math.round(Number(channel)))).toString(16).padStart(2, '0')).join('').toUpperCase()
  const alpha = match[4] === undefined ? 1 : Math.max(0, Math.min(1, Number(match[4])))
  return { hex, transparency: Math.round((1 - alpha) * 100) }
}

function baseOptions(element: ElementIR) {
  return {
    x: pxToInches(element.box.x), y: pxToInches(element.box.y),
    w: pxToInches(element.box.w), h: pxToInches(element.box.h), objectName: element.id,
  }
}

function textOptions(style: ElementStyleIR) {
  const foreground = color(style.color)
  const fill = color(style.backgroundColor)
  const line = color(style.borderColor)
  const align = ['left', 'center', 'right', 'justify'].includes(style.textAlign) ? style.textAlign : 'left'
  return {
    fontFace: style.fontFamily, fontSize: pxToPoints(style.fontSizePx), bold: style.fontWeight >= 600,
    italic: style.fontStyle === 'italic', color: foreground?.hex ?? '000000',
    margin: 0, breakLine: false, fit: 'none' as const, align,
    valign: style.verticalAlign === 'bottom' ? 'bottom' : style.verticalAlign === 'middle' ? 'mid' : 'top',
    lineSpacing: pxToPoints(style.lineHeightPx),
    ...(fill === undefined ? {} : { fill: { color: fill.hex, transparency: fill.transparency } }),
    ...(style.borderStyle === 'none' || style.borderWidthPx <= 0 || line === undefined
      ? { line: { type: 'none' as const } }
      : { line: { color: line.hex, transparency: line.transparency, width: pxToPoints(style.borderWidthPx) } }),
  }
}

function imageSource(pathOrUrl: string, workspace: string, artifactRoot: string): string {
  let path: string
  if (isLocalFilesystemPath(pathOrUrl)) {
    // `new URL('E:\\dir\\a.png')` parses successfully on Windows with the drive
    // letter as the scheme, so a bare platform path must never reach URL parsing.
    path = pathOrUrl
  } else {
    let parsed: URL | undefined
    try {
      parsed = new URL(pathOrUrl)
    } catch {
      parsed = undefined
    }
    if (parsed === undefined) path = pathOrUrl
    else if (parsed.protocol === 'file:') path = fileURLToPath(parsed)
    else throw new PptError('PPT_CREATE_ASSET_MISSING', 'PPTX images must be local frozen files')
  }
  if (!isPathInside(workspace, path) || !isPathInside(artifactRoot, path)) {
    throw new PptError('PPT_CREATE_ASSET_MISSING', `image is outside the artifact directory: ${path}`)
  }
  return path
}

function addNativeElement(pptx: PptxGenJS, slide: ReturnType<PptxGenJS['addSlide']>, element: ElementIR, workspace: string, artifactRoot: string): void {
  const base = baseOptions(element)
  if (element.kind === 'text') {
    const runs = (element.runs ?? []).map(run => ({
      text: run.text,
      options: {
        fontFace: run.fontFamily, fontSize: pxToPoints(run.fontSizePx), bold: run.fontWeight >= 600,
        italic: run.fontStyle === 'italic', color: color(run.color)?.hex ?? '000000',
        ...(run.textDecoration.includes('underline') ? { underline: { style: 'sng' as const } } : {}),
      },
    }))
    slide.addText((runs.length > 0 ? runs : element.text ?? '') as never, { ...base, ...textOptions(element.style) } as never)
    return
  }
  if (element.kind === 'image') {
    if (element.imagePath === undefined) throw new PptError('PPT_CREATE_ASSET_MISSING', `image path missing for ${element.id}`)
    const path = imageSource(element.imagePath, workspace, artifactRoot)
    const fit = element.style.objectFit === 'contain' ? 'contain' : 'cover'
    slide.addImage({ path, ...base, sizing: { type: fit, w: base.w, h: base.h }, transparency: Math.round((1 - element.style.opacity) * 100) })
    return
  }
  if (element.kind === 'svg') {
    if (element.svg === undefined) throw new PptError('PPT_CREATE_UNSUPPORTED_ELEMENT', `SVG markup missing for ${element.id}`)
    const data = `data:image/svg+xml;base64,${Buffer.from(element.svg).toString('base64')}`
    slide.addImage({ data, ...base, sizing: { type: 'contain', w: base.w, h: base.h } })
    return
  }
  if (element.kind === 'table') {
    slide.addTable((element.table ?? []).map(row => row.map(cell => ({ text: cell }))), {
      ...base, border: { type: 'solid', color: color(element.style.borderColor)?.hex ?? '999999', pt: pxToPoints(Math.max(1, element.style.borderWidthPx)) },
      color: color(element.style.color)?.hex ?? '000000', fill: color(element.style.backgroundColor)?.hex ?? 'FFFFFF',
      fontFace: element.style.fontFamily, fontSize: pxToPoints(element.style.fontSizePx), margin: 0,
    } as never)
    return
  }
  const fill = color(element.style.backgroundColor)
  const line = color(element.style.borderColor)
  const radius = Number.parseFloat(element.style.borderRadius)
  const shape = element.box.w <= 2 || element.box.h <= 2
    ? pptx.ShapeType.line
    : /%/u.test(element.style.borderRadius) && radius >= 50 ? pptx.ShapeType.ellipse
      : radius > 0 ? pptx.ShapeType.roundRect : pptx.ShapeType.rect
  slide.addShape(shape, {
    ...base,
    fill: fill === undefined ? { type: 'none' } : { color: fill.hex, transparency: fill.transparency },
    line: element.style.borderStyle === 'none' || element.style.borderWidthPx <= 0 || line === undefined
      ? { type: 'none' } : { color: line.hex, transparency: line.transparency, width: pxToPoints(element.style.borderWidthPx) },
  } as never)
}

function relationshipSource(path: string): string {
  if (path === '_rels/.rels') return ''
  const marker = '/_rels/'
  const index = path.indexOf(marker)
  if (index < 0 || !path.endsWith('.rels')) return ''
  return `${path.slice(0, index)}/${path.slice(index + marker.length, -'.rels'.length)}`
}

export function inspectPptxPackage(data: Uint8Array, expectedPages: number): PptxPackageInspection {
  let files: Record<string, Uint8Array>
  try { files = unzipSync(data) } catch (error) {
    throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX is not a readable ZIP package', { cause: error })
  }
  const names = Object.keys(files).sort()
  for (const required of ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels']) {
    if (files[required] === undefined) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `PPTX entry is missing: ${required}`)
  }
  const slideNames = names.filter(name => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).sort((a, b) => Number(a.match(/\d+/u)![0]) - Number(b.match(/\d+/u)![0]))
  if (slideNames.length !== expectedPages) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `expected ${expectedPages} slides, found ${slideNames.length}`)
  slideNames.forEach((name, index) => {
    if (name !== `ppt/slides/slide${index + 1}.xml`) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `slide sequence is not contiguous at ${name}`)
  })
  const presentation = strFromU8(files['ppt/presentation.xml']!)
  const size = /<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/u.exec(presentation)
  if (size === null) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'presentation slide size is missing')
  const widthEmu = Number(size[1]); const heightEmu = Number(size[2])
  if (widthEmu !== 12_192_000 || heightEmu !== 6_858_000) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `unexpected slide size: ${widthEmu}x${heightEmu}`)
  const presentationRels = strFromU8(files['ppt/_rels/presentation.xml.rels']!)
  const relationTargets = new Map<string, string>()
  for (const match of presentationRels.matchAll(/<Relationship\b([^>]*)>/giu)) {
    const id = /\bId="([^"]+)"/iu.exec(match[1]!)?.[1]
    const target = /\bTarget="([^"]+)"/iu.exec(match[1]!)?.[1]
    if (id !== undefined && target !== undefined) relationTargets.set(id, posix.normalize(posix.join('ppt', target)))
  }
  const orderedSlideIds = [...presentation.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"/giu)].map(match => match[1]!)
  if (orderedSlideIds.length !== expectedPages) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'presentation slide order list does not match page count')
  orderedSlideIds.forEach((id, index) => {
    if (relationTargets.get(id) !== `ppt/slides/slide${index + 1}.xml`) {
      throw new PptError('PPT_CREATE_INVALID_PACKAGE', `presentation slide order is invalid at page ${index + 1}`)
    }
  })

  for (const name of names.filter(entry => entry.endsWith('.rels'))) {
    const xml = strFromU8(files[name]!)
    if (/TargetMode="External"/iu.test(xml)) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `external relationship is forbidden: ${name}`)
    const source = relationshipSource(name)
    const base = source === '' ? '' : posix.dirname(source)
    for (const match of xml.matchAll(/<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*>/giu)) {
      const target = match[1]!
      if (/^[a-z]+:/iu.test(target)) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `non-package relationship target: ${target}`)
      const resolved = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(base, target))
      if (files[resolved] === undefined) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `relationship target is missing: ${resolved}`)
    }
  }
  for (const name of names.filter(entry => entry.startsWith('ppt/media/') && !entry.endsWith('/'))) {
    if (files[name]!.byteLength === 0) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `empty media part: ${name}`)
  }
  return { pageCount: slideNames.length, widthEmu, heightEmu, entries: names }
}

export async function createPptx(
  browser: BrowserRuntime,
  owner: SessionOwner,
  workspace: string,
  htmlPathInput: string,
  outlinePathInput: string,
  outputPathInput: string,
  fallbackMode: PptFallbackMode = 'reject',
  signal?: AbortSignal,
  transitions?: SlideTransitionPlan,
  animations?: AnimationPlan,
): Promise<PptCreateResult> {
  throwIfAborted(signal, 'PPT_CREATE_ABORTED')
  const [htmlPath, outlinePath, outputPath] = await Promise.all([
    resolveWorkspacePath(workspace, htmlPathInput, { mustExist: true, kind: 'file' }),
    resolveWorkspacePath(workspace, outlinePathInput, { mustExist: true, kind: 'file' }),
    resolveWorkspacePath(workspace, outputPathInput),
  ])
  const artifactRoot = dirname(outlinePath)
  if (dirname(htmlPath) !== artifactRoot || dirname(outputPath) !== artifactRoot) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', 'HTML, outline, and PPTX output must share one artifact directory')
  }
  const outline = validatePptOutline(JSON.parse(await readFile(outlinePath, 'utf8')))
  try {
    await validateDeckHtmlSource(workspace, artifactRoot, await readFile(htmlPath, 'utf8'), outline.length)
  } catch (error) {
    const issues = error instanceof PptError && Array.isArray(error.details?.issues) ? error.details.issues.map(String) : []
    if (issues.some(issue => /missing or invalid local asset/iu.test(issue))) {
      throw new PptError('PPT_CREATE_ASSET_MISSING', 'HTML references a missing or invalid local asset', { cause: error, details: { issues } })
    }
    throw error
  }
  let ir: DeckIR
  try {
    ir = await browser.extractDeckIr(owner, workspace, htmlPath, outline.length, signal)
  } catch (error) {
    if (signal?.aborted) throw new PptError('PPT_CREATE_ABORTED', 'PPTX creation was cancelled', { cause: error })
    throw error
  }
  outline.forEach((slide, index) => {
    ir.slides[index]!.speakerNotes = slide.content.flatMap(item => item.kind === 'note' && item.purpose === 'speaker' ? [item.text] : [])
  })
  const rasterized: RasterizedElementRecord[] = []
  for (const slide of ir.slides) {
    for (const element of slide.elements) {
      if (element.unsupportedReason === undefined) continue
      if (fallbackMode !== 'rasterize-element') {
        throw new PptError('PPT_CREATE_UNSUPPORTED_ELEMENT', `page ${slide.page} element ${element.id}: ${element.unsupportedReason}`)
      }
      const target = join(artifactRoot, 'assets', 'rasterized', `page-${slide.page}-${element.id}.png`)
      const imagePath = await browser.rasterizeElement(owner, workspace, htmlPath, element.id, target, signal)
      rasterized.push({ page: slide.page, element_id: element.id, reason: element.unsupportedReason, image_path: imagePath })
      element.kind = 'image'
      element.imagePath = join(workspace, imagePath)
      delete element.unsupportedReason
    }
  }

  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'DSH_PPT_16_9', width: SLIDE_WIDTH_IN, height: SLIDE_HEIGHT_IN })
  pptx.layout = 'DSH_PPT_16_9'
  pptx.author = 'DSH PPT'
  pptx.company = 'DSH'
  pptx.subject = 'Editable PPTX generated from constrained HTML'
  pptx.title = outline[0]?.title ?? 'Presentation'
  let nativeElementCount = 0
  for (const slideIr of ir.slides) {
    const slide = pptx.addSlide()
    for (const element of slideIr.elements) {
      addNativeElement(pptx, slide, element, workspace, artifactRoot)
      if (!rasterized.some(item => item.page === slideIr.page && item.element_id === element.id)) nativeElementCount += 1
    }
    if (slideIr.speakerNotes.length > 0) slide.addNotes(slideIr.speakerNotes.join('\n\n'))
  }
  const temporary = join(artifactRoot, `.deck.${process.pid}.${Date.now()}.pptx`)
  try {
    await pptx.writeFile({ fileName: temporary, compression: true })
    throwIfAborted(signal, 'PPT_CREATE_ABORTED')
    const written = new Uint8Array(await readFile(temporary))
    // The transition and animation plans rewrite slide parts inside the package before
    // the atomic commit. Without a plan the bytes written by pptxgenjs are committed
    // untouched, so a deck that opts out stays byte-comparable with its source.
    let bytes: Uint8Array = written
    if (transitions !== undefined) bytes = rewritePptxTransitions(bytes, transitions)
    if (animations !== undefined) bytes = rewritePptxAnimations(bytes, animations)
    inspectPptxPackage(bytes, outline.length)
    await atomicWriteFile(outputPath, bytes, { signal })
    return {
      pptx_path: workspaceRelative(workspace, outputPath), page_count: outline.length,
      native_element_count: nativeElementCount, rasterized_elements: rasterized, structural_status: 'passed',
    }
  } catch (error) {
    if (signal?.aborted) throw new PptError('PPT_CREATE_ABORTED', 'PPTX creation was cancelled', { cause: error })
    throw error
  } finally {
    await rm(temporary, { force: true })
  }
}
