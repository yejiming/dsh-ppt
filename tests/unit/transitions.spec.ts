import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync, zipSync } from 'fflate'
import PptxGenJS from 'pptxgenjs'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { PptError } from '../../src/errors.ts'
import { SLIDE_HEIGHT_IN, SLIDE_WIDTH_IN } from '../../src/ir.ts'
import { inspectPptxPackage } from '../../src/pptx.ts'
import {
  injectSlideTransition,
  inspectZipEntries,
  planSlideTransitions,
  rewritePptxTransitions,
  transitionElementXml,
  type SlideTransition,
  type SlideTransitionPlan,
  type ZipEntryInfo,
} from '../../src/transitions.ts'

const SLIDE_NAMESPACE = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
const DRAWING_NAMESPACE = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const SLIDE_PART = /^ppt\/slides\/slide\d+\.xml$/u

function slideXml(options: { clrMapOvr: boolean; timing: boolean }): string {
  const clrMapOvr = options.clrMapOvr ? '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1"/></p:clrMapOvr>' : ''
  const timing = options.timing ? '<p:timing><p:tnLst/></p:timing>' : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${SLIDE_NAMESPACE} ${DRAWING_NAMESPACE}><p:cSld><p:spTree><p:nvGrpSpPr/></p:spTree></p:cSld>${clrMapOvr}${timing}</p:sld>`
}

interface XmlNode {
  nodeName: string
  children: { length: number; item(index: number): XmlNode | null }
}

interface XmlDocument {
  documentElement: XmlNode | null
}

async function childNames(xml: string): Promise<string[]> {
  const { Window } = await import('happy-dom')
  const domWindow = new Window()
  const Parser = (domWindow as unknown as { DOMParser: new () => { parseFromString(text: string, type: string): unknown } }).DOMParser
  const document = new Parser().parseFromString(xml, 'text/xml') as unknown as XmlDocument
  domWindow.close()
  const root = document.documentElement
  if (root === null) throw new Error('slide XML did not parse into a root element')
  const names: string[] = []
  for (let index = 0; index < root.children.length; index += 1) {
    const child = root.children.item(index)
    if (child !== null) names.push(child.nodeName)
  }
  return names
}

function thrownBy(run: () => unknown): PptError {
  let failure: unknown
  try {
    run()
  } catch (error) {
    failure = error
  }
  if (failure instanceof PptError) return failure
  throw new Error(`expected a PptError, received ${failure === undefined ? 'no error' : String(failure)}`)
}

function rawEntryRange(bytes: Uint8Array, info: ZipEntryInfo): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const nameLength = view.getUint16(info.local_header_offset + 26, true)
  const extraLength = view.getUint16(info.local_header_offset + 28, true)
  const start = info.local_header_offset
  return bytes.slice(start, start + 30 + nameLength + extraLength + info.compressed_size)
}

function entryRange(bytes: Uint8Array, name: string): Uint8Array {
  const info = inspectZipEntries(bytes).find(entry => entry.name === name)
  if (info === undefined) throw new Error(`zip entry is missing: ${name}`)
  return rawEntryRange(bytes, info)
}

let workspace: string
let deck: Uint8Array

beforeAll(async () => {
  workspace = await mkdtemp(join(tmpdir(), 'dsh-ppt-transitions-'))
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'DSH_PPT_16_9', width: SLIDE_WIDTH_IN, height: SLIDE_HEIGHT_IN })
  pptx.layout = 'DSH_PPT_16_9'
  pptx.author = 'DSH PPT'
  for (const title of ['Cover', 'Argument', 'Timeline']) {
    pptx.addSlide().addText(title, { x: 1, y: 1, w: 8, h: 1 })
  }
  const target = join(workspace, 'deck.pptx')
  await pptx.writeFile({ fileName: target, compression: true })
  deck = new Uint8Array(await readFile(target))
})

afterAll(async () => {
  await rm(workspace, { recursive: true, force: true })
})

describe('slide transition elements', () => {
  it('renders only elements ECMA-376 defines and rejects unknown input', () => {
    expect(transitionElementXml({ type: 'fade', speed: 'slow' })).toBe('<p:transition spd="slow"><p:fade/></p:transition>')
    expect(transitionElementXml({ type: 'dissolve' })).toBe('<p:transition><p:dissolve/></p:transition>')
    expect(transitionElementXml({ type: 'push', direction: 'left' })).toBe('<p:transition><p:push dir="l"/></p:transition>')
    expect(transitionElementXml({ type: 'cover', direction: 'up', speed: 'med' })).toBe('<p:transition spd="med"><p:cover dir="u"/></p:transition>')
    expect(transitionElementXml({ type: 'cut' })).toBe('<p:transition><p:cut/></p:transition>')
    expect(transitionElementXml({ type: 'fade', advanceAfterMs: 4000 })).toBe('<p:transition advTm="4000"><p:fade/></p:transition>')

    const unknown = thrownBy(() => transitionElementXml({ type: 'spin' as SlideTransition['type'] }))
    expect(unknown.code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(unknown.message).toContain('spin')
    expect(thrownBy(() => transitionElementXml({ type: 'fade', direction: 'left' })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => transitionElementXml({ type: 'push', direction: 'diagonal' as SlideTransition['direction'] })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => transitionElementXml({ type: 'push', advanceAfterMs: 0 })).code).toBe('PPT_CREATE_INPUT_INVALID')
  })

  it('inserts after clrMapOvr and keeps the CT_Slide order', async () => {
    const injected = injectSlideTransition(slideXml({ clrMapOvr: true, timing: true }), { type: 'dissolve', speed: 'med' })
    expect(await childNames(injected)).toEqual(['p:cSld', 'p:clrMapOvr', 'p:transition', 'p:timing'])
    expect(injected).toContain('</p:clrMapOvr><p:transition spd="med"><p:dissolve/></p:transition><p:timing>')
    expect(injected.match(/<p:transition\b/gu)).toHaveLength(1)
  })

  it('falls back to the cSld boundary when clrMapOvr is absent', async () => {
    const injected = injectSlideTransition(slideXml({ clrMapOvr: false, timing: true }), { type: 'cover', direction: 'left' })
    expect(await childNames(injected)).toEqual(['p:cSld', 'p:transition', 'p:timing'])
    expect(injected).toContain('</p:cSld><p:transition><p:cover dir="l"/></p:transition><p:timing>')
  })

  it('declares the presentation namespace when the root omits it', () => {
    const injected = injectSlideTransition('<p:sld><p:cSld/></p:sld>', { type: 'fade' })
    expect(injected).toBe('<p:sld><p:cSld/><p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:fade/></p:transition></p:sld>')
  })

  it('rejects a second injection instead of stacking two transitions', () => {
    const once = injectSlideTransition(slideXml({ clrMapOvr: true, timing: true }), { type: 'fade' })
    const again = thrownBy(() => injectSlideTransition(once, { type: 'wipe', direction: 'right' }))
    expect(again.code).toBe('PPT_CREATE_INVALID_PACKAGE')
    expect(once.match(/<p:transition\b/gu)).toHaveLength(1)
  })

  it('rejects XML that is not a complete p:sld part', () => {
    expect(thrownBy(() => injectSlideTransition('<p:notes><p:cSld/></p:notes>', { type: 'fade' })).code).toBe('PPT_CREATE_INVALID_PACKAGE')
    expect(thrownBy(() => injectSlideTransition('<p:sld><p:cSld/></p:slide>', { type: 'fade' })).code).toBe('PPT_CREATE_INVALID_PACKAGE')
  })
})

describe('slide transition rhythm', () => {
  it('maps slide roles onto a paced default sequence', () => {
    const plan = planSlideTransitions([
      { page: 1, type: 'cover' },
      { page: 2, type: 'agenda' },
      { page: 3, type: 'content' },
      { page: 4, type: 'timeline' },
      { page: 5, type: 'unknown-role' },
    ])
    expect(plan.get(1)).toEqual({ type: 'fade', speed: 'slow' })
    expect(plan.get(2)).toEqual({ type: 'push', direction: 'left', speed: 'fast' })
    expect(plan.get(3)).toEqual({ type: 'dissolve', speed: 'med' })
    expect(plan.get(4)).toEqual({ type: 'cover', direction: 'up', speed: 'med' })
    expect(plan.get(5)).toEqual({ type: 'dissolve', speed: 'med' })
    expect(thrownBy(() => planSlideTransitions([{ page: 0, type: 'cover' }])).code).toBe('PPT_CREATE_INPUT_INVALID')
  })
})

describe('PPTX transition rewrite', () => {
  it('rewrites arranged slide parts and leaves every other entry byte-identical', () => {
    const plan: SlideTransitionPlan = new Map<number, SlideTransition>([
      [1, { type: 'fade', speed: 'slow' }],
      [2, { type: 'push', direction: 'left' }],
    ])
    const before = inspectZipEntries(deck)
    const injected = rewritePptxTransitions(deck, plan)
    const after = inspectZipEntries(injected)

    expect(after.map(entry => entry.name)).toEqual(before.map(entry => entry.name))
    expect(after).toHaveLength(before.length)
    expect(after.filter(entry => SLIDE_PART.test(entry.name))).toHaveLength(3)
    for (const [index, entry] of before.entries()) {
      const updated = after[index]!
      expect(updated.name).toBe(entry.name)
      expect(updated.method).toBe(entry.method)
      if (SLIDE_PART.test(entry.name)) continue
      expect(rawEntryRange(injected, updated)).toEqual(rawEntryRange(deck, entry))
    }
    expect(entryRange(injected, 'ppt/slides/slide1.xml')).not.toEqual(entryRange(deck, 'ppt/slides/slide1.xml'))
    expect(entryRange(injected, 'ppt/slides/slide2.xml')).not.toEqual(entryRange(deck, 'ppt/slides/slide2.xml'))
    expect(entryRange(injected, 'ppt/slides/slide3.xml')).toEqual(entryRange(deck, 'ppt/slides/slide3.xml'))
  })

  it('keeps the archive complete and the package structurally valid', () => {
    const plan = planSlideTransitions([
      { page: 1, type: 'cover' },
      { page: 2, type: 'content' },
      { page: 3, type: 'timeline' },
    ])
    const injected = rewritePptxTransitions(deck, plan)
    const files = unzipSync(injected)

    expect(Object.keys(files)).toEqual(Object.keys(unzipSync(deck)))
    expect(inspectPptxPackage(injected, 3).pageCount).toBe(3)

    const slides = [1, 2, 3].map(page => strFromU8(files[`ppt/slides/slide${page}.xml`]!))
    for (const xml of slides) {
      expect(xml.match(/<p:transition\b/gu)).toHaveLength(1)
      expect(xml.indexOf('<p:transition')).toBeGreaterThan(xml.indexOf('</p:clrMapOvr>'))
    }
    expect(slides[0]).toContain('<p:fade/>')
    expect(slides[1]).toContain('<p:dissolve/>')
    expect(slides[2]).toContain('<p:cover dir="u"/>')
  })

  it('leaves unarranged pages untouched and returns the input for an empty plan', () => {
    const injected = rewritePptxTransitions(deck, new Map([[1, { type: 'cut' }]]))
    const files = unzipSync(injected)
    expect(strFromU8(files['ppt/slides/slide1.xml']!)).toContain('<p:cut/>')
    expect(strFromU8(files['ppt/slides/slide2.xml']!)).not.toContain('<p:transition')
    expect(strFromU8(files['ppt/slides/slide3.xml']!)).not.toContain('<p:transition')
    expect(rewritePptxTransitions(deck, new Map())).toBe(deck)
  })

  it('keeps the original compression method when rewriting a deflated package', () => {
    const deflated = zipSync(unzipSync(deck), { level: 6 })
    const before = inspectZipEntries(deflated)
    expect(before.find(entry => entry.name === 'ppt/slides/slide1.xml')?.method).toBe(8)

    const injected = rewritePptxTransitions(deflated, new Map([[1, { type: 'wipe', direction: 'right' }]]))
    const after = inspectZipEntries(injected)
    expect(after.map(entry => entry.name)).toEqual(before.map(entry => entry.name))
    expect(after.find(entry => entry.name === 'ppt/slides/slide1.xml')?.method).toBe(8)
    expect(entryRange(injected, 'ppt/slides/slide2.xml')).toEqual(entryRange(deflated, 'ppt/slides/slide2.xml'))
    expect(strFromU8(unzipSync(injected)['ppt/slides/slide1.xml']!)).toContain('<p:wipe dir="r"/>')
  })

  it('fails loudly when a page is arranged twice or not at all', () => {
    const once = rewritePptxTransitions(deck, new Map([[1, { type: 'fade' }]]))
    const repeated = thrownBy(() => rewritePptxTransitions(once, new Map([[1, { type: 'wipe', direction: 'left' }]])))
    expect(repeated.code).toBe('PPT_CREATE_INVALID_PACKAGE')
    expect(strFromU8(unzipSync(once)['ppt/slides/slide1.xml']!).match(/<p:transition\b/gu)).toHaveLength(1)

    const missing = thrownBy(() => rewritePptxTransitions(deck, new Map([[9, { type: 'fade' }]])))
    expect(missing.code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(missing.message).toContain('9')

    const unknown = thrownBy(() => rewritePptxTransitions(deck, new Map([[1, { type: 'morph' as SlideTransition['type'] }]])))
    expect(unknown.code).toBe('PPT_CREATE_INPUT_INVALID')
  })
})
