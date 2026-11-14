import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strFromU8, unzipSync } from 'fflate'
import PptxGenJS from 'pptxgenjs'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import {
  injectSlideAnimation,
  planSlideAnimations,
  rewritePptxAnimations,
  normalizeTextAnimation,
  TEXT_ANIMATION_EFFECTS,
  TEXT_ANIMATION_MAX_PER_PAGE,
  type TextAnimation,
} from '../../src/animation.ts'
import { PptError } from '../../src/errors.ts'
import { SLIDE_HEIGHT_IN, SLIDE_WIDTH_IN } from '../../src/ir.ts'
import { inspectPptxPackage } from '../../src/pptx.ts'
import { inspectZipEntries, type ZipEntryInfo } from '../../src/zip.ts'

const SLIDE_PART = /^ppt\/slides\/slide\d+\.xml$/u
const SHAPES_NAMESPACE = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

function slideXml(options: { clrMapOvr?: boolean; transition?: boolean; timing?: string; shapes?: string } = {}): string {
  const clrMapOvr = options.clrMapOvr === false ? '' : '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
  const transition = options.transition === true ? '<p:transition spd="med"><p:dissolve/></p:transition>' : ''
  const timing = options.timing ?? ''
  const shapes = options.shapes ?? '<p:sp><p:nvSpPr><p:cNvPr id="2" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:p><a:r><a:t>a</a:t></a:r></a:p></p:txBody></p:sp>'
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${SHAPES_NAMESPACE}><p:cSld><p:spTree>${shapes}</p:spTree></p:cSld>${clrMapOvr}${transition}${timing}</p:sld>`
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

describe('text animation validation', () => {
  it('accepts every catalogued effect and normalizes the target', () => {
    for (const effect of TEXT_ANIMATION_EFFECTS) {
      expect(normalizeTextAnimation({ target: '  title  ', effect }).target).toBe('title')
    }
  })

  it('rejects unknown effects, directions, starts, and durations', () => {
    expect(thrownBy(() => normalizeTextAnimation({ target: 'a', effect: 'spin' as TextAnimation['effect'] })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => normalizeTextAnimation({ target: '', effect: 'fade' })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => normalizeTextAnimation({ target: 'a', effect: 'fade', direction: 'up' })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => normalizeTextAnimation({ target: 'a', effect: 'fade', start: 'later' as TextAnimation['start'] })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => normalizeTextAnimation({ target: 'a', effect: 'fade', durationMs: 0 })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => normalizeTextAnimation({ target: 'a', effect: 'fade', durationMs: 60_001 })).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => normalizeTextAnimation({ target: 'a', effect: 'wipe', direction: 'diagonal' as TextAnimation['direction'] })).code).toBe('PPT_CREATE_INPUT_INVALID')
  })

  it('keeps a direction only on effects that render one', () => {
    expect(normalizeTextAnimation({ target: 'a', effect: 'wipe', direction: 'left' }).direction).toBe('left')
    expect(normalizeTextAnimation({ target: 'a', effect: 'fly-in', direction: 'up' }).direction).toBe('up')
    expect(normalizeTextAnimation({ target: 'a', effect: 'zoom' }).direction).toBeUndefined()
  })

  it('rejects a duplicate animation for the same build unit and an oversized page', () => {
    expect(thrownBy(() => planSlideAnimations([{
      page: 1,
      animations: [{ target: 'a', effect: 'fade' }, { target: 'a', effect: 'wipe' }],
    }])).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => planSlideAnimations([{
      page: 1,
      animations: Array.from({ length: TEXT_ANIMATION_MAX_PER_PAGE + 1 }, (_, index) => ({ target: `shape-${index}`, effect: 'fade' as const })),
    }])).code).toBe('PPT_CREATE_INPUT_INVALID')
    expect(thrownBy(() => planSlideAnimations([{ page: 0, animations: [] }])).code).toBe('PPT_CREATE_INPUT_INVALID')
    // The same target may be animated once as a whole and once per paragraph.
    const plan = planSlideAnimations([{ page: 1, animations: [
      { target: 'a', effect: 'fade' },
      { target: 'a', effect: 'fade', byParagraph: true },
    ] }])
    expect(plan.get(1)).toHaveLength(2)
  })
})

describe('slide animation injection', () => {
  it('emits the preset triple, the visibility set, and the timeline envelope', () => {
    const injected = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'fade' }])
    expect(injected).toContain('<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">')
    expect(injected).toContain('nodeType="mainSeq"')
    expect(injected).toContain('presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"')
    expect(injected).toContain('<p:attrName>style.visibility</p:attrName>')
    expect(injected).toContain('<p:to><p:strVal val="visible"/></p:to>')
    expect(injected).toContain('<p:animEffect transition="in" filter="fade">')
    expect(injected).toContain('<p:spTgt spid="2"/>')
    expect(injected).toContain('<p:prevCondLst><p:cond evt="onPrev" delay="0">')
    expect(injected).toContain('<p:nextCondLst><p:cond evt="onNext" delay="0">')
    // The effect time node is written before its children, as PowerPoint does.
    const effectId = Number(/<p:cTn id="(\d+)" presetID="10"/u.exec(injected)![1])
    const childId = Number(/<p:cTn id="(\d+)" dur="1" fill="hold">/u.exec(injected)![1])
    expect(effectId).toBeLessThan(childId)
  })

  it('places the timeline after clrMapOvr and after an existing transition', () => {
    const without = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'fade' }])
    expect(without).toContain('</p:clrMapOvr><p:timing>')
    const withTransition = injectSlideAnimation(slideXml({ transition: true }), [{ target: 'title', effect: 'fade' }])
    expect(withTransition).toContain('</p:transition><p:timing>')
    expect(withTransition).not.toContain('</p:clrMapOvr><p:timing>')
  })

  it('replaces an existing timeline instead of stacking a second one', () => {
    const once = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'fade' }])
    const twice = injectSlideAnimation(once, [{ target: 'title', effect: 'wipe' }])
    expect(twice.match(/<p:timing\b/gu)).toHaveLength(1)
    expect(twice).toContain('filter="wipe(down)"')
    expect(twice).not.toContain('filter="fade"')
  })

  it('writes a direction dependent subtype and filter', () => {
    const wipe = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'wipe', direction: 'left' }])
    expect(wipe).toContain('presetSubtype="8"')
    expect(wipe).toContain('filter="wipe(left)"')
    // Peek keeps PowerPoint's inverted pairing: from-bottom subtype, upward filter.
    const peek = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'peek', direction: 'down' }])
    expect(peek).toContain('presetSubtype="4"')
    expect(peek).toContain('filter="wipe(up)"')
  })

  it('writes motion origins as formulas and scales as floats', () => {
    const zoom = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'zoom' }])
    expect(zoom).toContain('<p:attrName>ppt_w</p:attrName>')
    expect(zoom).toContain('<p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>')
    expect(zoom).toContain('<p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav>')
    const fly = injectSlideAnimation(slideXml(), [{ target: 'title', effect: 'fly-in', direction: 'down' }])
    expect(fly).toContain('<p:tav tm="0"><p:val><p:strVal val="1+#ppt_h/2"/></p:val></p:tav>')
    expect(fly).toContain('<p:cBhvr additive="base">')
  })

  it('builds one step per paragraph when asked', () => {
    const shapes = '<p:sp><p:nvSpPr><p:cNvPr id="4" name="body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody>'
      + '<a:p><a:r><a:t>one</a:t></a:r></a:p><a:p><a:r><a:t>two</a:t></a:r></a:p><a:p><a:r><a:t>three</a:t></a:r></a:p>'
      + '</p:txBody></p:sp>'
    const injected = injectSlideAnimation(slideXml({ shapes }), [{ target: 'body', effect: 'fade', byParagraph: true }])
    expect([...new Set([...injected.matchAll(/<p:pRg st="(\d+)" end="(\d+)"\/>/gu)].map(match => match[1]))]).toEqual(['0', '1', '2'])
    expect(injected.match(/nodeType="clickEffect"/gu)).toHaveLength(3)
    expect(injected).not.toContain('nodeType="afterEffect"')
  })

  it('marks with-previous and after-previous steps without wrapping them', () => {
    const injected = injectSlideAnimation(slideXml({
      shapes: '<p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody/></p:sp>'
        + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="b"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody/></p:sp>'
        + '<p:sp><p:nvSpPr><p:cNvPr id="4" name="c"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody/></p:sp>',
    }), [
      { target: 'a', effect: 'fade' },
      { target: 'b', effect: 'fade', start: 'with-previous' },
      { target: 'c', effect: 'fade', start: 'after-previous' },
    ])
    expect(injected.match(/nodeType="clickEffect"/gu)).toHaveLength(1)
    expect(injected.match(/nodeType="withEffect"/gu)).toHaveLength(1)
    expect(injected.match(/nodeType="afterEffect"/gu)).toHaveLength(1)
    // PowerPoint rejects a hand-written afterEffect nested inside its own delay
    // node, so the effect nodes must stay direct siblings inside one group.
    expect(injected.match(/<p:stCondLst><p:cond delay="indefinite"\/><\/p:stCondLst>/gu)).toHaveLength(1)
  })

  it('surfaces a missing target instead of emitting a dangling reference', () => {
    expect(thrownBy(() => injectSlideAnimation(slideXml(), [{ target: 'ghost', effect: 'fade' }])).code).toBe('PPT_CREATE_INPUT_INVALID')
  })
})

let workspace: string
let deck: Uint8Array

beforeAll(async () => {
  workspace = await mkdtemp(join(tmpdir(), 'dsh-ppt-animation-'))
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'DSH_PPT_16_9', width: SLIDE_WIDTH_IN, height: SLIDE_HEIGHT_IN })
  pptx.layout = 'DSH_PPT_16_9'
  pptx.author = 'DSH PPT'
  for (const page of [1, 2, 3]) {
    const slide = pptx.addSlide()
    slide.addText(`page ${page}`, { x: 1, y: 1, w: 8, h: 1, objectName: `title-${page}` })
  }
  const target = join(workspace, 'deck.pptx')
  await pptx.writeFile({ fileName: target, compression: true })
  deck = new Uint8Array(await readFile(target))
})

afterAll(async () => {
  await rm(workspace, { recursive: true, force: true })
})

describe('PPTX animation rewrite', () => {
  it('rewrites only the planned slide parts and leaves every other entry byte-identical', () => {
    const plan = planSlideAnimations([{ page: 2, animations: [{ target: 'title-2', effect: 'fade' }] }])
    const before = inspectZipEntries(deck)
    const injected = rewritePptxAnimations(deck, plan)
    const after = inspectZipEntries(injected)

    expect(after.map(entry => entry.name)).toEqual(before.map(entry => entry.name))
    for (const [index, entry] of before.entries()) {
      if (SLIDE_PART.test(entry.name)) continue
      expect(rawEntryRange(injected, after[index]!)).toEqual(rawEntryRange(deck, entry))
    }
    expect(rawEntryRange(injected, entryRange0(after, 'ppt/slides/slide1.xml'))).toEqual(rawEntryRange(deck, entryRange0(before, 'ppt/slides/slide1.xml')))
    expect(rawEntryRange(injected, entryRange0(after, 'ppt/slides/slide3.xml'))).toEqual(rawEntryRange(deck, entryRange0(before, 'ppt/slides/slide3.xml')))
    expect(rawEntryRange(injected, entryRange0(after, 'ppt/slides/slide2.xml'))).not.toEqual(rawEntryRange(deck, entryRange0(before, 'ppt/slides/slide2.xml')))
  })

  it('keeps the archive complete and the package structurally valid', () => {
    const injected = rewritePptxAnimations(deck, planSlideAnimations([
      { page: 1, animations: [{ target: 'title-1', effect: 'fade', byParagraph: true }] },
      { page: 3, animations: [{ target: 'title-3', effect: 'zoom' }] },
    ]))
    const files = unzipSync(injected)
    expect(Object.keys(files)).toEqual(Object.keys(unzipSync(deck)))
    expect(inspectPptxPackage(injected, 3).pageCount).toBe(3)
    expect(strFromU8(files['ppt/slides/slide1.xml']!)).toContain('<p:timing>')
    expect(strFromU8(files['ppt/slides/slide2.xml']!)).not.toContain('<p:timing>')
  })

  it('returns the input untouched for an empty plan and reports an unmatched page', () => {
    expect(rewritePptxAnimations(deck, planSlideAnimations([]))).toBe(deck)
    expect(thrownBy(() => rewritePptxAnimations(deck, planSlideAnimations([{ page: 9, animations: [{ target: 'title-1', effect: 'fade' }] }]))).code)
      .toBe('PPT_CREATE_INPUT_INVALID')
  })
})

function entryRange0(entries: readonly ZipEntryInfo[], name: string): ZipEntryInfo {
  const info = entries.find(entry => entry.name === name)
  if (info === undefined) throw new Error(`zip entry is missing: ${name}`)
  return info
}
