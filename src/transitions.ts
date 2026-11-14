import { strFromU8, strToU8 } from 'fflate'
import { PptError } from './errors.ts'
import { rewriteZip } from './zip.ts'

/**
 * Slide transitions live outside the pptxgenjs feature set, so a deck opts in by
 * declaring a per-page plan and letting the PPTX writer rewrite only the slide XML
 * parts after pptxgenjs produced the package. Every other zip entry is copied
 * verbatim so the result stays byte-comparable with the untouched archive.
 */

export const SLIDE_TRANSITION_TYPES = ['cut', 'fade', 'dissolve', 'push', 'wipe', 'cover', 'pull'] as const
export type SlideTransitionType = typeof SLIDE_TRANSITION_TYPES[number]

export const SLIDE_TRANSITION_DIRECTIONS = ['left', 'right', 'up', 'down'] as const
export type SlideTransitionDirection = typeof SLIDE_TRANSITION_DIRECTIONS[number]

export const SLIDE_TRANSITION_SPEEDS = ['slow', 'med', 'fast'] as const
export type SlideTransitionSpeed = typeof SLIDE_TRANSITION_SPEEDS[number]

export interface SlideTransition {
  type: SlideTransitionType
  direction?: SlideTransitionDirection
  speed?: SlideTransitionSpeed
  advanceAfterMs?: number
  advanceOnClick?: boolean
}

export interface SlideRhythmPage {
  page: number
  type: string
}

export type SlideTransitionPlan = ReadonlyMap<number, SlideTransition>

export { inspectZipEntries, type ZipEntryInfo } from './zip.ts'

const PRESENTATION_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const SLIDE_PART = /^ppt\/slides\/slide(\d+)\.xml$/u
const DIRECTIONAL_TYPES: readonly SlideTransitionType[] = ['push', 'wipe', 'cover', 'pull']
const DIRECTION_ATTRIBUTES: Record<SlideTransitionDirection, string> = { left: 'l', right: 'r', up: 'u', down: 'd' }

/** Default rhythm: one transition family per slide role so pacing reads as authored, not uniform. */
const SLIDE_RHYTHM = new Map<string, SlideTransition>([
  ['cover', { type: 'fade', speed: 'slow' }],
  ['agenda', { type: 'push', direction: 'left', speed: 'fast' }],
  ['section', { type: 'cover', direction: 'left', speed: 'med' }],
  ['content', { type: 'dissolve', speed: 'med' }],
  ['comparison', { type: 'wipe', direction: 'left', speed: 'med' }],
  ['timeline', { type: 'cover', direction: 'up', speed: 'med' }],
  ['process', { type: 'push', direction: 'up', speed: 'med' }],
  ['data', { type: 'wipe', direction: 'right', speed: 'med' }],
  ['quote', { type: 'dissolve', speed: 'slow' }],
  ['summary', { type: 'fade', speed: 'med' }],
  ['ending', { type: 'fade', speed: 'slow' }],
])
const DEFAULT_RHYTHM: SlideTransition = { type: 'dissolve', speed: 'med' }

export function isSlideTransitionType(value: string): value is SlideTransitionType {
  return (SLIDE_TRANSITION_TYPES as readonly string[]).includes(value)
}

export function normalizeSlideTransition(transition: SlideTransition): SlideTransition {
  const type: string = transition.type
  if (!isSlideTransitionType(type)) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', `unsupported slide transition type: ${type}`, {
      details: { supported: [...SLIDE_TRANSITION_TYPES] },
    })
  }
  if (transition.direction !== undefined) {
    if (!DIRECTIONAL_TYPES.includes(type)) {
      throw new PptError('PPT_CREATE_INPUT_INVALID', `slide transition ${type} does not accept a direction`)
    }
    if (!(SLIDE_TRANSITION_DIRECTIONS as readonly string[]).includes(transition.direction)) {
      throw new PptError('PPT_CREATE_INPUT_INVALID', `unsupported slide transition direction: ${String(transition.direction)}`)
    }
  }
  if (transition.speed !== undefined && !(SLIDE_TRANSITION_SPEEDS as readonly string[]).includes(transition.speed)) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', `unsupported slide transition speed: ${String(transition.speed)}`)
  }
  if (transition.advanceAfterMs !== undefined && (!Number.isInteger(transition.advanceAfterMs) || transition.advanceAfterMs <= 0)) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', `slide transition advanceAfterMs must be a positive integer: ${String(transition.advanceAfterMs)}`)
  }
  if (transition.advanceOnClick !== undefined && transition.advanceOnClick !== false) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', 'slide transition advanceOnClick may only be set to false')
  }
  const normalized: SlideTransition = { type }
  if (transition.direction !== undefined) normalized.direction = transition.direction
  if (transition.speed !== undefined) normalized.speed = transition.speed
  if (transition.advanceAfterMs !== undefined) normalized.advanceAfterMs = transition.advanceAfterMs
  if (transition.advanceOnClick === false) normalized.advanceOnClick = false
  return normalized
}

export function transitionElementXml(transition: SlideTransition): string {
  const normalized = normalizeSlideTransition(transition)
  const attributes: string[] = []
  if (normalized.advanceOnClick === false) attributes.push('advClick="0"')
  if (normalized.advanceAfterMs !== undefined) attributes.push(`advTm="${normalized.advanceAfterMs}"`)
  if (normalized.speed !== undefined && normalized.type !== 'cut') attributes.push(`spd="${normalized.speed}"`)
  const direction = DIRECTIONAL_TYPES.includes(normalized.type) && normalized.direction !== undefined
    ? ` dir="${DIRECTION_ATTRIBUTES[normalized.direction]}"`
    : ''
  const open = attributes.length === 0 ? '<p:transition>' : `<p:transition ${attributes.join(' ')}>`
  return `${open}<p:${normalized.type}${direction}/></p:transition>`
}

function elementEndOffset(body: string, name: string): number {
  const match = new RegExp(`<${name}\\b[^>]*?(/?)>`, 'u').exec(body)
  if (match === null) return -1
  if (match[1] === '/') return match.index + match[0].length
  const close = body.indexOf(`</${name}>`, match.index + match[0].length)
  if (close < 0) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `slide XML has an unterminated ${name} element`)
  return close + name.length + 3
}

/** CT_Slide order is cSld, clrMapOvr, transition, timing, extLst. */
function transitionInsertionOffset(body: string): number {
  const csldEnd = elementEndOffset(body, 'p:cSld')
  if (csldEnd < 0) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML has no p:cSld element')
  const clrMapEnd = elementEndOffset(body, 'p:clrMapOvr')
  return clrMapEnd >= csldEnd ? clrMapEnd : csldEnd
}

export function injectSlideTransition(xml: string, transition: SlideTransition): string {
  const element = transitionElementXml(transition)
  const root = /<p:sld\b[^>]*>/u.exec(xml)
  if (root === null) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML does not contain a p:sld root element')
  const bodyStart = root.index + root[0].length
  const bodyEnd = xml.lastIndexOf('</p:sld>')
  if (bodyEnd < bodyStart) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML has an unterminated p:sld root element')
  const body = xml.slice(bodyStart, bodyEnd)
  if (/<p:transition\b/u.test(body)) {
    throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML already declares a p:transition element')
  }
  const payload = /xmlns:p=/u.test(root[0]) ? element : element.replace('<p:transition', `<p:transition xmlns:p="${PRESENTATION_NAMESPACE}"`)
  const at = bodyStart + transitionInsertionOffset(body)
  return `${xml.slice(0, at)}${payload}${xml.slice(at)}`
}

export function planSlideTransitions(pages: readonly SlideRhythmPage[]): SlideTransitionPlan {
  const plan = new Map<number, SlideTransition>()
  for (const page of pages) {
    if (!Number.isInteger(page.page) || page.page < 1) {
      throw new PptError('PPT_CREATE_INPUT_INVALID', `slide transition page must be a positive integer: ${String(page.page)}`)
    }
    plan.set(page.page, { ...(SLIDE_RHYTHM.get(page.type) ?? DEFAULT_RHYTHM) })
  }
  return plan
}

export function rewritePptxTransitions(data: Uint8Array, plan: SlideTransitionPlan): Uint8Array {
  if (plan.size === 0) return data
  const targets = new Map<number, SlideTransition>()
  for (const [page, transition] of plan) {
    if (!Number.isInteger(page) || page < 1) {
      throw new PptError('PPT_CREATE_INPUT_INVALID', `slide transition page must be a positive integer: ${String(page)}`)
    }
    targets.set(page, transition)
  }
  const injected = new Set<number>()
  const output = rewriteZip(data, {
    transform: (name, read) => {
      const match = SLIDE_PART.exec(name)
      const page = match === null ? undefined : Number(match[1])
      const transition = page === undefined ? undefined : targets.get(page)
      if (page === undefined || transition === undefined) return undefined
      injected.add(page)
      return strToU8(injectSlideTransition(strFromU8(read()), transition))
    },
  })
  for (const page of targets.keys()) {
    if (!injected.has(page)) throw new PptError('PPT_CREATE_INPUT_INVALID', `no slide part matches transition page ${page}`)
  }
  return output
}
