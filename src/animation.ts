import { strFromU8, strToU8 } from 'fflate'
import { PptError } from './errors.ts'
import { rewriteZip } from './zip.ts'

/**
 * Text and shape animation lives in the slide's `p:timing` timeline, which
 * pptxgenjs does not emit at all. A deck opts in by declaring, per page, which
 * element plays which entrance effect; the finished package is then rewritten so
 * only `ppt/slides/slideN.xml` changes.
 *
 * The effect catalogue and every literal in the emitted XML below were taken from
 * timelines that PowerPoint 16.0 wrote itself (probed through COM), not from the
 * prose of the specification, because the specification deliberately leaves the
 * `presetID` numbering unspecified:
 *
 * - `presetClass` partitions the `presetID` numbering space. `presetID="10"` is
 *   Fade only under `presetClass="entr"`; the same number means something else
 *   under `emph`, `exit`, or `path`. This module therefore always writes both.
 * - `presetSubtype` is a direction bitmask, not a per-effect enum:
 *   1 = from top, 2 = from right, 4 = from bottom, 8 = from left, 16 = in,
 *   32 = out. Derived values are the bitwise or: 10 = 2|8 = horizontal,
 *   5 = 1|4 = vertical, 9 = top-left, 3 = top-right, 6 = bottom-right,
 *   12 = bottom-left. Confirmed against the defaults PowerPoint emits
 *   (Fade 0, Blinds 10, Box 16, Strips 12, Wheel 1, Zoom 16, Split 21).
 * - The renderer obeys the effect *nodes*, not the preset triple: `p:animEffect`
 *   carries a filter string and `p:anim` carries explicit from/to values, while
 *   the preset triple only drives the UI's effect list. That is why a direction
 *   change is expressed by rewriting the node, and why `peek` can legitimately
 *   pair `presetSubtype="4"` with `filter="wipe(up)"` (PowerPoint emits exactly
 *   that: peek's filter runs opposite to its subtype).
 * - Every entrance effect PowerPoint emitted carried a `p:set` that forces
 *   `style.visibility` to `visible`; without it the shape can stay hidden.
 */

export const TEXT_ANIMATION_EFFECTS = [
  'appear',
  'flash-once',
  'fade',
  'dissolve',
  'wedge',
  'wipe',
  'blinds',
  'checkerboard',
  'random-bars',
  'box',
  'circle',
  'diamond',
  'plus',
  'split',
  'strips',
  'wheel',
  'zoom',
  'fly-in',
  'crawl',
  'peek',
  'stretch',
  'swivel',
  'spiral',
  'bounce',
  'credits',
  'float-in',
  'grow-turn',
  'rise-up',
  'unfold',
] as const
export type TextAnimationEffect = typeof TEXT_ANIMATION_EFFECTS[number]

export const TEXT_ANIMATION_DIRECTIONS = ['left', 'right', 'up', 'down'] as const
export type TextAnimationDirection = typeof TEXT_ANIMATION_DIRECTIONS[number]

export const TEXT_ANIMATION_STARTS = ['on-click', 'with-previous', 'after-previous'] as const
export type TextAnimationStart = typeof TEXT_ANIMATION_STARTS[number]

export interface TextAnimation {
  /** Element id from the deck IR, which pptxgenjs writes as `p:cNvPr/@name`. */
  target: string
  effect: TextAnimationEffect
  direction?: TextAnimationDirection
  /** Build the effect one paragraph at a time instead of animating the whole box. */
  byParagraph?: boolean
  start?: TextAnimationStart
  durationMs?: number
}

export interface SlideAnimationPlanEntry {
  page: number
  animations: readonly TextAnimation[]
}

export type AnimationPlan = ReadonlyMap<number, readonly TextAnimation[]>

export const TEXT_ANIMATION_DEFAULT_DURATION_MS = 500
export const TEXT_ANIMATION_MAX_PER_PAGE = 24

const PRESENTATION_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const SLIDE_PART = /^ppt\/slides\/slide(\d+)\.xml$/u
const SHAPE_ID = /<p:cNvPr id="(\d+)" name="([^"]*)"/gu

/** Bit values of the `presetSubtype` direction mask. */
const SUBTYPE_TOP = 1
const SUBTYPE_RIGHT = 2
const SUBTYPE_BOTTOM = 4
const SUBTYPE_LEFT = 8
const SUBTYPE_IN = 16
const SUBTYPE_OUT = 32

const DIRECTION_BITS: Record<TextAnimationDirection, number> = {
  up: SUBTYPE_TOP,
  right: SUBTYPE_RIGHT,
  down: SUBTYPE_BOTTOM,
  left: SUBTYPE_LEFT,
}

/** `flash-once` shares Fade's node shape but keeps its own preset id. */
type EffectNodes =
  | { kind: 'visibility' }
  | { kind: 'filter'; filter: string; directional?: Record<TextAnimationDirection, string> }
  | { kind: 'motion'; motions: (direction: TextAnimationDirection | undefined) => readonly Motion[] }

interface Motion {
  attr: 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' | 'style.rotation'
  from: string | number
  to: string | number
  /** PowerPoint emits `additive="base"` on motion behaviours but not on scale ones. */
  additive?: boolean
}

interface EffectDefinition {
  presetID: number
  /** Default subtype PowerPoint emits for the variant with no direction given. */
  subtype: number
  /** Present only when the effect offers a direction; its absence rejects a direction. */
  subtypeFor?: (direction: TextAnimationDirection) => number
  nodes: EffectNodes
}

const directionBit = (direction: TextAnimationDirection): number => DIRECTION_BITS[direction]

/** Effects whose direction only collapses to a horizontal or vertical axis. */
const axesSubtype = (horizontal: number, vertical: number) => (direction: TextAnimationDirection): number =>
  direction === 'up' || direction === 'down' ? vertical : horizontal

const HORIZONTAL = SUBTYPE_RIGHT | SUBTYPE_LEFT
const VERTICAL = SUBTYPE_TOP | SUBTYPE_BOTTOM

/**
 * A motion that slides the shape in from an off-slide edge. The from-bottom
 * pair is PowerPoint's own output; the other edges are its mirror images about
 * the slide centre, using the same `1 + size/2` offset convention.
 */
function edgeMotion(direction: TextAnimationDirection): Motion {
  if (direction === 'down') return { attr: 'ppt_y', from: '1+#ppt_h/2', to: '#ppt_y', additive: true }
  if (direction === 'up') return { attr: 'ppt_y', from: '-#ppt_h/2', to: '#ppt_y', additive: true }
  if (direction === 'right') return { attr: 'ppt_x', from: '1+#ppt_w/2', to: '#ppt_x', additive: true }
  return { attr: 'ppt_x', from: '-#ppt_w/2', to: '#ppt_x', additive: true }
}

function horizontalVertical(horizontal: string, vertical: string): Record<TextAnimationDirection, string> {
  return { left: horizontal, right: horizontal, up: vertical, down: vertical }
}

const EFFECT_DEFINITIONS: Record<TextAnimationEffect, EffectDefinition> = {
  appear: { presetID: 1, subtype: 0, nodes: { kind: 'visibility' } },
  'flash-once': { presetID: 11, subtype: 0, nodes: { kind: 'visibility' } },
  fade: { presetID: 10, subtype: 0, nodes: { kind: 'filter', filter: 'fade' } },
  dissolve: { presetID: 9, subtype: 0, nodes: { kind: 'filter', filter: 'dissolve' } },
  wedge: { presetID: 20, subtype: 0, nodes: { kind: 'filter', filter: 'wedge' } },
  wipe: {
    presetID: 22, subtype: SUBTYPE_BOTTOM, subtypeFor: directionBit,
    nodes: { kind: 'filter', filter: 'wipe(down)', directional: { up: 'wipe(up)', down: 'wipe(down)', left: 'wipe(left)', right: 'wipe(right)' } },
  },
  blinds: {
    presetID: 3, subtype: HORIZONTAL, subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
    nodes: { kind: 'filter', filter: 'blinds(horizontal)', directional: horizontalVertical('blinds(horizontal)', 'blinds(vertical)') },
  },
  checkerboard: {
    presetID: 5, subtype: HORIZONTAL, subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
    nodes: { kind: 'filter', filter: 'checkerboard(across)', directional: horizontalVertical('checkerboard(across)', 'checkerboard(down)') },
  },
  'random-bars': {
    presetID: 14, subtype: HORIZONTAL, subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
    nodes: { kind: 'filter', filter: 'randombar(horizontal)', directional: horizontalVertical('randombar(horizontal)', 'randombar(vertical)') },
  },
  // Box, Circle, Diamond, and Plus carry no direction in the filter string: the
  // direction is expressed purely through the subtype bitmask.
  box: { presetID: 4, subtype: SUBTYPE_IN, subtypeFor: directionBit, nodes: { kind: 'filter', filter: 'box(in)' } },
  circle: { presetID: 6, subtype: SUBTYPE_IN, subtypeFor: directionBit, nodes: { kind: 'filter', filter: 'circle(in)' } },
  diamond: { presetID: 8, subtype: SUBTYPE_IN, subtypeFor: directionBit, nodes: { kind: 'filter', filter: 'diamond(in)' } },
  plus: { presetID: 13, subtype: SUBTYPE_IN, subtypeFor: directionBit, nodes: { kind: 'filter', filter: 'plus(in)' } },
  split: { presetID: 16, subtype: SUBTYPE_IN | SUBTYPE_BOTTOM | SUBTYPE_TOP, nodes: { kind: 'filter', filter: 'barn(inVertical)' } },
  strips: { presetID: 18, subtype: SUBTYPE_BOTTOM | SUBTYPE_LEFT, nodes: { kind: 'filter', filter: 'strips(downLeft)' } },
  wheel: { presetID: 21, subtype: 1, nodes: { kind: 'filter', filter: 'wheel(1)' } },
  zoom: {
    presetID: 23, subtype: SUBTYPE_IN,
    nodes: { kind: 'motion', motions: () => [{ attr: 'ppt_w', from: 0, to: '#ppt_w' }, { attr: 'ppt_h', from: 0, to: '#ppt_h' }] },
  },
  'fly-in': { presetID: 2, subtype: SUBTYPE_BOTTOM, subtypeFor: directionBit, nodes: { kind: 'motion', motions: direction => [edgeMotion(direction ?? 'down')] } },
  crawl: { presetID: 7, subtype: SUBTYPE_BOTTOM, subtypeFor: directionBit, nodes: { kind: 'motion', motions: direction => [edgeMotion(direction ?? 'down')] } },
  peek: {
    // PowerPoint pairs peek's from-bottom subtype with an upward wipe filter.
    presetID: 12, subtype: SUBTYPE_BOTTOM, subtypeFor: directionBit,
    nodes: {
      kind: 'filter', filter: 'wipe(up)',
      directional: { up: 'wipe(down)', down: 'wipe(up)', left: 'wipe(right)', right: 'wipe(left)' },
    },
  },
  stretch: {
    presetID: 17, subtype: HORIZONTAL, subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
    nodes: { kind: 'motion', motions: direction => (direction === 'up' || direction === 'down')
      ? [{ attr: 'ppt_h', from: 0, to: '#ppt_h' }]
      : [{ attr: 'ppt_w', from: 0, to: '#ppt_w' }] },
  },
  swivel: {
    presetID: 19, subtype: HORIZONTAL, subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
    nodes: { kind: 'motion', motions: direction => (direction === 'up' || direction === 'down')
      ? [{ attr: 'ppt_h', from: 0, to: '#ppt_h' }]
      : [{ attr: 'ppt_w', from: 0, to: '#ppt_w' }] },
  },
  spiral: {
    presetID: 15, subtype: 0,
    nodes: {
      kind: 'motion',
      motions: () => [
        { attr: 'ppt_w', from: 0, to: '#ppt_w' },
        { attr: 'ppt_h', from: 0, to: '#ppt_h' },
        { attr: 'ppt_x', from: '#ppt_x', to: '#ppt_x' },
        { attr: 'ppt_y', from: '#ppt_y', to: '#ppt_y' },
      ],
    },
  },
  bounce: {
    presetID: 26, subtype: 0,
    nodes: { kind: 'motion', motions: () => [{ attr: 'ppt_y', from: '#ppt_y-0.25', to: '#ppt_y', additive: true }] },
  },
  credits: {
    presetID: 28, subtype: 0,
    nodes: { kind: 'motion', motions: () => [{ attr: 'ppt_y', from: '#ppt_y+1', to: '#ppt_y-1', additive: true }] },
  },
  'float-in': {
    presetID: 30, subtype: 0,
    nodes: {
      kind: 'motion',
      motions: () => [
        { attr: 'style.rotation', from: -90, to: 0 },
        { attr: 'ppt_x', from: '#ppt_x+0.4', to: '#ppt_x-0.05', additive: true },
        { attr: 'ppt_y', from: '#ppt_y-0.4', to: '#ppt_y+0.1', additive: true },
      ],
    },
  },
  'grow-turn': {
    presetID: 31, subtype: 0,
    nodes: {
      kind: 'motion',
      motions: () => [
        { attr: 'ppt_w', from: 0, to: '#ppt_w' },
        { attr: 'ppt_h', from: 0, to: '#ppt_h' },
        { attr: 'style.rotation', from: 90, to: 0 },
      ],
    },
  },
  'rise-up': {
    presetID: 37, subtype: 0,
    nodes: { kind: 'motion', motions: () => [{ attr: 'ppt_y', from: '#ppt_y+1', to: '#ppt_y-.03', additive: true }] },
  },
  unfold: {
    // PowerPoint writes presetID 40 for Unfold even though the automation enum
    // calls Unfold 37; the XML id and the enum value are not the same space.
    presetID: 40, subtype: 0,
    nodes: { kind: 'motion', motions: () => [{ attr: 'ppt_x', from: '#ppt_x-.1', to: '#ppt_x', additive: true }] },
  },
}

export function isTextAnimationEffect(value: string): value is TextAnimationEffect {
  return (TEXT_ANIMATION_EFFECTS as readonly string[]).includes(value)
}

export function isTextAnimationDirection(value: string): value is TextAnimationDirection {
  return (TEXT_ANIMATION_DIRECTIONS as readonly string[]).includes(value)
}

function fail(message: string): never {
  throw new PptError('PPT_CREATE_INPUT_INVALID', message)
}

export function normalizeTextAnimation(animation: TextAnimation): TextAnimation {
  if (typeof animation.target !== 'string' || animation.target.trim() === '') {
    fail('text animation target must be a non-empty element id')
  }
  if (!isTextAnimationEffect(animation.effect)) {
    fail(`unsupported text animation effect: ${String(animation.effect)}`)
  }
  const definition = EFFECT_DEFINITIONS[animation.effect]
  if (animation.direction !== undefined) {
    if (!isTextAnimationDirection(animation.direction)) fail(`unsupported text animation direction: ${String(animation.direction)}`)
    if (definition.subtypeFor === undefined) {
      fail(`text animation effect ${animation.effect} does not accept a direction`)
    }
  }
  if (animation.start !== undefined && !(TEXT_ANIMATION_STARTS as readonly string[]).includes(animation.start)) {
    fail(`unsupported text animation start: ${String(animation.start)}`)
  }
  if (animation.durationMs !== undefined && (!Number.isInteger(animation.durationMs) || animation.durationMs < 1 || animation.durationMs > 60_000)) {
    fail(`text animation durationMs must be an integer between 1 and 60000: ${String(animation.durationMs)}`)
  }
  const normalized: TextAnimation = { target: animation.target.trim(), effect: animation.effect }
  if (animation.direction !== undefined) normalized.direction = animation.direction
  if (animation.byParagraph === true) normalized.byParagraph = true
  if (animation.start !== undefined) normalized.start = animation.start
  if (animation.durationMs !== undefined) normalized.durationMs = animation.durationMs
  return normalized
}

export function planSlideAnimations(entries: readonly SlideAnimationPlanEntry[]): AnimationPlan {
  const plan = new Map<number, readonly TextAnimation[]>()
  for (const entry of entries) {
    if (!Number.isInteger(entry.page) || entry.page < 1) {
      fail(`text animation page must be a positive integer: ${String(entry.page)}`)
    }
    if (entry.animations.length > TEXT_ANIMATION_MAX_PER_PAGE) {
      fail(`page ${entry.page} declares ${entry.animations.length} text animations; the limit is ${TEXT_ANIMATION_MAX_PER_PAGE}`)
    }
    const seen = new Set<string>()
    const animations = entry.animations.map(animation => {
      const normalized = normalizeTextAnimation(animation)
      const key = `${normalized.target}\u0000${normalized.byParagraph === true ? 'paragraph' : 'shape'}`
      if (seen.has(key)) fail(`page ${entry.page} declares two animations for ${normalized.target}; a target may be animated once per build unit`)
      seen.add(key)
      return normalized
    })
    plan.set(entry.page, animations)
  }
  return plan
}

interface ShapeIndex {
  idByName: Map<string, string>
  paragraphCountByName: Map<string, number>
}

/** `p:spTgt/@spid` addresses `p:cNvPr/@id`, so the id has to come from the slide part. */
function indexShapes(xml: string): ShapeIndex {
  const idByName = new Map<string, string>()
  for (const match of xml.matchAll(SHAPE_ID)) idByName.set(match[2]!, match[1]!)
  const paragraphCountByName = new Map<string, number>()
  for (const [name, id] of idByName) {
    const start = xml.indexOf(`<p:cNvPr id="${id}"`)
    if (start < 0) continue
    const end = xml.indexOf('</p:sp>', start)
    const block = xml.slice(start, end < 0 ? undefined : end)
    // pptxgenjs can emit several `a:pPr` runs inside a single `a:p`, so the
    // paragraph count is whichever of the two markers the writer used.
    const paragraphs = [...block.matchAll(/<a:p[ >]/gu)].length
    const paragraphProperties = [...block.matchAll(/<a:pPr[ >]/gu)].length
    paragraphCountByName.set(name, Math.max(paragraphs, paragraphProperties))
  }
  return { idByName, paragraphCountByName }
}

function escapeAttribute(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;').replace(/"/gu, '&quot;')
}

class TimelineBuilder {
  private nextId = 1

  id(): number {
    const value = this.nextId
    this.nextId += 1
    return value
  }
}

function targetXml(spid: string, paragraph: number | undefined): string {
  if (paragraph === undefined) return `<p:spTgt spid="${spid}"/>`
  return `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${paragraph}" end="${paragraph}"/></p:txEl></p:spTgt>`
}

function visibilitySet(builder: TimelineBuilder, target: string): string {
  return `<p:set><p:cBhvr><p:cTn id="${builder.id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
    + `<p:tgtEl>${target}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>`
    + '<p:to><p:strVal val="visible"/></p:to></p:set>'
}

function animEffectNode(builder: TimelineBuilder, target: string, filter: string, durationMs: number): string {
  return `<p:animEffect transition="in" filter="${escapeAttribute(filter)}"><p:cBhvr><p:cTn id="${builder.id()}" dur="${durationMs}"/>`
    + `<p:tgtEl>${target}</p:tgtEl></p:cBhvr></p:animEffect>`
}

/**
 * A numeric from/to is written as `p:fltVal` and a formula as `p:strVal`, because
 * PowerPoint rejects a formula it cannot parse but happily reads a literal — this
 * is exactly how it writes Zoom's zero scale itself.
 */
function animNode(builder: TimelineBuilder, target: string, motion: Motion, durationMs: number): string {
  const value = (input: string | number): string => typeof input === 'number'
    ? `<p:fltVal val="${input}"/>`
    : `<p:strVal val="${escapeAttribute(input)}"/>`
  const capability = motion.additive === true ? '<p:cBhvr additive="base">' : '<p:cBhvr>'
  return `<p:anim calcmode="lin" valueType="num">${capability}<p:cTn id="${builder.id()}" dur="${durationMs}" fill="hold"/>`
    + `<p:tgtEl>${target}</p:tgtEl><p:attrNameLst><p:attrName>${motion.attr}</p:attrName></p:attrNameLst></p:cBhvr>`
    + `<p:tavLst><p:tav tm="0"><p:val>${value(motion.from)}</p:val></p:tav>`
    + `<p:tav tm="100000"><p:val>${value(motion.to)}</p:val></p:tav></p:tavLst></p:anim>`
}

function effectXml(builder: TimelineBuilder, animation: TextAnimation, shape: ShapeIndex, start: TextAnimationStart): string {
  const spid = shape.idByName.get(animation.target)
  if (spid === undefined) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', `text animation target is not present on the slide: ${animation.target}`)
  }
  const definition = EFFECT_DEFINITIONS[animation.effect]
  const durationMs = animation.durationMs ?? TEXT_ANIMATION_DEFAULT_DURATION_MS
  const paragraphs = animation.byParagraph === true ? (shape.paragraphCountByName.get(animation.target) ?? 0) : 0
  if (animation.byParagraph === true && paragraphs < 1) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', `byParagraph animation target has no text paragraphs: ${animation.target}`)
  }

  const units: (number | undefined)[] = animation.byParagraph === true
    ? Array.from({ length: paragraphs }, (_, index) => index)
    : [undefined]

  const subtype = definition.subtypeFor !== undefined && animation.direction !== undefined
    ? definition.subtypeFor(animation.direction)
    : definition.subtype

  const buildOne = (paragraph: number | undefined): string => {
    // The effect time node is allocated before its children so the ids run
    // parent-first, matching the order PowerPoint itself writes.
    const effectId = builder.id()
    const target = targetXml(spid, paragraph)
    const parts: string[] = [visibilitySet(builder, target)]
    if (definition.nodes.kind === 'visibility') {
      // `appear` and `flash-once` are pure visibility toggles with no effect node.
    } else if (definition.nodes.kind === 'filter') {
      const filter = animation.direction === undefined
        ? definition.nodes.filter
        : definition.nodes.directional?.[animation.direction] ?? definition.nodes.filter
      parts.push(animEffectNode(builder, target, filter, durationMs))
    } else {
      // The timeline only animates the box as a whole; a per-paragraph motion
      // would need separate timing nodes per paragraph run.
      for (const motion of definition.nodes.motions(animation.direction)) {
        parts.push(animNode(builder, target, motion, durationMs))
      }
    }
    return `<p:par><p:cTn id="${effectId}" presetID="${definition.presetID}" presetClass="entr" presetSubtype="${subtype}" fill="hold" grpId="0" nodeType="__NODE__">`
      + `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${parts.join('')}</p:childTnLst></p:cTn></p:par>`
  }

  const nodeType = start === 'with-previous'
    ? 'withEffect'
    : start === 'after-previous' ? 'afterEffect' : 'clickEffect'
  if (units.length === 1) return buildOne(units[0]).replace('__NODE__', nodeType)
  // A paragraph build plays each paragraph as its own click step, so every unit
  // is a clickEffect and the whole block occupies one timeline group.
  return units.map(paragraph => buildOne(paragraph).replace('__NODE__', 'clickEffect')).join('')
}

interface PlannedAnimation {
  animation: TextAnimation
  start: TextAnimationStart
}

/** Every `on-click` opens a new timeline group; the first animation always opens one. */
function planAnimationGroups(animations: readonly TextAnimation[]): PlannedAnimation[][] {
  const groups: PlannedAnimation[][] = []
  for (const animation of animations) {
    const requested = animation.start ?? 'on-click'
    if (groups.length === 0 || requested === 'on-click') {
      // A leading `with-previous` or `after-previous` has no host, so it becomes a click step.
      groups.push([{ animation, start: 'on-click' }])
      continue
    }
    groups[groups.length - 1]!.push({ animation, start: requested })
  }
  return groups
}

export function animationTimingXml(animations: readonly TextAnimation[], shape: ShapeIndex): string {
  if (animations.length === 0) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', 'a slide animation timeline needs at least one animation')
  }
  const builder = new TimelineBuilder()
  const rootId = builder.id()
  const seqId = builder.id()
  const groups: string[] = []
  for (const group of planAnimationGroups(animations)) {
    const groupId = builder.id()
    const innerId = builder.id()
    let effects = ''
    for (const planned of group) {
      // Every effect in a group is a direct sibling and `nodeType` alone carries
      // the timing relationship. PowerPoint writes an `afterEffect` inside an
      // extra delay wrapper, but that shape is rejected on open when it is
      // hand-written (verified both ways against PowerPoint 16), while the
      // sibling form with the same `nodeType="afterEffect"` opens cleanly.
      effects += effectXml(builder, planned.animation, shape, planned.start)
    }
    // PowerPoint nests the effect list two levels deep: group -> inner par -> effects.
    groups.push(`<p:par><p:cTn id="${groupId}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>`
      + `<p:par><p:cTn id="${innerId}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>`
      + '</p:childTnLst></p:cTn></p:par>')
  }
  return '<p:timing><p:tnLst><p:par>'
    + `<p:cTn id="${rootId}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`
    + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${seqId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groups.join('')}</p:childTnLst></p:cTn>`
    + '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
    + '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
    + '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
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
function timingInsertionOffset(body: string): number {
  const csldEnd = elementEndOffset(body, 'p:cSld')
  if (csldEnd < 0) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML has no p:cSld element')
  const clrMapEnd = elementEndOffset(body, 'p:clrMapOvr')
  const transitionEnd = elementEndOffset(body, 'p:transition')
  return Math.max(csldEnd, clrMapEnd, transitionEnd)
}

export function injectSlideAnimation(xml: string, animations: readonly TextAnimation[]): string {
  const root = /<p:sld\b[^>]*>/u.exec(xml)
  if (root === null) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML does not contain a p:sld root element')
  const bodyStart = root.index + root[0].length
  const bodyEnd = xml.lastIndexOf('</p:sld>')
  if (bodyEnd < bodyStart) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML has an unterminated p:sld root element')
  const body = xml.slice(bodyStart, bodyEnd)
  const withoutTiming = body.replace(/<p:timing\b[\s\S]*?<\/p:timing>/u, '')
  if (withoutTiming.includes('<p:timing')) {
    throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'slide XML has an unterminated p:timing element')
  }
  const payload = animationTimingXml(animations, indexShapes(xml))
  const namespaced = /xmlns:p=/u.test(root[0])
    ? payload
    : payload.replace('<p:timing', `<p:timing xmlns:p="${PRESENTATION_NAMESPACE}"`)
  const at = bodyStart + timingInsertionOffset(withoutTiming)
  const rebuilt = `${withoutTiming.slice(0, at)}${namespaced}${withoutTiming.slice(at)}`
  return `${xml.slice(0, bodyStart)}${rebuilt}${xml.slice(bodyEnd)}`
}

export function rewritePptxAnimations(data: Uint8Array, plan: AnimationPlan): Uint8Array {
  if (plan.size === 0) return data
  const targets = new Map<number, readonly TextAnimation[]>()
  for (const [page, animations] of plan) {
    if (!Number.isInteger(page) || page < 1) fail(`text animation page must be a positive integer: ${String(page)}`)
    targets.set(page, animations)
  }
  const injected = new Set<number>()
  const output = rewriteZip(data, {
    transform: (name, read) => {
      const match = SLIDE_PART.exec(name)
      const page = match === null ? undefined : Number(match[1])
      const animations = page === undefined ? undefined : targets.get(page)
      if (page === undefined || animations === undefined) return undefined
      injected.add(page)
      return strToU8(injectSlideAnimation(strFromU8(read()), animations))
    },
  })
  for (const page of targets.keys()) {
    if (!injected.has(page)) fail(`no slide part matches text animation page ${page}`)
  }
  return output
}
