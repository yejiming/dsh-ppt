/**
 * Slide taxonomy shared by the outline schema, the art direction schema and the
 * built-in theme library. It lives in its own module because all three need it:
 * keeping it in `outline.ts` made `outline.ts -> themes.ts -> outline.ts` a
 * cycle, which breaks the moment any of those modules is loaded first.
 */

export const SLIDE_TYPES = ['cover', 'agenda', 'section', 'content', 'comparison', 'timeline', 'process', 'data', 'quote', 'summary', 'ending'] as const

export const SLIDE_LAYOUTS = [
  'cover', 'center', 'title-content', 'split', 'two-column', 'three-column', 'grid', 'hero-image',
  'image-left', 'image-right', 'timeline-horizontal', 'timeline-vertical', 'process-horizontal',
  'process-vertical', 'chart-focus', 'quote-focus', 'full-bleed', 'closing',
] as const

export type SlideType = typeof SLIDE_TYPES[number]
export type SlideLayout = typeof SLIDE_LAYOUTS[number]

/**
 * Which layouts can carry which slide roles. A layout that appears here for the
 * wrong role is rejected before anything is rendered, so this table is the
 * single authority behind both the schema and the model-facing layout guide.
 */
export const SLIDE_TYPE_LAYOUTS: Record<SlideLayout, readonly SlideType[]> = {
  cover: ['cover'],
  center: ['cover', 'section', 'quote', 'ending'],
  'title-content': ['agenda', 'content', 'summary'],
  split: ['content', 'comparison', 'data'],
  'two-column': ['agenda', 'content', 'comparison', 'data', 'summary'],
  'three-column': ['agenda', 'content', 'summary'],
  grid: ['agenda', 'content', 'data', 'summary'],
  'hero-image': ['cover', 'section', 'content'],
  'image-left': ['content', 'quote'],
  'image-right': ['content', 'quote'],
  'timeline-horizontal': ['timeline'],
  'timeline-vertical': ['timeline'],
  'process-horizontal': ['process'],
  'process-vertical': ['process'],
  'chart-focus': ['data'],
  'quote-focus': ['quote'],
  'full-bleed': ['cover', 'section', 'quote', 'ending'],
  closing: ['ending'],
}

/** Slide roles that may legitimately carry no visible content item. */
export const CONTENT_OPTIONAL_TYPES: readonly SlideType[] = ['cover', 'section', 'ending']

/** Slide roles whose layout must be compatible with an ordered point sequence. */
export const SEQUENCED_TYPES: readonly SlideType[] = ['timeline', 'process']
