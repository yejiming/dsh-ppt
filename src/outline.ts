import { isAbsolute } from 'node:path'
import { rm } from 'node:fs/promises'
import { z } from 'zod'
import { allocateArtifactDirectory } from './artifacts.ts'
import { artDirectionFindings, validateArtDirection } from './art-direction.ts'
import type { ArtDirection } from './art-direction.ts'
import { atomicWriteJson } from './atomic.ts'
import { PptError } from './errors.ts'
import { resolveRegisteredFont, type DiscoveredFont } from './fonts.ts'
import { workspaceRelative } from './paths.ts'
import { findTheme, themeFindingsForPlan, validateTheme, type PptTheme } from './themes.ts'

import { CONTENT_OPTIONAL_TYPES, SEQUENCED_TYPES, SLIDE_LAYOUTS, SLIDE_TYPE_LAYOUTS, SLIDE_TYPES } from './slide-taxonomy.ts'

// Re-exported so existing importers keep one path to the taxonomy.
export { SLIDE_LAYOUTS, SLIDE_TYPES } from './slide-taxonomy.ts'

const noMarkup = (value: string): boolean => !/[\r\n]/u.test(value) && !/<\/?[a-z][^>]*>/iu.test(value)

function cleanString(max: number, multiline = false) {
  return z.string().transform(value => value.normalize('NFC').trim()).pipe(
    z.string().refine(value => [...value].length >= 1 && [...value].length <= max, `must contain 1..${max} Unicode code points`)
      .refine(value => multiline || noMarkup(value), 'must not contain newlines or HTML'),
  )
}

function safeReference(max: number) {
  return cleanString(max).refine((value) => {
    try {
      const url = new URL(value)
      return (url.protocol === 'http:' || url.protocol === 'https:') && url.username === '' && url.password === ''
    } catch {
      if (isAbsolute(value) || value.includes('\\')) return false
      const segments = value.split('/')
      return segments.every(segment => segment !== '..' && segment !== '.' && segment.length > 0)
    }
  }, 'must be a public HTTP(S) URL or safe workspace-relative path')
}

const Point = z.strictObject({
  kind: z.literal('point'), text: cleanString(180), label: cleanString(40).optional(), group: cleanString(40).optional(),
  level: z.union([z.literal(1), z.literal(2)]).default(1), emphasis: z.boolean().default(false),
})

const Data = z.strictObject({
  kind: z.literal('data'), label: cleanString(60),
  value: z.union([z.number().finite(), cleanString(40)]), unit: cleanString(20).optional(),
  source: safeReference(500).optional(), note: cleanString(120).optional(), group: cleanString(40).optional(),
  emphasis: z.boolean().default(false),
})

const Image = z.strictObject({
  kind: z.literal('image'), role: z.enum(['hero', 'supporting', 'background', 'portrait', 'logo', 'diagram']),
  intent: cleanString(160), query: cleanString(160).optional(), asset: safeReference(240).optional(),
  caption: cleanString(120).optional(), group: cleanString(40).optional(),
}).superRefine((item, context) => {
  if ((item.query === undefined) === (item.asset === undefined)) context.addIssue({ code: 'custom', message: 'image requires exactly one of query or asset' })
  if (item.role === 'background' && item.caption !== undefined) context.addIssue({ code: 'custom', path: ['caption'], message: 'background images cannot have captions' })
})

const Chart = z.strictObject({
  kind: z.literal('chart'), chart_type: z.enum(['bar', 'line', 'area', 'pie', 'donut', 'scatter', 'bubble', 'radar', 'waterfall', 'funnel', 'heatmap', 'table']),
  subject: cleanString(120), data_ref: safeReference(240).optional(), takeaway: cleanString(180), group: cleanString(40).optional(),
})

const Note = z.strictObject({
  kind: z.literal('note'), purpose: z.enum(['speaker', 'production']), text: cleanString(500, true),
})

export const OutlineContentItemSchema = z.discriminatedUnion('kind', [Point, Data, Image, Chart, Note])

const Style = z.strictObject({
  layout: z.enum(SLIDE_LAYOUTS), background: z.enum(['light', 'dark', 'accent', 'image']),
  accent: z.string().regex(/^#[0-9A-Fa-f]{6}$/u).transform(value => value.toUpperCase()),
  // Registry membership is enforced later by resolveFontPlan, which substitutes
  // a deterministic fallback and reports it. Keeping the schema open means a
  // registry newer than the running host degrades instead of failing outright.
  title_font: cleanString(80), body_font: cleanString(80), visual_direction: cleanString(200),
})

const Slide = z.strictObject({
  page: z.number().int().min(1).max(60), type: z.enum(SLIDE_TYPES), title: cleanString(80),
  content: z.array(OutlineContentItemSchema).max(12), style: Style,
}).superRefine((slide, context) => {
  const visible = slide.content.filter(item => item.kind !== 'note')
  const notes = slide.content.filter(item => item.kind === 'note')
  if (visible.length > 8) context.addIssue({ code: 'custom', path: ['content'], message: 'a slide can contain at most 8 visible items' })
  if (notes.length > 2) context.addIssue({ code: 'custom', path: ['content'], message: 'a slide can contain at most 2 notes' })
  if (![...CONTENT_OPTIONAL_TYPES].includes(slide.type as 'cover') && visible.length === 0) {
    context.addIssue({ code: 'custom', path: ['content'], message: `${slide.type} requires at least one visible item` })
  }
  const titleLimit = slide.type === 'cover' ? 80 : 60
  if ([...slide.title].length > titleLimit) context.addIssue({ code: 'custom', path: ['title'], message: `${slide.type} title exceeds ${titleLimit} code points` })

  if (!SLIDE_TYPE_LAYOUTS[slide.style.layout].includes(slide.type)) {
    context.addIssue({ code: 'custom', path: ['style', 'layout'], message: `${slide.style.layout} is incompatible with ${slide.type}` })
  }
  const images = slide.content.filter(item => item.kind === 'image')
  const backgrounds = images.filter(item => item.role === 'background')
  if (slide.style.background === 'image' && backgrounds.length !== 1) context.addIssue({ code: 'custom', path: ['content'], message: 'image background requires exactly one background image item' })
  if (slide.style.background !== 'image' && backgrounds.length > 0) context.addIssue({ code: 'custom', path: ['content'], message: 'background image item requires style.background=image' })
  if (['hero-image', 'image-left', 'image-right'].includes(slide.style.layout) && images.every(item => item.role === 'background')) {
    context.addIssue({ code: 'custom', path: ['content'], message: `${slide.style.layout} requires a non-background image` })
  }
  const charts = slide.content.filter(item => item.kind === 'chart')
  if (slide.style.layout === 'chart-focus' && charts.length !== 1) context.addIssue({ code: 'custom', path: ['content'], message: 'chart-focus requires exactly one chart' })
  if (slide.type === 'data' && slide.style.layout !== 'chart-focus' && charts.length > 2) context.addIssue({ code: 'custom', path: ['content'], message: 'data slides allow at most two charts' })
  for (const [index, chart] of slide.content.entries()) {
    if (chart.kind !== 'chart' || chart.data_ref !== undefined) continue
    const dataReady = slide.content.some(item => item.kind === 'data')
    const pending = slide.content.some(item => item.kind === 'note' && item.purpose === 'production' && /(?:待补.*数据|data.*pending)/iu.test(item.text))
    if (!dataReady && !pending) context.addIssue({ code: 'custom', path: ['content', index, 'data_ref'], message: 'chart without data_ref requires a data item or explicit pending-data production note' })
  }
  if (slide.type === 'comparison') {
    const groups = new Set(visible.flatMap(item => 'group' in item && item.group !== undefined ? [item.group] : []))
    if (groups.size < 2) context.addIssue({ code: 'custom', path: ['content'], message: 'comparison requires at least two explicit groups' })
  }
  if (SEQUENCED_TYPES.includes(slide.type)) {
    const points = slide.content.filter(item => item.kind === 'point').length
    if (points < 3 || points > 8) context.addIssue({ code: 'custom', path: ['content'], message: `${slide.type} requires 3..8 point items` })
  }
})

export const PptOutlineSchema = z.array(Slide).min(1).max(60).superRefine((slides, context) => {
  slides.forEach((slide, index) => {
    if (slide.page !== index + 1) context.addIssue({ code: 'custom', path: [index, 'page'], message: `page must be ${index + 1}` })
  })
})

export type PptOutline = z.infer<typeof PptOutlineSchema>

export interface OutlineWriteResult {
  artifact_dir: string
  outline_path: string
  design_plan_path?: string
  design_status: 'directed' | 'legacy'
  page_count: number
  type_counts: Record<string, number>
  fonts: string[]
  warnings: string[]
  blocking_warnings: string[]
  /** Present only when the caller pinned the deck to a built-in theme. */
  theme?: {
    id: string
    name: string
    palette_source: string
    accent: string
    accent_inverted: string
    findings: string[]
  }
}

/**
 * Check that an authored outline and plan still belong to the theme they name.
 * Selecting a theme and then using unrelated colours or fonts is the failure
 * this catches: the deck reads as neither the theme nor the brief.
 */
export function themeConformanceFindings(
  theme: PptTheme,
  outline: PptOutline,
  designPlan: ArtDirection | undefined,
): string[] {
  const findings: string[] = []
  if (designPlan === undefined) {
    findings.push(`THEME_PLAN_MISSING: theme ${theme.id} was selected but no art_direction was supplied, so nothing enforces the theme`)
  } else {
    for (const finding of themeFindingsForPlan(theme, designPlan)) findings.push(`${finding.code}: ${finding.message}`)
  }
  const palette = new Set([
    ...theme.palette.background, ...theme.palette.surface, ...theme.palette.text,
    theme.palette.accent, theme.palette.accent_inverted,
  ])
  const families = new Set(Object.values(theme.typography).map(role => role.family))
  for (const slide of outline) {
    if (!palette.has(slide.style.accent)) {
      findings.push(`THEME_ACCENT_DRIFT (page ${slide.page}): style.accent ${slide.style.accent} is not defined by theme ${theme.id}`)
    }
    for (const [role, family] of [['title_font', slide.style.title_font], ['body_font', slide.style.body_font]] as const) {
      if (!families.has(family)) {
        findings.push(`THEME_FONT_DRIFT (page ${slide.page}, ${role}): ${family} is not one of theme ${theme.id} typography families`)
      }
    }
  }
  return findings
}

export function resolveTheme(themeId: string): PptTheme {
  return validateTheme(findTheme(themeId))
}

export interface OutlineFontResolutionOptions {
  discovered: readonly DiscoveredFont[]
  platform?: NodeJS.Platform
}

function bodyText(slide: PptOutline[number]): string {
  const values: string[] = []
  for (const item of slide.content) {
    if (item.kind === 'point') values.push(item.label ?? '', item.text)
    else if (item.kind === 'data') values.push(item.label, String(item.value), item.unit ?? '', item.note ?? '')
    else if (item.kind === 'image') values.push(item.caption ?? '', item.intent)
    else if (item.kind === 'chart') values.push(item.subject, item.takeaway)
  }
  return values.filter(Boolean).join(' ') || slide.title
}

function resolveFontPlan(
  outline: PptOutline,
  designPlan: ArtDirection | undefined,
  options: OutlineFontResolutionOptions | undefined,
): { outline: PptOutline; designPlan?: ArtDirection; warnings: string[] } {
  if (options === undefined) return { outline, ...(designPlan === undefined ? {} : { designPlan }), warnings: [] }
  const resolvedOutline = structuredClone(outline)
  const resolvedDesign = designPlan === undefined ? undefined : structuredClone(designPlan)
  const warnings = new Set<string>()
  for (const slide of resolvedOutline) {
    const title = resolveRegisteredFont(slide.style.title_font, slide.title, options.discovered, options.platform)
    const body = resolveRegisteredFont(slide.style.body_font, bodyText(slide), options.discovered, options.platform)
    slide.style.title_font = title.resolved.name as typeof slide.style.title_font
    slide.style.body_font = body.resolved.name as typeof slide.style.body_font
    if (title.warning !== undefined) warnings.add(`FONT_FALLBACK (page ${slide.page}, title): ${title.warning}`)
    if (body.warning !== undefined) warnings.add(`FONT_FALLBACK (page ${slide.page}, body): ${body.warning}`)
  }
  if (resolvedDesign !== undefined) {
    const titleSample = resolvedOutline.map(slide => slide.title).join(' ')
    const bodySample = resolvedOutline.map(bodyText).join(' ')
    const samples: Record<keyof ArtDirection['typography'], string> = {
      display: titleSample, body: bodySample, latin: 'AaZz09', code: 'AaZz09_{}[]();',
    }
    for (const role of Object.keys(samples) as Array<keyof ArtDirection['typography']>) {
      const requested = resolvedDesign.typography[role].family
      const resolved = resolveRegisteredFont(requested, samples[role], options.discovered, options.platform)
      resolvedDesign.typography[role].family = resolved.resolved.name as typeof requested
      if (resolved.warning !== undefined) warnings.add(`FONT_FALLBACK (art_direction.${role}): ${resolved.warning}`)
    }
  }
  return { outline: resolvedOutline, ...(resolvedDesign === undefined ? {} : { designPlan: resolvedDesign }), warnings: [...warnings] }
}

/**
 * Compress zod issues into one line a caller can act on. The hosting surface
 * renders `PptError.message` but not its `details`, so the offending JSON paths
 * must reach the message itself.
 */
function describeOutlineIssues(issues: readonly { path: readonly PropertyKey[]; message: string }[]): string {
  const rendered = issues.map(issue => {
    const path = issue.path.length === 0 ? '<root>' : issue.path.join('.')
    const message = issue.message.length > 200 ? `${issue.message.slice(0, 197)}...` : issue.message
    return `${path}: ${message}`
  })
  const shown = rendered.slice(0, 8)
  return rendered.length > shown.length ? `${shown.join('; ')}; +${rendered.length - shown.length} more` : shown.join('; ')
}

export function validatePptOutline(value: unknown): PptOutline {
  const result = PptOutlineSchema.safeParse(value)
  if (result.success) return result.data
  const issues = result.error.issues.map(issue => ({ path: issue.path, message: issue.message }))
  throw new PptError('PPT_OUTLINE_INVALID', `PPT outline validation failed: ${describeOutlineIssues(issues)}`, {
    details: { issues: issues.map(issue => ({ path: issue.path.join('.'), message: issue.message })) },
  })
}

export async function writePptOutline(
  workspace: string,
  artifactTitle: string,
  value: unknown,
  outputRoot = 'ppt-output',
  signal?: AbortSignal,
  artDirection?: unknown,
  fontResolution?: OutlineFontResolutionOptions,
  themeId?: string,
): Promise<OutlineWriteResult> {
  const validatedOutline = validatePptOutline(value)
  const validatedDesignPlan = artDirection === undefined ? undefined : validateArtDirection(artDirection, validatedOutline.length)
  const theme = themeId === undefined ? undefined : resolveTheme(themeId)
  const resolved = resolveFontPlan(validatedOutline, validatedDesignPlan, fontResolution)
  const outline = resolved.outline
  const designPlan = resolved.designPlan
  const paths = await allocateArtifactDirectory(workspace, artifactTitle, outputRoot)
  try {
    await atomicWriteJson(paths.outline, outline, { signal })
    if (designPlan !== undefined) await atomicWriteJson(paths.designPlan, designPlan, { signal })
  } catch (error) {
    await Promise.all([rm(paths.outline, { force: true }), rm(paths.designPlan, { force: true })])
    throw error
  }
  const typeCounts: Record<string, number> = {}
  const fonts = new Set<string>()
  const blocking: string[] = []
  for (const slide of outline) {
    typeCounts[slide.type] = (typeCounts[slide.type] ?? 0) + 1
    fonts.add(slide.style.title_font)
    fonts.add(slide.style.body_font)
    if (slide.content.some(item => item.kind === 'chart' && item.data_ref === undefined && !slide.content.some(other => other.kind === 'data'))) {
      blocking.push(`page ${slide.page}: chart data is explicitly pending`)
    }
  }
  if (designPlan !== undefined) for (const role of Object.values(designPlan.typography)) fonts.add(role.family)
  const themeFindings = theme === undefined ? [] : themeConformanceFindings(theme, outline, designPlan)
  return {
    artifact_dir: workspaceRelative(workspace, paths.root), outline_path: workspaceRelative(workspace, paths.outline),
    ...(designPlan === undefined ? {} : { design_plan_path: workspaceRelative(workspace, paths.designPlan) }),
    design_status: designPlan === undefined ? 'legacy' : 'directed',
    page_count: outline.length, type_counts: typeCounts, fonts: [...fonts].sort(),
    warnings: [
      ...resolved.warnings,
      ...(designPlan === undefined
        ? ['ART_DIRECTION_MISSING: legacy outline created without design-plan.json']
        : artDirectionFindings(designPlan).map(finding => `${finding.code}${finding.page === undefined ? '' : ` (page ${finding.page})`}: ${finding.message}`)),
      ...themeFindings,
    ],
    blocking_warnings: blocking,
    ...(theme === undefined ? {} : {
      theme: {
        id: theme.id, name: theme.name, palette_source: theme.palette_source,
        accent: theme.palette.accent, accent_inverted: theme.palette.accent_inverted,
        findings: themeFindings,
      },
    }),
  }
}
