import { z } from 'zod'
import { PptError } from './errors.ts'

export const ART_COMPOSITIONS = [
  'hero', 'editorial-split', 'asymmetric-split', 'process', 'layered', 'data-focus', 'quote', 'full-bleed', 'closing',
] as const
export const ART_DENSITIES = ['low', 'medium', 'high'] as const
export const ART_BACKGROUNDS = ['base', 'inverse', 'accent', 'image'] as const
export const ART_TITLE_TREATMENTS = ['statement', 'question', 'label', 'number-led'] as const
export const ART_ANCHOR_KINDS = ['none', 'typography', 'image', 'data', 'code', 'diagram'] as const
export const ART_FRAME_POLICIES = ['none', 'single', 'grouped'] as const
export const ART_ROLES = ['title', 'subtitle', 'body', 'metric', 'code', 'diagram', 'visual-anchor', 'supporting', 'frame'] as const

const text = (max: number) => z.string().transform(value => value.normalize('NFC').trim()).pipe(
  z.string().refine(value => [...value].length >= 1 && [...value].length <= max, `must contain 1..${max} Unicode code points`)
    .refine(value => !/[\r\n]/u.test(value) && !/<\/?[a-z][^>]*>/iu.test(value), 'must not contain newlines or HTML'),
)
const color = z.string().regex(/^#[0-9A-Fa-f]{6}$/u).transform(value => value.toUpperCase())
// Font families are checked against the registry after parsing (resolveFontPlan
// does it deterministically), not at the schema boundary. A registry newer than
// the running host must degrade to a fallback, never reject a valid plan.
const fontRole = z.strictObject({ family: text(120), weight: z.number().int().min(100).max(900) })

const visualAnchor = z.strictObject({
  kind: z.enum(ART_ANCHOR_KINDS),
  role: text(80),
  min_area_ratio: z.number().finite().min(0.05).max(0.9).optional(),
}).superRefine((value, context) => {
  if (value.kind === 'none' && value.min_area_ratio !== undefined) context.addIssue({ code: 'custom', path: ['min_area_ratio'], message: 'none anchor cannot define min_area_ratio' })
  if (value.kind !== 'none' && value.min_area_ratio === undefined) context.addIssue({ code: 'custom', path: ['min_area_ratio'], message: 'visual anchor requires min_area_ratio' })
})

const artSlide = z.strictObject({
  page: z.number().int().min(1).max(60),
  job: text(160),
  takeaway: text(180),
  composition: z.enum(ART_COMPOSITIONS),
  density: z.enum(ART_DENSITIES),
  background_role: z.enum(ART_BACKGROUNDS),
  title_treatment: z.enum(ART_TITLE_TREATMENTS),
  visual_anchor: visualAnchor,
  frame_policy: z.enum(ART_FRAME_POLICIES),
  allow_intentional_repeat: z.boolean().default(false),
})

export const ArtDirectionSchema = z.strictObject({
  version: z.literal(1).default(1),
  concept: text(100),
  audience_effect: text(180),
  palette: z.strictObject({
    background: z.array(color).min(1).max(4),
    surface: z.array(color).min(1).max(4),
    accent: color,
    text: z.array(color).min(1).max(4),
  }),
  typography: z.strictObject({
    display: fontRole,
    body: fontRole,
    latin: fontRole,
    code: fontRole,
  }),
  rhythm: z.strictObject({
    background_sequence: z.array(z.enum(ART_BACKGROUNDS)).min(1).max(60),
    max_grouped_frame_slides: z.number().int().min(0).max(60).default(2),
    max_same_composition_run: z.number().int().min(1).max(6).default(1),
  }),
  slides: z.array(artSlide).min(1).max(60),
}).superRefine((plan, context) => {
  if (plan.rhythm.background_sequence.length !== plan.slides.length) {
    context.addIssue({ code: 'custom', path: ['rhythm', 'background_sequence'], message: 'background_sequence must contain one entry per slide' })
  }
  plan.slides.forEach((slide, index) => {
    if (slide.page !== index + 1) context.addIssue({ code: 'custom', path: ['slides', index, 'page'], message: `page must be ${index + 1}` })
    if (plan.rhythm.background_sequence[index] !== slide.background_role) {
      context.addIssue({ code: 'custom', path: ['slides', index, 'background_role'], message: 'background_role must match rhythm.background_sequence' })
    }
  })
})

export type ArtDirection = z.infer<typeof ArtDirectionSchema>

export interface DesignFinding {
  code: string
  severity: 'warning' | 'error'
  message: string
  page?: number
}

export function validateArtDirection(value: unknown, expectedPages?: number): ArtDirection {
  const result = ArtDirectionSchema.safeParse(value)
  if (!result.success) {
    throw new PptError('PPT_ART_DIRECTION_INVALID', 'PPT art direction validation failed', {
      details: { issues: result.error.issues.map(issue => ({ path: issue.path.join('.'), message: issue.message })) },
    })
  }
  if (expectedPages !== undefined && result.data.slides.length !== expectedPages) {
    throw new PptError('PPT_ART_DIRECTION_INVALID', 'PPT art direction page count does not match outline', {
      details: { issues: [{ path: 'slides', message: `expected ${expectedPages} pages, received ${result.data.slides.length}` }] },
    })
  }
  return result.data
}

export function artDirectionFindings(plan: ArtDirection): DesignFinding[] {
  const findings: DesignFinding[] = []
  const grouped = plan.slides.filter(slide => slide.frame_policy === 'grouped')
  if (grouped.length > plan.rhythm.max_grouped_frame_slides) {
    findings.push({
      code: 'ART_GROUPED_FRAMES_OVER_BUDGET', severity: 'warning',
      message: `${grouped.length} grouped-frame slides exceed the planned maximum of ${plan.rhythm.max_grouped_frame_slides}`,
      page: grouped[plan.rhythm.max_grouped_frame_slides]?.page,
    })
  }
  let run = 1
  for (let index = 1; index < plan.slides.length; index += 1) {
    const current = plan.slides[index]!
    const previous = plan.slides[index - 1]!
    run = current.composition === previous.composition ? run + 1 : 1
    if (run > plan.rhythm.max_same_composition_run && !current.allow_intentional_repeat) {
      findings.push({ code: 'ART_COMPOSITION_REPEATED', severity: 'warning', message: `composition ${current.composition} repeats beyond the planned run length`, page: current.page })
    }
  }
  if (plan.slides.length >= 6 && new Set(plan.slides.map(slide => slide.composition)).size < 3) {
    findings.push({ code: 'ART_COMPOSITION_VARIETY_LOW', severity: 'warning', message: 'decks with six or more slides should use at least three composition families' })
  }
  return findings
}

export function artDirectionReviewChecklist(plan: ArtDirection): string[] {
  const checklist = [
    `Confirm the deck expresses the art direction concept “${plan.concept}” and intended audience effect “${plan.audience_effect}”.`,
    'Compare adjacent slides for deliberate rhythm rather than accidental repetition of silhouette, density, background, or card structure.',
  ]
  for (const slide of plan.slides) {
    const anchor = slide.visual_anchor.kind === 'none' ? 'no visual anchor' : `${slide.visual_anchor.kind} visual anchor`
    checklist.push(`Page ${slide.page}: verify ${slide.composition}, ${slide.density} density, ${slide.background_role} background, ${anchor}, and ${slide.frame_policy} frame policy deliver “${slide.takeaway}”.`)
  }
  return checklist.slice(0, 30)
}
