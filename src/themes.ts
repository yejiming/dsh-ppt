import { ART_BACKGROUNDS, ART_COMPOSITIONS, ART_DENSITIES, ART_FRAME_POLICIES, ART_TITLE_TREATMENTS, type ArtDirection } from './art-direction.ts'
import { PptError } from './errors.ts'
import { registeredFont } from './fonts.ts'
import { SLIDE_TYPES } from './slide-taxonomy.ts'

/**
 * A deck theme is a reusable design recipe: colours, type roles, a page rhythm
 * and the decoration vocabulary that holds them together. The plugin shipped
 * without one, so every deck was designed from nothing and the quality ceiling
 * was whatever the model improvised that turn. A theme does not replace the
 * agent's design pass — it gives the pass a vetted starting point and a
 * verifiable palette, the way a template does for a human designer.
 */

export type ArtComposition = typeof ART_COMPOSITIONS[number]
export type ArtBackground = typeof ART_BACKGROUNDS[number]
export type ArtDensity = typeof ART_DENSITIES[number]
export type ArtTitleTreatment = typeof ART_TITLE_TREATMENTS[number]
export type ArtFramePolicy = typeof ART_FRAME_POLICIES[number]

export interface ThemeFontRole {
  family: string
  weight: number
}

export interface PptTheme {
  id: string
  /** Display name shown to the user when the theme is offered. */
  name: string
  concept: string
  audience_effect: string
  scenes: readonly string[]
  /**
   * Where the palette values come from. Every built-in theme traces to a
   * published design system so a colour can be audited instead of trusted; a
   * theme authored from a brand book or a public palette site records that
   * source here too.
   */
  palette_source: string
  palette: {
    background: readonly string[]
    surface: readonly string[]
    accent: string
    /**
     * Accent for the inverted tonal group. A single accent cannot clear 4.5:1 on
     * both a light and a dark field, so every theme carries the pair and the plan
     * tells the model which one each page uses.
     */
    accent_inverted: string
    text: readonly string[]
  }
  typography: {
    display: ThemeFontRole
    body: ThemeFontRole
    latin: ThemeFontRole
    code: ThemeFontRole
  }
  /** Composition preference, repeated across the deck in page order. */
  composition_cycle: readonly ArtComposition[]
  /** Background rhythm preference, repeated across the deck in page order. */
  background_cycle: readonly ArtBackground[]
  /** Rhythm limits the theme's own decoration was designed against. */
  rhythm_limits: {
    max_grouped_frame_slides: number
    max_same_composition_run: number
  }
  /** Decoration moves a page may use; every entry must survive the HTML whitelist. */
  decoration: readonly string[]
  /** What each composition means inside this theme. */
  layout_notes: Readonly<Partial<Record<ArtComposition, string>>>
}

const HEX = /^#[0-9A-F]{6}$/u

/** WCAG 2.1 relative luminance of an #RRGGBB colour. */
export function relativeLuminance(color: string): number {
  const hex = color.replace('#', '')
  const channel = (offset: number): number => {
    const value = Number.parseInt(hex.slice(offset, offset + 2), 16) / 255
    return value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4
  }
  return 0.2126 * channel(0) + 0.7152 * channel(2) + 0.0722 * channel(4)
}

/** WCAG 2.1 contrast ratio between two #RRGGBB colours, from 1 to 21. */
export function contrastRatio(a: string, b: string): number {
  const first = relativeLuminance(a)
  const second = relativeLuminance(b)
  const lighter = Math.max(first, second)
  const darker = Math.min(first, second)
  return (lighter + 0.05) / (darker + 0.05)
}

export interface ThemeFinding {
  code: string
  severity: 'warning' | 'error'
  message: string
}

/**
 * Contrast is the one part of "looks good" that can be checked by machine, so
 * every theme is held to it at load time rather than discovered in review.
 *
 * A palette is a set of tonal groups, not a cross product: `background[i]`,
 * `surface[i]` and `text[i]` belong together, so text is paired by index. The
 * accent is a single value that must stay readable on the primary group; on an
 * inverted group it is expected to be swapped for the theme's inverted accent,
 * which is reported as a warning rather than a failure.
 */
export function themeFindings(theme: PptTheme): ThemeFinding[] {
  const findings: ThemeFinding[] = []
  const check = (label: string, foreground: string, background: string, minimum: number): void => {
    const ratio = contrastRatio(foreground, background)
    if (ratio < minimum) {
      findings.push({
        code: 'THEME_CONTRAST_LOW', severity: 'error',
        message: `${theme.id}: ${label} contrast ${ratio.toFixed(2)}:1 is below ${minimum}:1 (${foreground} on ${background})`,
      })
    }
  }
  const groups = Math.max(theme.palette.background.length, theme.palette.surface.length, theme.palette.text.length)
  for (let index = 0; index < groups; index += 1) {
    const background = theme.palette.background[index % theme.palette.background.length]!
    const surface = theme.palette.surface[index % theme.palette.surface.length]!
    const text = theme.palette.text[index % theme.palette.text.length]!
    check(`body text ${text} on background ${background}`, text, background, 4.5)
    check(`body text ${text} on surface ${surface}`, text, surface, 4.5)
    if (index === 0) {
      check(`accent on background ${background}`, theme.palette.accent, background, 4.5)
      check(`accent on surface ${surface}`, theme.palette.accent, surface, 4.5)
      continue
    }
    check(`inverted accent on background ${background}`, theme.palette.accent_inverted, background, 4.5)
    check(`inverted accent on surface ${surface}`, theme.palette.accent_inverted, surface, 4.5)
  }
  return findings
}

function requireColor(theme: string, label: string, value: string): void {
  if (!HEX.test(value)) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme} ${label} must be #RRGGBB: ${value}`)
  }
}

function requireFont(theme: string, role: string, family: string): void {
  if (registeredFont(family) === undefined) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme} ${role} font is not in the approved registry: ${family}`)
  }
}

export function validateTheme(theme: PptTheme): PptTheme {
  if (!/^[a-z0-9][a-z0-9-]{1,40}$/u.test(theme.id)) {
    throw new PptError('PPT_THEME_INVALID', `theme id must be lower-kebab-case: ${theme.id}`)
  }
  if (theme.palette.background.length === 0 || theme.palette.background.length > 4) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} needs 1..4 background colours`)
  }
  if (theme.palette.surface.length === 0 || theme.palette.surface.length > 4) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} needs 1..4 surface colours`)
  }
  if (theme.palette.text.length === 0 || theme.palette.text.length > 4) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} needs 1..4 text colours`)
  }
  theme.palette.background.forEach((value, index) => requireColor(theme.id, `background[${index}]`, value))
  theme.palette.surface.forEach((value, index) => requireColor(theme.id, `surface[${index}]`, value))
  theme.palette.text.forEach((value, index) => requireColor(theme.id, `text[${index}]`, value))
  requireColor(theme.id, 'accent', theme.palette.accent)
  for (const role of ['display', 'body', 'latin', 'code'] as const) {
    const entry = theme.typography[role]
    requireFont(theme.id, role, entry.family)
    if (!Number.isInteger(entry.weight) || entry.weight < 100 || entry.weight > 900 || entry.weight % 100 !== 0) {
      throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} ${role} weight must be a 100..900 multiple of 100: ${entry.weight}`)
    }
  }
  if (theme.composition_cycle.length === 0) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} needs at least one composition`)
  }
  if (theme.background_cycle.length === 0) {
    throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} needs at least one background role`)
  }
  for (const composition of theme.composition_cycle) {
    if (!(ART_COMPOSITIONS as readonly string[]).includes(composition)) {
      throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} has unknown composition: ${composition}`)
    }
  }
  for (const background of theme.background_cycle) {
    if (!(ART_BACKGROUNDS as readonly string[]).includes(background)) {
      throw new PptError('PPT_THEME_INVALID', `theme ${theme.id} has unknown background role: ${background}`)
    }
  }
  return theme
}

export interface ThemeSummary {
  id: string
  name: string
  concept: string
  /** Copied out of the theme so the catalogue is plain mutable JSON for the tool wire. */
  scenes: string[]
  palette_source: string
  accent: string
  accent_inverted: string
  display_font: string
  body_font: string
  signature: string
}

export interface ThemeCatalog {
  themes: ThemeSummary[]
  usage: string
  warnings: string[]
}

const USAGE = [
  'ppt_themes returns reusable design recipes. Call it with no arguments to choose one, then pass the chosen theme_id to ppt_outline.',
  'A theme is a starting point for the Art Direction pass, not a substitute for it: keep concept and audience_effect specific to this deck, and write each page job and takeaway from the actual content.',
  'Call ppt_themes with theme_id and page_types to get the visual half of the plan for every page: composition, density, background role, title treatment, frame policy and the exact colours that page should use.',
  'layout_notes and decoration say how each composition is built and which ornament this theme permits, so the HTML follows the recipe instead of inventing one.',
  'Every palette traces to a published design system, named in palette_source. When the deck must follow a client brand instead, read the brand palette from its public style guide with browser_visit, then build the plan on those values and keep the same contrast discipline: body text and accent at 4.5:1 or better against every background and surface, and a separate accent for the inverted tonal group.',
].join(' ')

export function themeSummary(theme: PptTheme): ThemeSummary {
  return {
    id: theme.id,
    name: theme.name,
    concept: theme.concept,
    scenes: [...theme.scenes],
    palette_source: theme.palette_source,
    accent: theme.palette.accent,
    accent_inverted: theme.palette.accent_inverted,
    display_font: theme.typography.display.family,
    body_font: theme.typography.body.family,
    signature: theme.decoration[0] ?? '',
  }
}

export function listThemes(themes: readonly PptTheme[] = PPT_THEMES, scene?: string): ThemeCatalog {
  const selected = scene === undefined || scene.trim() === ''
    ? themes
    : themes.filter(theme => theme.scenes.some(entry => entry.toLowerCase().includes(scene.trim().toLowerCase())))
  return {
    themes: selected.map(themeSummary),
    usage: USAGE,
    warnings: selected.length === 0 ? [`no theme matches scene "${scene ?? ''}"; call ppt_themes without a scene to list every theme`] : [],
  }
}

export function findTheme(id: string, themes: readonly PptTheme[] = PPT_THEMES): PptTheme {
  const theme = themes.find(candidate => candidate.id === id)
  if (theme === undefined) {
    throw new PptError('PPT_THEME_UNKNOWN', `unknown theme: ${id}`, {
      details: { available: themes.map(candidate => candidate.id) },
    })
  }
  return theme
}

export interface ThemePagePlan {
  page: number
  type: typeof SLIDE_TYPES[number]
  composition: ArtComposition
  density: ArtDensity
  background_role: ArtBackground
  title_treatment: ArtTitleTreatment
  frame_policy: ArtFramePolicy
}

/**
 * Expand the theme's cycles into one visual plan per page. The caller keeps the
 * content decisions (job, takeaway, visual anchor), so this returns only the
 * structural half of an Art Direction plan.
 */
export function planThemePages(
  theme: PptTheme,
  types: readonly (typeof SLIDE_TYPES[number])[],
): ThemePagePlan[] {
  const covers = new Set(['cover', 'section'])
  let groupedBudget = theme.rhythm_limits.max_grouped_frame_slides
  let previous: ArtComposition | undefined
  return types.map((type, index) => {
    const composition = compositionFor(theme, type, index, previous)
    previous = composition
    const background = theme.background_cycle[index % theme.background_cycle.length]!
    // `grouped` means at least two frame elements on the page, and the theme
    // budgets how many pages the deck may spend that way. Spending the budget
    // here keeps artDirectionFindings clean instead of reporting an overrun.
    const wantsGrouped = composition === 'layered' || composition === 'data-focus'
    const frame_policy: ArtFramePolicy = wantsGrouped && groupedBudget > 0 ? 'grouped' : 'none'
    if (frame_policy === 'grouped') groupedBudget -= 1
    return {
      page: index + 1,
      type,
      composition,
      density: densityFor(composition),
      background_role: background,
      title_treatment: covers.has(type) ? 'statement' as const : 'label' as const,
      frame_policy,
    }
  })
}

/**
 * Which compositions can carry which slide roles. Derived from the outline
 * compatibility table in src/outline.ts, so a theme can never hand the outline
 * a composition the page type cannot express.
 */
const COMPOSITION_TYPES: Record<ArtComposition, readonly (typeof SLIDE_TYPES[number])[]> = {
  hero: ['cover', 'section'],
  'editorial-split': ['agenda', 'content', 'summary'],
  'asymmetric-split': ['content', 'comparison', 'data'],
  process: ['process', 'timeline'],
  layered: ['agenda', 'content', 'summary', 'data'],
  'data-focus': ['data'],
  quote: ['quote'],
  'full-bleed': ['cover', 'section', 'quote', 'ending'],
  closing: ['ending'],
}

/** Used when a theme's cycle carries nothing the page type can express. */
const COMPOSITION_DEFAULTS: Record<(typeof SLIDE_TYPES[number]), ArtComposition> = {
  cover: 'hero',
  agenda: 'layered',
  section: 'hero',
  content: 'editorial-split',
  comparison: 'asymmetric-split',
  timeline: 'process',
  process: 'process',
  data: 'data-focus',
  quote: 'quote',
  summary: 'layered',
  ending: 'closing',
}

/**
 * Walk the theme's own composition cycle from the current page onwards and take
 * the first entry that fits this slide role *and* differs from the previous
 * page's composition, so a deck keeps its theme's character instead of being
 * forced onto one generic composition per role, and does not repeat a shape
 * unless the role leaves the theme no alternative.
 */
function compositionFor(
  theme: PptTheme,
  type: (typeof SLIDE_TYPES[number]),
  index: number,
  previous?: ArtComposition,
): ArtComposition {
  const cycle = theme.composition_cycle
  const candidates: ArtComposition[] = []
  for (let offset = 0; offset < cycle.length; offset += 1) {
    const candidate = cycle[(index + offset) % cycle.length]!
    if (COMPOSITION_TYPES[candidate].includes(type) && !candidates.includes(candidate)) candidates.push(candidate)
  }
  if (candidates.length === 0) return COMPOSITION_DEFAULTS[type]
  return candidates.find(candidate => candidate !== previous) ?? candidates[0]!
}

function densityFor(composition: ArtComposition): ArtDensity {
  if (composition === 'hero' || composition === 'quote' || composition === 'full-bleed' || composition === 'closing') return 'low'
  if (composition === 'layered' || composition === 'data-focus') return 'high'
  return 'medium'
}

/**
 * The theme fields an Art Direction plan inherits verbatim. Everything else in
 * the plan stays the agent's own composition work.
 */
export function themePaletteAndType(theme: PptTheme): Pick<ArtDirection, 'palette' | 'typography'> {
  return {
    palette: {
      background: [...theme.palette.background],
      surface: [...theme.palette.surface],
      accent: theme.palette.accent,
      text: [...theme.palette.text],
    },
    typography: {
      display: { ...theme.typography.display },
      body: { ...theme.typography.body },
      latin: { ...theme.typography.latin },
      code: { ...theme.typography.code },
    },
  }
}

/**
 * Confirm an authored plan still belongs to the theme it claims. Palette drift
 * is the failure this catches: a deck that names a theme and then uses unrelated
 * colours reads as neither.
 */
export function themeFindingsForPlan(theme: PptTheme, plan: ArtDirection): ThemeFinding[] {
  const findings: ThemeFinding[] = []
  const allowed = new Set([
    ...theme.palette.background, ...theme.palette.surface, ...theme.palette.text,
    theme.palette.accent, theme.palette.accent_inverted,
  ])
  const declared = [plan.palette.accent, ...plan.palette.background, ...plan.palette.surface, ...plan.palette.text]
  const foreign = [...new Set(declared.filter(value => !allowed.has(value)))].sort()
  if (foreign.length > 0) {
    findings.push({
      code: 'THEME_PALETTE_DRIFT', severity: 'warning',
      message: `theme ${theme.id} does not define ${foreign.join(', ')}; keep extra colours out or pick a theme that already carries them`,
    })
  }
  if (plan.palette.accent !== theme.palette.accent) {
    findings.push({
      code: 'THEME_ACCENT_REPLACED', severity: 'warning',
      message: `theme ${theme.id} accent is ${theme.palette.accent} but the plan uses ${plan.palette.accent}`,
    })
  }
  const roles = ['display', 'body', 'latin', 'code'] as const
  for (const role of roles) {
    if (plan.typography[role].family !== theme.typography[role].family) {
      findings.push({
        code: 'THEME_TYPOGRAPHY_REPLACED', severity: 'warning',
        message: `theme ${theme.id} ${role} font is ${theme.typography[role].family} but the plan uses ${plan.typography[role].family}`,
      })
    }
  }
  return findings
}

/**
 * Twelve built-in themes. Every palette clears WCAG AA (4.5:1) for body text
 * against every background and surface it defines, and for the accent as well,
 * because the accent carries emphasis text and not only rules. The palettes come
 * from Carbon, Tailwind, Radix, Fluent 2 and Open Color values; the typography is
 * restricted to families this build can actually render (see FONT_REGISTRY).
 */
export const PPT_THEMES: readonly PptTheme[] = [
  {
    id: 'carbon-blueprint',
    name: '碳素蓝图',
    concept: 'Carbon blueprint clarity',
    audience_effect: 'Engineering leadership reads a system diagram, not a slide',
    scenes: ['技术方案', '架构评审', 'RFC 评审'],
    palette_source: 'IBM Carbon gray10/90 + blue70/blue40',
    palette: {
      background: ['#FFFFFF', '#161616'],
      surface: ['#F4F4F4', '#262626'],
      accent: '#0043CE',
      accent_inverted: '#78A9FF',
      text: ['#161616', '#F4F4F4'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Segoe UI', weight: 600 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'editorial-split', 'process', 'data-focus', 'asymmetric-split', 'closing'],
    background_cycle: ['base', 'base', 'inverse', 'base', 'base', 'accent'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Accent rule: an 8px accent bar pinned to the left edge of the content area (position:absolute; left:0; top:0; bottom:0).',
      'Eyebrow label: 2px accent top border, 16px text, letter-spacing 0.12em, text-transform:uppercase.',
      'Cards: surface fill, 1px #E0E0E0 border, 8px radius. There is no shadow in this design system.',
      'Page number bottom right at 14px in #6F6F6F.',
      'Process pages use a 1px connector line with 32px numbered columns.',
    ],
    layout_notes: {
      hero: 'Title at 72px over an accent bar; keep the subtitle to one line.',
      'editorial-split': '7fr/5fr split: claim on the left, evidence card on the right.',
      process: '1px spine across the content area, numbered nodes on a repeat(4,1fr) grid.',
      'data-focus': 'Metric at 88px in the latin face with tabular figures, unit at 24px.',
    },
  },
  {
    id: 'slate-review',
    name: '板岩复盘',
    concept: 'Quarterly ledger',
    audience_effect: 'Every claim lands on a number the room already trusts',
    scenes: ['季度复盘', '经营分析', 'OKR 回顾'],
    palette_source: 'Tailwind slate50/900 + blue700',
    palette: {
      background: ['#F8FAFC', '#0F172A'],
      surface: ['#E2E8F0', '#1E293B'],
      accent: '#1D4ED8',
      accent_inverted: '#60A5FA',
      text: ['#0F172A', '#F1F5F9'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Arial', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'data-focus', 'editorial-split', 'process', 'layered', 'closing'],
    background_cycle: ['base', 'base', 'base', 'inverse', 'base', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Data pages lead with an 88px metric in the latin face using tabular figures above a 1px baseline rule.',
      'Tables: left-aligned header at weight 400, 1px #E2E8F0 row rules, no vertical lines.',
      'Card groups: surface fill with a 1px border, offset by margin rather than shadow.',
    ],
    layout_notes: {
      hero: 'One line of title and one number; nothing else competes.',
      'data-focus': 'Metric block top left, supporting table or chart below with 24px gap.',
      'editorial-split': 'Claim left, the quarter-over-quarter numbers right.',
      layered: 'Three stacked surface cards, each offset 24px down and right of the last.',
    },
  },
  {
    id: 'midnight-raise',
    name: '午夜融资',
    concept: 'Midnight conviction',
    audience_effect: 'The room feels the market window closing and the ask as inevitable',
    scenes: ['融资路演', '战略汇报', '董事会'],
    palette_source: 'IBM Carbon blue10/blue40 + coolGray90',
    palette: {
      background: ['#0B1220', '#FFFFFF'],
      surface: ['#21272A', '#F2F4F8'],
      accent: '#78A9FF',
      accent_inverted: '#0043CE',
      text: ['#EDF5FF', '#0B1220'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Segoe UI', weight: 600 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'quote', 'editorial-split', 'layered', 'data-focus', 'closing'],
    background_cycle: ['inverse', 'inverse', 'accent', 'base', 'inverse', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'A single 1px cool-grey rule (1160x1) in the top-left of dark pages is the only geometry.',
      'Emphasis is accent text plus an 8px vertical bar; never a filled accent block behind body copy.',
      'Secondary text on dark pages uses #C6C6C6, which clears 10.96:1 against the base ink.',
      'The light inversion page switches the accent role to #0043CE for a 7.79:1 white-background read.',
    ],
    layout_notes: {
      hero: 'Dark full-bleed field, 72px title on the left third, 1px rule above it.',
      quote: 'A single line at 44px with the accent bar to its left, nothing else on the page.',
      'data-focus': 'One 88px figure per page; the trend line sits under it as a 1px rule.',
      layered: 'Two surface cards at 30% and 70% width, stacked with a 24px vertical gap.',
    },
  },
  {
    id: 'paper-ink',
    name: '纸墨',
    concept: 'Paper and ink',
    audience_effect: 'The deck reads like a considered essay rather than a sales pitch',
    scenes: ['学术报告', '白皮书', '深度长文'],
    palette_source: 'dsh-ppt official fixture field #F7F5F0/#111318 + IBM Carbon magenta70',
    palette: {
      background: ['#F7F5F0', '#111318'],
      surface: ['#FFFFFF', '#1C1F26'],
      accent: '#9F1853',
      accent_inverted: '#FF7EB6',
      text: ['#111318', '#FFFFFF'],
    },
    typography: {
      display: { family: 'SimSun', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Times New Roman', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'editorial-split', 'quote', 'process', 'full-bleed', 'closing'],
    background_cycle: ['base', 'base', 'base', 'base', 'inverse', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'A 6px crimson vertical bar sits to the left of the title block.',
      'A 1px warm-grey divider (height:1px; background:#E5E0DF; margin:24px 0) separates blocks.',
      'Footer citations set small numbers in the latin serif face.',
      'Tint is spent only on the quote page rule; page fields stay the paper colour.',
    ],
    layout_notes: {
      hero: 'Serif title at 72px with the crimson bar left of it; warm paper field behind.',
      'editorial-split': '8fr/4fr: the argument column left, a single pull quote right.',
      quote: 'Centred serif paragraph at 32px between two 1px dividers.',
      'full-bleed': 'One inverted paper field for the deck midpoint; title only.',
    },
  },
  {
    id: 'editorial-serif',
    name: '学刊',
    concept: 'Scholarly editorial',
    audience_effect: 'Reviewers read the argument as a paper with typed evidence',
    scenes: ['学术答辩', '研究报告', '政策解读'],
    palette_source: 'IBM Carbon warmGray10/90 + orange70',
    palette: {
      background: ['#F7F3F2', '#171414'],
      surface: ['#E5E0DF', '#272525'],
      accent: '#8A3800',
      accent_inverted: '#FF832B',
      text: ['#171414', '#F7F3F2'],
    },
    typography: {
      display: { family: 'SimSun', weight: 700 },
      body: { family: 'SimSun', weight: 400 },
      latin: { family: 'Times New Roman', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'editorial-split', 'data-focus', 'quote', 'process', 'closing'],
    background_cycle: ['base', 'base', 'base', 'base', 'inverse', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Serif throughout: SimSun for headings and body, Times New Roman for latin and figures.',
      'Tables take a 3px current-colour header underline and 1px row rules.',
      'Section numbers are set as a 240px serif numeral watermark at z-index:0 in a light grey.',
    ],
    layout_notes: {
      hero: 'Serif display at 72px, section numeral watermark behind it.',
      'editorial-split': 'Equal halves, serif body at 24px with 1.6 line height.',
      'data-focus': 'A table rather than a chart leads this theme: header rule plus row rules.',
      process: 'Numbered steps in serif numerals, 1px rule between each.',
    },
  },
  {
    id: 'telemetry-teal',
    name: '遥测青',
    concept: 'Instrument panel',
    audience_effect: 'Operators see live signals and the one threshold that moved',
    scenes: ['运维复盘', '数据看板汇报', '性能评审'],
    palette_source: 'IBM Carbon teal10/60/70/100 + gray100',
    palette: {
      background: ['#FFFFFF', '#081A1C'],
      surface: ['#D9FBFB', '#004144'],
      accent: '#005D5D',
      accent_inverted: '#3DDBD9',
      text: ['#081A1C', '#D9FBFB'],
    },
    typography: {
      display: { family: 'DengXian', weight: 700 },
      body: { family: 'DengXian', weight: 400 },
      latin: { family: 'Segoe UI', weight: 600 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'data-focus', 'process', 'editorial-split', 'layered', 'closing'],
    background_cycle: ['base', 'base', 'base', 'base', 'inverse', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Metric cards: teal-tint fill, 1px border, 24px padding.',
      'Threshold marker: a 2px solid accent rule across the content area.',
      'Step rail: an even grid of 32px numbers joined by a 1px line.',
    ],
    layout_notes: {
      hero: 'Title left, the single threshold rule beneath it.',
      'data-focus': 'Four metric cards on a repeat(4,1fr) grid, one highlighted with accent text.',
      process: 'Numbered rail with the accent rule marking the current stage.',
      layered: 'Two card rows on the dark page, 1px borders visible against the ink.',
    },
  },
  {
    id: 'graphite-minimal',
    name: '石墨极简',
    concept: 'Graphite minimal',
    audience_effect: 'Nothing decorative competes with the specification',
    scenes: ['产品规格', '内部对齐', '工程文档'],
    palette_source: 'Radix gray1/gray3/gray12 + indigo11',
    palette: {
      background: ['#FCFCFC', '#202020'],
      surface: ['#F0F0F0', '#393939'],
      accent: '#3A5BC7',
      accent_inverted: '#8DA4F0',
      text: ['#202020', '#FCFCFC'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Arial', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'editorial-split', 'data-focus', 'layered', 'process', 'closing'],
    background_cycle: ['base', 'base', 'base', 'base', 'base', 'inverse'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Zero ornament: the whole deck runs on 1px borders and type scale.',
      'Grid guides: 1px #F0F0F0 column rules, used only on hero and layered pages.',
      'Page numbers are set as a 240px grey numeral behind the top-right corner at z-index:0.',
    ],
    layout_notes: {
      hero: 'One 72px line, a 1px rule, and a grey numeral watermark.',
      'editorial-split': 'Strict 1fr/1fr halves divided by a 1px vertical rule.',
      'data-focus': 'A single table or metric; no chart chrome at all.',
      layered: 'Three cards on the same baseline, separated by 1px borders only.',
    },
  },
  {
    id: 'indigo-launch',
    name: '靛蓝发布',
    concept: 'Launch night',
    audience_effect: 'The feature set arrives as a staged reveal with a single colour of momentum',
    scenes: ['产品发布', '新版本宣讲', '大会 Keynote'],
    palette_source: 'Tailwind indigo950/900/800/300/50',
    palette: {
      background: ['#1E1B4B', '#EEF2FF'],
      surface: ['#312E81', '#E0E7FF'],
      accent: '#A5B4FC',
      accent_inverted: '#3730A3',
      text: ['#EEF2FF', '#1E1B4B'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'DengXian', weight: 400 },
      latin: { family: 'Segoe UI', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'full-bleed', 'quote', 'layered', 'editorial-split', 'closing'],
    background_cycle: ['inverse', 'inverse', 'accent', 'base', 'inverse', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Section pages are a full-bleed #312E81 field with a centred 44px title.',
      'Launch pages open with a 96px version number in the latin face.',
      'On light inversion pages the accent role moves to #3730A3 for an 8.88:1 read.',
    ],
    layout_notes: {
      hero: 'Deep indigo field, 72px title left, version number as the only ornament.',
      'full-bleed': 'Full-viewport indigo with the section name centred at 44px.',
      quote: 'One line of the roadmap promise, centred, at 44px.',
      layered: 'Feature cards in E0E7FF on the indigo field, 1px borders.',
    },
  },
  {
    id: 'amber-academy',
    name: '琥珀课堂',
    concept: 'Workshop chalk',
    audience_effect: 'Learners follow one numbered step at a time without losing the thread',
    scenes: ['培训', '工作坊', '新人上手'],
    palette_source: 'Tailwind amber50/100/700/950 + IBM Carbon yellow20',
    palette: {
      background: ['#FFFBEB', '#451A03'],
      surface: ['#FEF3C7', '#78350F'],
      accent: '#92400E',
      accent_inverted: '#FBBF24',
      text: ['#451A03', '#FFFBEB'],
    },
    typography: {
      display: { family: 'DengXian', weight: 700 },
      body: { family: 'DengXian', weight: 400 },
      latin: { family: 'Arial', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'process', 'process', 'editorial-split', 'data-focus', 'closing'],
    background_cycle: ['base', 'base', 'base', 'base', 'base', 'inverse'],
    rhythm_limits: { max_grouped_frame_slides: 3, max_same_composition_run: 2 },
    decoration: [
      'Step cards: amber-tint fill with a 1px #FDDC69 border.',
      'Number chips: 32px squares, 8px radius, accent fill, light numeral.',
      'This theme deliberately allows two consecutive process pages; the rhythm limit says so.',
    ],
    layout_notes: {
      hero: 'Warm field, 72px title, the workshop promise in one line.',
      process: 'Two consecutive step pages are expected; keep three to four steps each.',
      'editorial-split': 'Instruction left, worked example right.',
      'data-focus': 'A small table of before/after values rather than a chart.',
    },
  },
  {
    id: 'crimson-brief',
    name: '绯红简报',
    concept: 'Crimson brief',
    audience_effect: 'One proposition, one number, one deadline, stated with heat',
    scenes: ['营销提案', '立项申请', '一页纸决策'],
    palette_source: 'Tailwind rose50/700/900/950',
    palette: {
      background: ['#FFFFFF', '#4C0519'],
      surface: ['#FFF1F2', '#881337'],
      accent: '#BE123C',
      accent_inverted: '#FDA4AF',
      text: ['#4C0519', '#FFF1F2'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Arial', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'quote', 'data-focus', 'editorial-split', 'asymmetric-split', 'closing'],
    background_cycle: ['base', 'base', 'inverse', 'base', 'base', 'inverse'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'The single-proposition page is a quote composition with a 6px crimson bar on the left.',
      'The one decisive number is set at 88px in the accent colour with tabular figures.',
      'The closing page is a full #4C0519 field with #FFE4E6 text.',
    ],
    layout_notes: {
      hero: 'White field, 72px title, the deadline as a single accent line beneath.',
      quote: 'The core proposition at 44px beside the crimson bar.',
      'data-focus': 'Exactly one figure, 88px, nothing else on the page.',
      closing: 'Deep crimson field with the ask as the only text.',
    },
  },
  {
    id: 'fluent-azure',
    name: '流蓝企业',
    concept: 'Fluent azure',
    audience_effect: 'An enterprise audience recognizes a governed, reviewable proposal',
    scenes: ['企业汇报', '客户方案', '招标应答'],
    palette_source: 'Microsoft Fluent 2 brandWeb 10/30/80/160',
    palette: {
      background: ['#FFFFFF', '#0A2E4A'],
      surface: ['#EBF3FC', '#0F548C'],
      accent: '#0F6CBD',
      accent_inverted: '#A9D3F7',
      text: ['#061724', '#EBF3FC'],
    },
    typography: {
      display: { family: 'Microsoft YaHei', weight: 700 },
      body: { family: 'Microsoft YaHei', weight: 400 },
      latin: { family: 'Segoe UI', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'editorial-split', 'layered', 'process', 'data-focus', 'closing'],
    background_cycle: ['base', 'base', 'base', 'inverse', 'base', 'accent'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Fluent 2 feel: 600-weight headings, radii limited to 4 and 8px, spacing on a 4px base.',
      'Section band: an 8px accent strip along the bottom edge of section pages.',
      'Cards: surface fill, 1px border, radius 8, padding 24.',
    ],
    layout_notes: {
      hero: 'Title left with the section band along the bottom edge.',
      'editorial-split': 'Proposal claim left, scope card right.',
      layered: 'Three scope cards on a repeat(3,1fr) grid with 24px gaps.',
      'data-focus': 'Milestone table with a 1px header rule.',
    },
  },
  {
    id: 'moss-annual',
    name: '苔绿年报',
    concept: 'Annual moss',
    audience_effect: 'Stakeholders read steady, compounding movement instead of a spike',
    scenes: ['年度报告', 'ESG 披露', '长期规划'],
    palette_source: 'Tailwind green50/100/700/900',
    palette: {
      background: ['#F0FDF4', '#14532D'],
      surface: ['#DCFCE7', '#166534'],
      accent: '#15803D',
      accent_inverted: '#86EFAC',
      text: ['#14532D', '#F0FDF4'],
    },
    typography: {
      display: { family: 'Noto Sans SC', weight: 700 },
      body: { family: 'Noto Sans SC', weight: 400 },
      latin: { family: 'Arial', weight: 700 },
      code: { family: 'DejaVu Sans', weight: 400 },
    },
    composition_cycle: ['hero', 'data-focus', 'editorial-split', 'layered', 'process', 'closing'],
    background_cycle: ['base', 'base', 'base', 'base', 'inverse', 'base'],
    rhythm_limits: { max_grouped_frame_slides: 2, max_same_composition_run: 1 },
    decoration: [
      'Long-horizon charts sit on a green-tint band with 1px grid rules.',
      'Data cards: green-tint fill, 1px border, figures in the accent colour.',
      'Section pages are a full #14532D field with #F0FDF4 text.',
    ],
    layout_notes: {
      hero: 'Soft green field, 72px title, the reporting period beneath.',
      'data-focus': 'Multi-year comparison; prefer a table of years over a single spike chart.',
      'editorial-split': 'Commitment left, measured result right.',
      layered: 'Three period cards on the same baseline.',
    },
  },
]
