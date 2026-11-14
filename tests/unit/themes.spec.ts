import { describe, expect, it } from 'vitest'
import { ART_COMPOSITIONS, validateArtDirection } from '../../src/art-direction.ts'
import { validatePptOutline } from '../../src/outline.ts'
import { resolveTheme, themeConformanceFindings } from '../../src/outline.ts'
import { SLIDE_TYPES } from '../../src/outline.ts'
import {
  contrastRatio, findTheme, listThemes, planThemePages, PPT_THEMES, relativeLuminance,
  themeFindings, themeFindingsForPlan, themePaletteAndType, validateTheme,
} from '../../src/themes.ts'
import { MINIMAL_ART_DIRECTION } from '../fixtures/minimal-art-direction.ts'

/** Which compositions each slide role can actually be rendered as. */
const COMPOSITION_TYPES: Record<string, readonly string[]> = {
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

describe('theme library', () => {
  it('ships a theme for every common briefing scene', () => {
    expect(PPT_THEMES.length).toBeGreaterThanOrEqual(8)
    const ids = PPT_THEMES.map(theme => theme.id)
    expect(new Set(ids).size).toBe(ids.length)
    const scenes = PPT_THEMES.flatMap(theme => theme.scenes)
    for (const required of ['技术方案', '季度复盘', '融资路演', '学术报告', '产品发布', '培训']) {
      expect(scenes).toContain(required)
    }
  })

  it('validates every built-in theme against the real registry and schema', () => {
    for (const theme of PPT_THEMES) {
      expect(validateTheme(theme)).toBe(theme)
      expect(theme.composition_cycle.every(entry => (ART_COMPOSITIONS as readonly string[]).includes(entry))).toBe(true)
    }
  })

  it('holds every built-in palette to WCAG AA', () => {
    const findings = PPT_THEMES.flatMap(theme => themeFindings(theme).filter(finding => finding.severity === 'error'))
    expect(findings).toEqual([])
  })

  it('rejects a theme whose font is outside the approved registry', () => {
    const [first] = PPT_THEMES
    expect(() => validateTheme({ ...first!, typography: { ...first!.typography, display: { family: 'Inter', weight: 700 } } }))
      .toThrow(expect.objectContaining({ code: 'PPT_THEME_INVALID' }))
  })

  it('rejects a theme whose palette is not six-digit hex', () => {
    const [first] = PPT_THEMES
    expect(() => validateTheme({ ...first!, palette: { ...first!.palette, accent: '#fff' } }))
      .toThrow(expect.objectContaining({ code: 'PPT_THEME_INVALID' }))
  })

  it('computes WCAG contrast against the known reference values', () => {
    expect(contrastRatio('#000000', '#FFFFFF')).toBeCloseTo(21, 5)
    expect(contrastRatio('#FFFFFF', '#FFFFFF')).toBeCloseTo(1, 5)
    expect(relativeLuminance('#FFFFFF')).toBeCloseTo(1, 5)
  })

  it('plans only compositions the page role can render', () => {
    for (const theme of PPT_THEMES) {
      const plan = planThemePages(theme, [...SLIDE_TYPES])
      expect(plan.map(entry => entry.page)).toEqual(plan.map((_, index) => index + 1))
      for (const entry of plan) {
        expect(COMPOSITION_TYPES[entry.composition]).toContain(entry.type)
        expect(theme.background_cycle).toContain(entry.background_role)
      }
    }
  })

  it('avoids repeating the previous composition when the theme offers an alternative', () => {
    const theme = findTheme('carbon-blueprint', PPT_THEMES)
    const plan = planThemePages(theme, ['cover', 'content', 'content', 'content', 'ending'])
    const body = plan.slice(1, 4).map(entry => entry.composition)
    expect(body).toHaveLength(3)
    for (let index = 1; index < body.length; index += 1) expect(body[index]).not.toBe(body[index - 1])
  })

  it('keeps grouped frame pages inside the theme budget', () => {
    for (const theme of PPT_THEMES) {
      const plan = planThemePages(theme, [...SLIDE_TYPES])
      const grouped = plan.filter(entry => entry.frame_policy === 'grouped')
      expect(grouped.length).toBeLessThanOrEqual(theme.rhythm_limits.max_grouped_frame_slides)
    }
  })

  it('does not stack the same composition past the theme run limit when alternatives exist', () => {
    // Two slide roles have exactly one composition that can express them
    // (process and timeline both render as `process`), so only decks whose roles
    // leave the theme a choice are checked here.
    const plan = planThemePages(findTheme('slate-review', PPT_THEMES), ['cover', 'agenda', 'content', 'content', 'summary', 'ending'])
    let run = 1
    for (let index = 1; index < plan.length; index += 1) {
      run = plan[index]!.composition === plan[index - 1]!.composition ? run + 1 : 1
      expect(run).toBeLessThanOrEqual(2)
    }
  })

  it('filters the catalogue by scene and reports an empty result', () => {
    expect(listThemes(PPT_THEMES, '技术方案').themes.map(theme => theme.id)).toContain('carbon-blueprint')
    expect(listThemes(PPT_THEMES, '不存在的场景').warnings).toHaveLength(1)
  })

  it('reports an unknown theme with the available ids', () => {
    expect(() => findTheme('not-a-theme', PPT_THEMES)).toThrow(expect.objectContaining({ code: 'PPT_THEME_UNKNOWN' }))
  })

  it('carries the theme palette and typography into a plan shape', () => {
    const theme = findTheme('carbon-blueprint', PPT_THEMES)
    expect(themePaletteAndType(theme).palette.accent).toBe(theme.palette.accent)
    expect(themePaletteAndType(theme).typography.display).toEqual(theme.typography.display)
  })

  it('flags palette and typography drift away from the declared theme', () => {
    const theme = findTheme('carbon-blueprint', PPT_THEMES)
    const base = validateArtDirection(MINIMAL_ART_DIRECTION, 1)
    const drifted = {
      ...base,
      palette: { background: [...base.palette.background], surface: [...base.palette.surface], accent: '#FF00FF', text: [...base.palette.text] },
    }
    const codes = themeFindingsForPlan(theme, drifted).map(finding => finding.code)
    expect(codes).toContain('THEME_PALETTE_DRIFT')
    expect(codes).toContain('THEME_ACCENT_REPLACED')
    expect(codes).toContain('THEME_TYPOGRAPHY_REPLACED')
  })

  it('accepts a plan that stays inside its theme', () => {
    const theme = findTheme('carbon-blueprint', PPT_THEMES)
    const base = validateArtDirection(MINIMAL_ART_DIRECTION, 1)
    const { palette, typography } = themePaletteAndType(theme)
    expect(themeFindingsForPlan(theme, { ...base, palette, typography })).toEqual([])
  })

  it('reports conformance problems between an outline and the theme it names', () => {
    const theme = findTheme('carbon-blueprint', PPT_THEMES)
    const outline = validatePptOutline([{
      page: 1, type: 'cover', title: 'A', content: [],
      style: {
        layout: 'cover', background: 'light', accent: '#FF00FF',
        title_font: 'Arial', body_font: 'Arial', visual_direction: 'x',
      },
    }])

    // No plan at all: the theme cannot be enforced.
    expect(themeConformanceFindings(theme, outline, undefined)).toEqual(expect.arrayContaining([
      expect.stringContaining('THEME_PLAN_MISSING'),
    ]))

    const base = validateArtDirection(MINIMAL_ART_DIRECTION, 1)
    const drifted = {
      ...base,
      palette: { background: [...base.palette.background], surface: [...base.palette.surface], accent: '#FF00FF', text: [...base.palette.text] },
    }
    const findings = themeConformanceFindings(theme, outline, drifted)
    expect(findings).toEqual(expect.arrayContaining([
      expect.stringContaining('THEME_ACCENT_DRIFT'),
      expect.stringContaining('THEME_FONT_DRIFT'),
      expect.stringContaining('THEME_ACCENT_REPLACED'),
    ]))
  })

  it('resolves a theme id and rejects an unknown one', () => {
    expect(resolveTheme('paper-ink').name).toBe('纸墨')
    expect(() => resolveTheme('nope')).toThrow(expect.objectContaining({ code: 'PPT_THEME_UNKNOWN' }))
  })
})
