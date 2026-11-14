import { describe, expect, it } from 'vitest'
import { resolveRegisteredFont, type DiscoveredFont } from '../../src/fonts.ts'
import { validateArtDirection } from '../../src/art-direction.ts'
import { validatePptOutline } from '../../src/outline.ts'
import { MINIMAL_ART_DIRECTION } from '../fixtures/minimal-art-direction.ts'
import { MINIMAL_OUTLINE } from '../fixtures/minimal-outline.ts'

function discovered(name: string, characters: string): DiscoveredFont {
  const codePoints = new Set([...characters].map(character => character.codePointAt(0)!))
  return {
    name, file: `/fonts/${name}.ttf`, sha256: 'a'.repeat(64), familyName: name, postscriptName: null,
    weight: 'Regular', glyphCount: codePoints.size,
    supportsLatin: [...'AaZz09'].every(character => codePoints.has(character.codePointAt(0)!)),
    supportsCjk: [...'中文'].every(character => codePoints.has(character.codePointAt(0)!)), codePoints,
  }
}

describe('degrading when the registry is newer than the running build', () => {
  it('substitutes a registry-order fallback for an unregistered family instead of throwing', () => {
    const text = '中文'
    const result = resolveRegisteredFont('思源黑体 CN Heavy', text, [discovered('Microsoft YaHei', text)], 'win32')
    expect(result).toMatchObject({ requested: '思源黑体 CN Heavy', fallback: true, resolved: { name: 'Microsoft YaHei' } })
    expect(result.warning).toContain('not registered in this build')
  })

  it('picks the same fallback every time for the same inputs', () => {
    const text = 'AaZz09'
    const pool = [discovered('Arial', text), discovered('Microsoft YaHei', text)]
    const first = resolveRegisteredFont('Ghost Family', text, pool, 'win32')
    const second = resolveRegisteredFont('Ghost Family', text, pool, 'win32')
    expect(first.resolved.name).toBe(second.resolved.name)
  })

  it('still fails loudly when no installed font covers the requested text', () => {
    expect(() => resolveRegisteredFont('Ghost Family', '中文', [], 'win32'))
      .toThrow(/no installed approved font covers the requested text/u)
  })

  it('accepts an art direction authored against a newer registry', () => {
    const plan = {
      ...MINIMAL_ART_DIRECTION,
      typography: {
        display: { family: '思源黑体 CN Heavy', weight: 400 },
        body: { family: '思源黑体 CN Regular', weight: 400 },
        latin: { family: '思源黑体 CN Regular', weight: 400 },
        code: { family: '思源黑体 CN Regular', weight: 400 },
      },
    }
    expect(validateArtDirection(plan).typography.display.family).toBe('思源黑体 CN Heavy')
  })

  it('accepts an outline authored against a newer registry', () => {
    const outline = [{
      ...MINIMAL_OUTLINE[0],
      style: {
        ...MINIMAL_OUTLINE[0].style,
        title_font: '思源黑体 CN Heavy',
        body_font: '思源黑体 CN Regular',
      },
    }]
    expect(validatePptOutline(outline)[0]!.style.title_font).toBe('思源黑体 CN Heavy')
  })

  it('keeps rejecting a font name that carries markup', () => {
    const outline = [{
      ...MINIMAL_OUTLINE[0],
      style: { ...MINIMAL_OUTLINE[0].style, title_font: '<b>Liter</b>' },
    }]
    expect(() => validatePptOutline(outline)).toThrow(/PPT outline validation failed/u)
  })
})
