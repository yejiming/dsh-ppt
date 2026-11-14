import { describe, expect, it } from 'vitest'
import { occupancySilhouette, type SlideLeafBox } from '../../src/browser.ts'

const SLIDE = { width: 1280, height: 720 }
const EMPTY = Array.from({ length: 16 }, () => 0)
/** Page background leaf: `position:absolute; inset:0` inside a 1280x720 slide. */
const BACKGROUND: SlideLeafBox = { left: 0, top: 0, right: 1280, bottom: 720 }

function litRegions(occupancy: number[]): number {
  return occupancy.filter(value => value === 1).length
}

function intersectionOverUnion(left: number[], right: number[]): number {
  const intersection = left.filter((value, index) => value === 1 && right[index] === 1).length
  const union = left.filter((value, index) => value === 1 || right[index] === 1).length
  return union === 0 ? 0 : intersection / union
}

describe('slide occupancy silhouette', () => {
  it('ignores a full-bleed background layer', () => {
    expect(occupancySilhouette([BACKGROUND], SLIDE)).toEqual(EMPTY)
  })

  it('ignores a full-bleed layer with sub-pixel inset or a slide border', () => {
    const inset: SlideLeafBox = { left: 0.5, top: 0.5, right: 1279.5, bottom: 719.5 }
    const bordered: SlideLeafBox = { left: 1, top: 1, right: 1279, bottom: 719 }
    expect(occupancySilhouette([inset], SLIDE)).toEqual(EMPTY)
    expect(occupancySilhouette([bordered], SLIDE)).toEqual(EMPTY)
  })

  it('keeps counting a full-width header band', () => {
    const band: SlideLeafBox = { left: 0, top: 0, right: 1280, bottom: 90 }
    const occupancy = occupancySilhouette([band], SLIDE)
    expect(litRegions(occupancy)).toBe(4)
    expect(occupancy.slice(0, 4)).toEqual([1, 1, 1, 1])
    expect(occupancy.slice(4)).toEqual(Array.from({ length: 12 }, () => 0))
  })

  it('keeps counting a large block that does not cover the whole page', () => {
    const block: SlideLeafBox = { left: 40, top: 40, right: 1240, bottom: 680 }
    expect(occupancySilhouette([block], SLIDE)).toEqual(Array.from({ length: 16 }, () => 1))
    const nearFullPage: SlideLeafBox = { left: 8, top: 8, right: 1272, bottom: 712 }
    expect(litRegions(occupancySilhouette([nearFullPage], SLIDE))).toBe(16)
  })

  it('keeps the content silhouette of pages that share one background', () => {
    const left: SlideLeafBox = { left: 80, top: 160, right: 600, bottom: 560 }
    const right: SlideLeafBox = { left: 680, top: 160, right: 1200, bottom: 560 }
    const page1 = occupancySilhouette([BACKGROUND, left], SLIDE)
    const page2 = occupancySilhouette([BACKGROUND, right], SLIDE)
    const page3 = occupancySilhouette([BACKGROUND, left], SLIDE)
    expect(page1).toEqual([1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0])
    // Distinct layouts must not read as a repeated silhouette.
    expect(intersectionOverUnion(page1, page2)).toBeLessThanOrEqual(0.88)
    // Identical layouts must still be reported as repeated through the same 0.88 rule.
    expect(intersectionOverUnion(page1, page3)).toBeGreaterThan(0.88)
  })
})
