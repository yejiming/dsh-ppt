import { describe, expect, it } from 'vitest'
import { validatePptOutline } from '../../src/outline.ts'

interface CapturedError {
  code?: string
  message: string
  details?: { issues?: Array<{ path: string; message: string }> }
}

function capture(fn: () => unknown): CapturedError {
  try {
    fn()
  } catch (error) {
    return error as CapturedError
  }
  throw new Error('expected validatePptOutline to throw')
}

/**
 * Issue #1 called `ppt_outline` with a shorthand slide (`content` as a string and
 * `style` as a style name) and the surface showed only the constant
 * `PPT outline validation failed`, because the zod issues live in `details`,
 * which the tool-result view does not render.
 */
describe('PPT outline diagnostics', () => {
  it('names the offending JSON paths in the message for a shorthand issue #1 outline', () => {
    const shorthand = [{ page: 1, title: '标题', content: '内容', style: '简单', type: 'cover' }]
    const error = capture(() => validatePptOutline(shorthand))
    expect(error.code).toBe('PPT_OUTLINE_INVALID')
    expect(error.message).toContain('0.content')
    expect(error.message).toContain('0.style')
    expect(error.message).not.toBe('PPT outline validation failed')
    expect(error.details?.issues?.length).toBeGreaterThan(0)
    expect(error.details?.issues?.every(issue => typeof issue.path === 'string')).toBe(true)
  })

  it('keeps structured issues alongside the rendered line', () => {
    const pages = [{
      page: 2, type: 'content', title: 'Title', content: [{ kind: 'point', text: 'A point' }],
      style: {
        layout: 'title-content', background: 'light', accent: '#4F46E5',
        title_font: 'Liter', body_font: 'Liter', visual_direction: 'Clear hierarchy',
      },
    }]
    const error = capture(() => validatePptOutline(pages))
    expect(error.message).toContain('page must be 1')
    expect(error.message).toContain('0.page')
    expect(error.details?.issues).toEqual(expect.arrayContaining([expect.objectContaining({ path: '0.page' })]))
  })

  it('truncates the rendered line and counts the remaining issues', () => {
    const slides = Array.from({ length: 12 }, (_, index) => ({
      page: index + 1, type: 'content', title: 'Title', content: 'not-an-array', style: 'not-an-object',
    }))
    const error = capture(() => validatePptOutline(slides))
    expect(error.message).toMatch(/; \+\d+ more$/u)
    expect(error.message.length).toBeLessThan(2_000)
    expect(error.details?.issues?.length).toBeGreaterThan(8)
  })
})
