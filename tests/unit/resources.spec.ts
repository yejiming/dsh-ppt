import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { allocateArtifactDirectory, SourceManifest } from '../../src/artifacts.ts'
import { resolveRegisteredFont, type DiscoveredFont } from '../../src/fonts.ts'
import { SessionResourceRegistry } from '../../src/session-resources.ts'
import { createTestWorkspace } from '../helpers/workspace.ts'

describe('session resources and artifacts', () => {
  it('isolates owner resources and cleans only uncommitted temporary paths', async () => {
    const workspace = await createTestWorkspace()
    const registry = new SessionResourceRegistry()
    try {
      const ownerA = { agentId: 'a', sessionId: 'one' }
      const ownerB = { agentId: 'b', sessionId: 'two' }
      registry.open(ownerA, workspace.root)
      registry.open(ownerB, workspace.root)
      const temporary = join(workspace.root, 'temporary-a')
      const delivered = join(workspace.root, 'delivered.txt')
      await mkdir(temporary)
      await writeFile(delivered, 'keep')
      let disposed = false
      registry.trackTemporaryPath(ownerA, temporary)
      registry.track(ownerA, { label: 'browser', dispose() { disposed = true } })
      await registry.release(ownerA)
      expect(disposed).toBe(true)
      await expect(readFile(delivered, 'utf8')).resolves.toBe('keep')
      expect(registry.state(ownerB)?.workspace).toBe(workspace.root)
    } finally {
      await registry.dispose()
      await workspace.cleanup()
    }
  })

  it('allocates deterministic non-overwriting directories and appends hashed source records', async () => {
    const workspace = await createTestWorkspace()
    try {
      const first = await allocateArtifactDirectory(workspace.root, 'Quarterly Review')
      const second = await allocateArtifactDirectory(workspace.root, 'Quarterly Review')
      expect(first.root).not.toBe(second.root)
      expect(second.root).toMatch(/quarterly-review-2$/)
      const image = join(first.images, 'chart.png')
      await writeFile(image, 'fixture')
      const manifest = new SourceManifest(workspace.root, first.sourceManifest)
      const item = await manifest.append({
        original_url: 'https://images.example/chart.png',
        source_page: 'https://example/source',
        fetched_at: '2026-08-24T00:00:00.000Z',
        author: null,
        license: 'unknown',
        license_url: null,
        localFile: image,
      })
      expect(item.local_path).toBe('images/chart.png')
      expect(item.sha256).toMatch(/^[a-f0-9]{64}$/)
      expect(JSON.parse(await readFile(first.sourceManifest, 'utf8')).assets).toHaveLength(1)
    } finally {
      await workspace.cleanup()
    }
  })

  it('uses only deterministic registered font fallbacks with glyph coverage', () => {
    const latin = new Set([...'AaZz09Hello'].map(character => character.codePointAt(0)!))
    const fallback: DiscoveredFont = {
      name: 'HedvigLettersSans', file: '/fonts/hedvig.ttf', sha256: 'a'.repeat(64), familyName: 'HedvigLettersSans',
      postscriptName: null, weight: 'Regular', glyphCount: latin.size, supportsLatin: true, supportsCjk: false, codePoints: latin,
    }
    expect(resolveRegisteredFont('Liter', 'Hello', [fallback])).toMatchObject({ fallback: true, resolved: { name: 'HedvigLettersSans' } })
    // An unknown family degrades to a deterministic registry-order fallback rather
    // than aborting: a deck authored against a newer registry must still render.
    expect(resolveRegisteredFont('Unregistered', 'Hello', [fallback]))
      .toMatchObject({ requested: 'Unregistered', fallback: true, resolved: { name: 'HedvigLettersSans' } })
    expect(() => resolveRegisteredFont('MiSans', '中文', [fallback])).toThrow(expect.objectContaining({ code: 'PPT_DEPENDENCY_MISSING' }))
  })
})
