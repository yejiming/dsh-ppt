import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { inject } from '../../src/index.ts'

describe('package contract', () => {
  it('publishes the five runtime entries and three-platform support', async () => {
    const manifest = JSON.parse(await readFile(new URL('../../package.json', import.meta.url), 'utf8')) as {
      name: string
      main: string
      os: string[]
      exports: Record<string, string | { types: string; default: string }>
      files: string[]
    }
    expect(manifest.name).toBe('@yejiming/dsh-ppt')
    expect(manifest.os).toEqual(['darwin', 'linux', 'win32'])
    expect(Object.keys(manifest.exports)).toEqual(expect.arrayContaining(['.', './tools', './headless', './runtime', './schemas']))
    expect(manifest.files).toContain('cordis.patch.yml')
    expect(manifest.files).not.toContain('preset/**')
    expect(manifest.main).toBe('lib/index.mjs')
    expect(inject).toEqual(['sandbox', 'subprocess'])
    for (const value of Object.values(manifest.exports)) {
      if (typeof value !== 'object') continue
      expect(value.default).toMatch(/^\.\/lib\/.+\.mjs$/u)
      expect(value.types).toMatch(/^\.\/lib\/types\/.+\.d\.ts$/u)
    }
  })

  it('uses the DSH host sharp singleton instead of publishing a private native copy', async () => {
    const manifest = JSON.parse(await readFile(new URL('../../package.json', import.meta.url), 'utf8')) as {
      dependencies: Record<string, string>
      devDependencies: Record<string, string>
      peerDependencies: Record<string, string>
      peerDependenciesMeta: Record<string, { optional?: boolean }>
    }
    expect(manifest.dependencies).not.toHaveProperty('sharp')
    expect(manifest.devDependencies.sharp).toBe('0.35.4')
    expect(manifest.peerDependencies.sharp).toBe('^0.35.3')
    expect(manifest.peerDependenciesMeta.sharp).toEqual({ optional: true })
  })

  it('ships a valid host entry and a headless-only preset bridge', async () => {
    const plugin = JSON.parse(await readFile(new URL('../../dsh-plugin.json', import.meta.url), 'utf8')) as {
      version: string
      facets: { host: { entry: string } }
    }
    const manifest = JSON.parse(await readFile(new URL('../../package.json', import.meta.url), 'utf8')) as { version: string }
    expect(plugin.version).toBe(manifest.version)
    const patch = await readFile(new URL('../../cordis.patch.yml', import.meta.url), 'utf8')
    expect(plugin.facets.host.entry).toBe('lib/index.mjs')
    expect(patch).toContain("name: '@deepseek-ai/dsh-agent-preset'")
    expect(patch).toContain("name: '@deepseek-ai/dsh-agent-preset-registry'")
    expect(patch).toContain("name: '@yejiming/dsh-ppt/headless'")
    expect(patch.match(/inject: \[headlessStartup\]/gu)).toHaveLength(2)
    expect(patch.match(/process\.argv\.includes\('--profile=headless'\)/gu)).toHaveLength(2)
    expect(patch).toMatch(/id: headless-runner[\s\S]*?disabled: true/u)
    expect(patch).toContain('task: !!js ctx.headlessStartup.task')
  })
})
