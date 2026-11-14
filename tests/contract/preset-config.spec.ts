import { readFileSync } from 'node:fs'
import { Config as PresentationConfig } from '@deepseek-ai/dsh-agent-tool-presentation'
import AgentPreset from '@deepseek-ai/dsh-agent-preset'
import { AgentPresetRegistry } from '@deepseek-ai/dsh-agent-preset-registry'
import { Config as PersonaConfig } from '@deepseek-ai/dsh-persona'
import { Config as ToolAskUserConfig } from '@deepseek-ai/dsh-tool-ask-user'
import { Config as ToolBashConfig } from '@deepseek-ai/dsh-tool-bash'
import { Config as ToolFsConfig } from '@deepseek-ai/dsh-tool-fs'
import { Config as ToolTodoConfig } from '@deepseek-ai/dsh-tool-todo'
import yaml from 'js-yaml'
import { describe, expect, it } from 'vitest'

function loadPatch(): Array<Record<string, unknown>> {
  const source = readFileSync(new URL('../../cordis.patch.yml', import.meta.url), 'utf8')
  // `!!js` scalars stay unevaluated strings: only the row shape is under test.
  const schema = yaml.DEFAULT_SCHEMA.extend([
    new yaml.Type('tag:yaml.org,2002:js', { kind: 'scalar', construct: (data: string) => data }),
  ])
  return yaml.load(source, { schema }) as Array<Record<string, unknown>>
}

interface Row {
  id?: string
  name?: string
  config?: Record<string, unknown>
  insert?: Row[]
}

function allRows(entries: Row[]): Row[] {
  const rows: Row[] = []
  for (const entry of entries) {
    if (Array.isArray(entry.insert)) rows.push(...allRows(entry.insert))
    else rows.push(entry)
  }
  return rows
}

/**
 * Every plugin row carries config that only its own package can validate. A 0.1
 * preset file was migrated here by hand, and `persona` kept its removed `text`
 * key, so the preset failed to mount with `$.prefix missing required value`.
 * These assertions run each row through the owning package's published Config.
 */
describe('PPT preset row configuration', () => {
  const rows = allRows(loadPatch() as Row[])
  const preset = rows.find(row => row.id === 'preset-ppt')
  const registry = rows.find(row => row.id === 'ppt-headless-agent-preset-registry')

  // Schemastery schemas are callable validators, but each package's generic signature
  // differs; the map is typed by the one contract the assertions below rely on.
  const pluginConfig = {
    '@deepseek-ai/dsh-persona': PersonaConfig,
    '@deepseek-ai/dsh-agent-tool-presentation': PresentationConfig,
    '@deepseek-ai/dsh-tool-bash': ToolBashConfig,
    '@deepseek-ai/dsh-tool-fs': ToolFsConfig,
    '@deepseek-ai/dsh-tool-ask-user': ToolAskUserConfig,
    '@deepseek-ai/dsh-tool-todo': ToolTodoConfig,
  } as unknown as Record<string, (value: unknown) => unknown>

  it('declares a definition the preset package accepts', () => {
    expect(preset).toBeDefined()
    expect(() => AgentPreset.Config(preset!.config as never)).not.toThrow()
    expect(preset!.config?.id).toBe('ppt')
  })

  it('gives every plugin row config that its own package accepts', () => {
    const plugins = preset!.config?.plugins as Row[]
    expect(plugins.length).toBeGreaterThan(0)
    const checked: string[] = []
    for (const row of plugins) {
      const validate = row.name === undefined ? undefined : pluginConfig[row.name]
      if (validate === undefined || row.config === undefined) continue
      expect(() => validate(row.config), `${String(row.id)} (${String(row.name)}) config must satisfy its package Config`).not.toThrow()
      checked.push(String(row.id))
    }
    // Guards against silently skipping a row whose Config stopped being exported.
    expect(checked).toEqual(['persona', 'tool-presentation', 'tool-todo'])
  })

  it('keys the persona prompt on prefix, the field 0.2 requires', () => {
    const persona = (preset!.config?.plugins as Row[]).find(row => row.id === 'persona')
    expect(persona?.config).toHaveProperty('prefix')
    expect(persona?.config).not.toHaveProperty('text')
    expect(String(persona?.config?.prefix)).toContain('{{cwd}}')
  })

  it('names a default the registry can resolve for headless', () => {
    expect(registry).toBeDefined()
    expect(registry!.name).toBe('@deepseek-ai/dsh-agent-preset-registry')
    expect(() => AgentPresetRegistry.Config(registry!.config)).not.toThrow()
    expect(registry!.config?.default).toBe('ppt')
  })
})
