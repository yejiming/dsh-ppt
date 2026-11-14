/** DeepSeek Harness host-plane entry for the PPT design plugin. */
import type { Context } from '@deepseek-ai/cordis'
import type {} from '@deepseek-ai/dsh-agent-preset-registry'
import type {} from '@deepseek-ai/dsh-sandbox'
import type {} from '@deepseek-ai/dsh-subprocess'
import z from 'schemastery'
import { PptError, PPT_ERROR_CODES, type PptErrorCode } from './errors.ts'
import { isSupportedPlatform } from './platform.ts'
import { createPptRuntime, type PptRuntime, type PptRuntimeOptions } from './runtime.ts'
import { PPT_MODE_TOOL_NAMES, PPT_TOOL_NAMES, type PptToolName } from './schemas.ts'

declare module '@deepseek-ai/cordis' {
  interface Context {
    pptRuntime: PptRuntime
  }
}

export { createPptRuntime, PptError, PPT_ERROR_CODES, PPT_MODE_TOOL_NAMES, PPT_TOOL_NAMES }
export type { PptErrorCode, PptRuntime, PptRuntimeOptions, PptToolName }

export const name = 'dsh-ppt'
export const inject = ['sandbox', 'subprocess']

export interface Config {
  presetId: string
  pythonExecutable: string
  browserExecutable: string
  fontDirs: string[]
  outputRoot: string
}

export const Config = z.object({
  presetId: z.string().pattern(/^[a-z0-9][a-z0-9-]*$/).default('ppt'),
  pythonExecutable: z.string().default(process.platform === 'win32' ? 'python' : 'python3'),
  browserExecutable: z.string().default(''),
  fontDirs: z.array(z.string()).default([]),
  outputRoot: z.string().default('ppt-output'),
})

export function assertSupportedPlatform(platform: NodeJS.Platform = process.platform): void {
  if (!isSupportedPlatform(platform)) {
    throw new PptError('PPT_PLATFORM_UNSUPPORTED', `PPT mode supports macOS, Linux, and Windows; unsupported platform: ${platform}`)
  }
}

/**
 * Provide the host-plane PPT runtime. The model-facing PPT preset is declared by
 * this package's `cordis.patch.yml`, because the DSH agent-preset registry
 * composes presets from ordinary plugin rows and no longer scans
 * `<dshHome>/.agent-presets`.
 */
export async function apply(ctx: Context, config: Config): Promise<void> {
  assertSupportedPlatform()
  const runtime = createPptRuntime({
    context: ctx,
    pythonExecutable: config.pythonExecutable,
    browserExecutable: config.browserExecutable || undefined,
    fontDirs: config.fontDirs,
    outputRoot: config.outputRoot,
  })
  ctx.provide('pptRuntime', runtime)
  ctx.effect(() => () => runtime.dispose(), 'dsh-ppt: dispose runtime resources')
}
