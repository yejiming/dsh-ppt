/** PPT-aware replacement for DSH's one-shot headless runner. */
import { randomUUID } from 'node:crypto'
import type { Context } from '@deepseek-ai/cordis'
import { installModelSelection } from '@deepseek-ai/dsh-agent'
import type {} from '@deepseek-ai/dsh-agent-default-model'
import type {} from '@deepseek-ai/dsh-agent-preset-registry'
import { createUserMessage } from '@deepseek-ai/dsh-llm'
import { SessionId, type SessionEvent } from '@deepseek-ai/dsh-session'
import z from 'schemastery'
import { PptError } from './errors.ts'
import { PPT_MODE_TOOL_NAMES } from './schemas.ts'

export const name = 'ppt-headless-runner'
export const inject = ['agentDefaultModel', 'agentPresets', 'agents', 'sessions']

export interface Config {
  task: string
  presetId: string
}

export const Config = z.object({
  task: z.string().required(),
  presetId: z.string().pattern(/^[a-z0-9][a-z0-9-]*$/).default('ppt'),
})

interface HeadlessIo {
  stdout: { write(chunk: string): unknown }
  stderr: { write(chunk: string): unknown }
  exit(code: number): void
}

type TurnEndReason = Extract<SessionEvent, { type: 'turn/end' }>['data']['reason']

function summarize(events: readonly SessionEvent[], firstSeq: number): { text: string; reason?: TurnEndReason } {
  let started = false
  let text = ''
  let reason: TurnEndReason | undefined
  for (const event of events) {
    if (event.seq < firstSeq) continue
    if (event.type === 'turn/start') {
      started = true
      continue
    }
    if (!started) continue
    if (event.type === 'assistant/message') {
      const joined = event.data.message.content
        .filter(block => block.type === 'text')
        .map(block => block.text)
        .join('')
      if (joined !== '') text = joined
    }
    if (event.type === 'turn/end') reason = event.data.reason
  }
  return { text, ...(reason === undefined ? {} : { reason }) }
}

async function run(ctx: Context, config: Config, io: HeadlessIo): Promise<void> {
  await ctx.get('loader')?.await()
  const agents = ctx.get('agents')
  const defaultModel = ctx.get('agentDefaultModel')
  const presets = ctx.get('agentPresets')
  const sessions = ctx.get('sessions')
  if (agents === undefined || defaultModel === undefined || presets === undefined || sessions === undefined) return

  const selection = defaultModel.currentSelection()
  const preset = await presets.resolve(config.presetId)
  const { agent } = await agents.create({
    sessionId: SessionId(`session-${randomUUID()}`),
    meta: { cwd: process.cwd(), agentPreset: preset.id },
    agentOptions: { provider: selection.provider, model: selection.model },
    setup: async (agentCtx) => {
      installModelSelection(agentCtx, { current: selection, assembled: undefined })
      await presets.mount(agentCtx, preset.id)
    },
  })
  await agent.whenIdle()
  const allowed = new Set<string>(PPT_MODE_TOOL_NAMES)
  const visible = agent.ctx.tools.schemas(agent).map(tool => tool.name).sort()
  const unexpected = visible.filter(toolName => !allowed.has(toolName))
  const missing = PPT_MODE_TOOL_NAMES.filter(toolName => !visible.includes(toolName))
  if (unexpected.length > 0 || missing.length > 0) {
    throw new PptError(
      'PPT_CAPABILITY_UNAVAILABLE',
      `PPT headless tool surface mismatch; unexpected: ${unexpected.join(', ') || 'none'}; missing: ${missing.join(', ') || 'none'}`,
    )
  }
  const firstSeq = agent.session.seq
  agent.followup(createUserMessage({
    content: [{ type: 'text', text: config.task }],
    source: { kind: 'user' },
  }))
  await agent.whenIdle()
  await sessions.flush(agent.session)
  const outcome = summarize(agent.session.snapshotEvents(firstSeq), firstSeq)
  io.stdout.write(`${outcome.text}\n`)
  if (outcome.reason?.kind === 'error') io.stderr.write(`dsh: ${outcome.reason.error.code}: ${outcome.reason.error.message}\n`)
  io.exit(outcome.reason?.kind === 'completed' ? 0 : 1)
}

export function apply(ctx: Context, config: Config): void {
  const exit = ctx.get('appExit')
  if (exit === undefined) throw new Error('ppt-headless-runner: launcher appExit service is unavailable')
  const io: HeadlessIo = { stdout: process.stdout, stderr: process.stderr, exit }
  run(ctx, config, io).catch((error: unknown) => {
    io.stderr.write(`dsh: ${error instanceof Error ? error.message : String(error)}\n`)
    io.exit(1)
  })
}
