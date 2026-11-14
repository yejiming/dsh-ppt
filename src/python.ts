import { readdir, stat } from 'node:fs/promises'
import { extname, join } from 'node:path'
import type { SandboxProvider } from '@deepseek-ai/dsh-sandbox'
import type { SubprocessRuntime } from '@deepseek-ai/dsh-subprocess'
import sharp from 'sharp'
import { PptError, asPptError, throwIfAborted } from './errors.ts'
import { DEFAULT_LIMITS, boundedInteger } from './limits.ts'
import { resolveWorkspacePath, workspaceRelative } from './paths.ts'
import type { SessionOwner } from './session-resources.ts'
import { SessionResourceRegistry } from './session-resources.ts'
import { runCollected } from './subprocess.ts'

export interface PythonArtifact {
  path: string
  size: number
  mime_type: string
  width?: number
  height?: number
}

export interface PythonExecutionInput {
  code: string
  cwd?: string
  timeout_ms?: number
  expected_outputs?: string[]
}

export interface PythonExecutionResult {
  exit_code: number
  stdout: string
  stderr: string
  stdout_truncated: boolean
  stderr_truncated: boolean
  duration_ms: number
  artifacts: PythonArtifact[]
}

interface FileStamp { size: number; mtimeMs: number }

const MIME_BY_EXTENSION: Readonly<Record<string, string>> = Object.freeze({
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.webp': 'image/webp',
  '.gif': 'image/gif', '.svg': 'image/svg+xml', '.csv': 'text/csv', '.json': 'application/json',
  '.txt': 'text/plain', '.pdf': 'application/pdf',
})

async function snapshotFiles(root: string, signal?: AbortSignal): Promise<Map<string, FileStamp>> {
  const files = new Map<string, FileStamp>()
  const queue = [root]
  let seen = 0
  while (queue.length > 0) {
    throwIfAborted(signal)
    const directory = queue.pop()!
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      if (entry.name === '.git' || entry.isSymbolicLink()) continue
      const path = join(directory, entry.name)
      if (entry.isDirectory()) queue.push(path)
      else if (entry.isFile()) {
        seen += 1
        if (seen > 10_000) throw new PptError('PPT_RESOURCE_LIMIT', 'Python workspace scan exceeds 10000 files')
        const info = await stat(path)
        files.set(path, { size: info.size, mtimeMs: info.mtimeMs })
      }
    }
  }
  return files
}

async function artifactMetadata(workspace: string, path: string): Promise<PythonArtifact> {
  const info = await stat(path)
  if (info.size > DEFAULT_LIMITS.maxGeneratedFileBytes) {
    throw new PptError('PPT_RESOURCE_LIMIT', `${workspaceRelative(workspace, path)} exceeds the generated-file size limit`)
  }
  const mime = MIME_BY_EXTENSION[extname(path).toLowerCase()] ?? 'application/octet-stream'
  const base = { path: workspaceRelative(workspace, path), size: info.size, mime_type: mime }
  if (!mime.startsWith('image/')) return base
  try {
    const metadata = await sharp(path, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).metadata()
    return {
      ...base,
      ...(metadata.width === undefined ? {} : { width: metadata.width }),
      ...(metadata.height === undefined ? {} : { height: metadata.height }),
    }
  } catch (error) {
    throw new PptError('PYTHON_EXECUTION_FAILED', `generated image is invalid: ${base.path}`, { cause: error })
  }
}

export class PythonRuntime {
  constructor(
    private readonly subprocess: SubprocessRuntime | undefined,
    private readonly sandbox: SandboxProvider | undefined,
    private readonly resources: SessionResourceRegistry,
    private readonly executable = process.platform === 'win32' ? 'python' : 'python3',
  ) {}

  async execute(owner: SessionOwner, workspace: string, input: PythonExecutionInput, signal?: AbortSignal): Promise<PythonExecutionResult> {
    throwIfAborted(signal)
    if (this.subprocess === undefined || this.sandbox === undefined) {
      throw new PptError('PPT_CAPABILITY_UNAVAILABLE', 'Python requires DSH subprocess and sandbox services')
    }
    if (input.code.trim().length === 0 || [...input.code].length > 200_000) {
      throw new PptError('PYTHON_EXECUTION_FAILED', 'Python code must contain 1..200000 Unicode code points')
    }
    const cwd = await resolveWorkspacePath(workspace, input.cwd ?? '.', { mustExist: true, kind: 'directory' })
    const expected = await Promise.all((input.expected_outputs ?? []).map(path => resolveWorkspacePath(workspace, path)))
    if (expected.length > DEFAULT_LIMITS.maxGeneratedFiles) throw new PptError('PPT_RESOURCE_LIMIT', 'too many expected Python outputs')
    const timeoutMs = boundedInteger(input.timeout_ms ?? DEFAULT_LIMITS.maxPythonMs, 'timeout_ms', 1_000, DEFAULT_LIMITS.maxPythonMs)
    this.resources.open(owner, workspace)
    const resolvedExecutable = await this.subprocess.resolveExecutable(this.executable, undefined, signal)
    const environment = { MPLBACKEND: 'Agg', PYTHONNOUSERSITE: '1', PYTHONDONTWRITEBYTECODE: '1' }
    const preflight = await this.sandbox.confine(
      [resolvedExecutable, '-c', 'import matplotlib, PIL, cv2; matplotlib.use("Agg"); print("ready")'],
      { mode: 'read-only', workspaceRoot: workspace },
    )
    if (preflight.enforcement !== 'full') throw new PptError('PPT_CAPABILITY_UNAVAILABLE', 'Python sandbox enforcement is partial')
    let dependency
    try {
      dependency = await runCollected(this.subprocess, preflight.argv, {
        cwd, env: environment, signal, timeoutMs: Math.min(timeoutMs, 15_000), maxOutputBytes: 8_192,
      })
    } catch (error) {
      if (signal?.aborted) throwIfAborted(signal)
      throw asPptError(error, 'PYTHON_DEPENDENCY_MISSING', 'Python dependency preflight failed')
    }
    if (dependency.exitCode !== 0) {
      throw new PptError('PYTHON_DEPENDENCY_MISSING', 'Python requires matplotlib, Pillow, and opencv-python', {
        details: { stderr: dependency.stderr.slice(0, 2_000), executable: resolvedExecutable },
      })
    }

    const before = await snapshotFiles(cwd, signal)
    const source = [
      'import os',
      'os.environ["MPLBACKEND"] = "Agg"',
      'import matplotlib',
      'matplotlib.use("Agg")',
      'import PIL, cv2',
      input.code,
    ].join('\n')
    const confined = await this.sandbox.confine([resolvedExecutable, '-'], { mode: 'workspace-write', workspaceRoot: workspace })
    if (confined.enforcement !== 'full') throw new PptError('PPT_CAPABILITY_UNAVAILABLE', 'Python sandbox enforcement is partial')
    const started = Date.now()
    let result
    try {
      result = await runCollected(this.subprocess, confined.argv, {
        cwd, env: environment, signal, stdin: source, timeoutMs,
        maxOutputBytes: DEFAULT_LIMITS.maxPythonOutputChars * 4,
      })
    } catch (error) {
      if (signal?.aborted) throwIfAborted(signal)
      throw asPptError(error, 'PYTHON_EXECUTION_FAILED', 'Python execution failed')
    }
    const after = await snapshotFiles(cwd, signal)
    const changed = [...after.entries()].filter(([path, stamp]) => {
      const old = before.get(path)
      return old === undefined || old.size !== stamp.size || old.mtimeMs !== stamp.mtimeMs
    }).map(([path]) => path)
    if (changed.length > DEFAULT_LIMITS.maxGeneratedFiles) {
      throw new PptError('PPT_RESOURCE_LIMIT', `Python generated more than ${DEFAULT_LIMITS.maxGeneratedFiles} files`)
    }
    for (const path of expected) {
      if (!after.has(path)) throw new PptError('PYTHON_EXECUTION_FAILED', `expected output was not created: ${workspaceRelative(workspace, path)}`)
    }
    const artifacts = await Promise.all(changed.sort().map(path => artifactMetadata(workspace, path)))
    const exitCode = result.exitCode ?? -1
    if (exitCode !== 0) {
      throw new PptError('PYTHON_EXECUTION_FAILED', `Python exited with code ${exitCode}`, {
        details: { exit_code: exitCode, stdout: result.stdout, stderr: result.stderr },
      })
    }
    return {
      exit_code: exitCode,
      stdout: result.stdout.slice(0, DEFAULT_LIMITS.maxPythonOutputChars),
      stderr: result.stderr.slice(0, DEFAULT_LIMITS.maxPythonOutputChars),
      stdout_truncated: result.stdoutTruncated || result.stdout.length > DEFAULT_LIMITS.maxPythonOutputChars,
      stderr_truncated: result.stderrTruncated || result.stderr.length > DEFAULT_LIMITS.maxPythonOutputChars,
      duration_ms: Date.now() - started,
      artifacts,
    }
  }
}
