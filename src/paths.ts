import { lstat, mkdir, realpath } from 'node:fs/promises'
import { dirname, isAbsolute, relative, resolve, sep } from 'node:path'
import { PptError } from './errors.ts'

export interface ResolveWorkspacePathOptions {
  mustExist?: boolean
  kind?: 'file' | 'directory' | 'either'
  createParent?: boolean
}

// A Windows drive rooted path (`E:\dir\file`, `c:/dir/file`) or any rooted path
// (`/dir/file`, `\\server\share\file`, `//server/share/file`). Platform independent on
// purpose: on Windows a drive path would otherwise be parsed by `new URL` as an `e:`
// scheme and rejected as a non-HTTP URL before the filesystem branch is reached.
const DRIVE_ROOTED_PATH = /^[A-Za-z]:[\\/]/
const ROOTED_PATH = /^[\\/]/

export function isLocalFilesystemPath(input: string): boolean {
  return DRIVE_ROOTED_PATH.test(input) || ROOTED_PATH.test(input)
}

export function isPathInside(root: string, target: string): boolean {
  const rel = relative(resolve(root), resolve(target))
  return rel === '' || (!rel.startsWith(`..${sep}`) && rel !== '..' && !isAbsolute(rel))
}

export function workspaceRelative(root: string, target: string): string {
  if (!isPathInside(root, target)) {
    throw new PptError('PPT_PATH_OUTSIDE_WORKSPACE', `path is outside workspace: ${target}`)
  }
  return relative(resolve(root), resolve(target)).split(sep).join('/') || '.'
}

async function nearestExistingParent(path: string): Promise<string> {
  let cursor = resolve(path)
  for (;;) {
    try {
      return await realpath(cursor)
    } catch {
      const parent = dirname(cursor)
      if (parent === cursor) throw new PptError('PPT_PATH_INVALID', `cannot resolve path parent: ${path}`)
      cursor = parent
    }
  }
}

export async function resolveWorkspacePath(
  workspaceRoot: string,
  input: string,
  options: ResolveWorkspacePathOptions = {},
): Promise<string> {
  if (typeof input !== 'string' || input.trim().length === 0 || input.includes('\0')) {
    throw new PptError('PPT_PATH_INVALID', 'path must be a non-empty string without NUL bytes')
  }
  const root = await realpath(resolve(workspaceRoot))
  const candidate = resolve(root, input)
  if (!isPathInside(root, candidate)) {
    throw new PptError('PPT_PATH_OUTSIDE_WORKSPACE', `path is outside workspace: ${input}`)
  }

  const existingParent = await nearestExistingParent(candidate)
  if (!isPathInside(root, existingParent)) {
    throw new PptError('PPT_PATH_OUTSIDE_WORKSPACE', `path resolves outside workspace through a symlink: ${input}`)
  }

  if (options.createParent) await mkdir(dirname(candidate), { recursive: true })
  if (options.mustExist) {
    let stat
    try {
      stat = await lstat(candidate)
    } catch (error) {
      throw new PptError('PPT_PATH_INVALID', `path does not exist: ${input}`, { cause: error })
    }
    const kind = options.kind ?? 'either'
    if (kind === 'file' && !stat.isFile()) {
      throw new PptError('PPT_PATH_INVALID', `path is not a regular file: ${input}`)
    }
    if (kind === 'directory' && !stat.isDirectory()) {
      throw new PptError('PPT_PATH_INVALID', `path is not a directory: ${input}`)
    }
    const resolvedExisting = await realpath(candidate)
    if (!isPathInside(root, resolvedExisting)) {
      throw new PptError('PPT_PATH_OUTSIDE_WORKSPACE', `path resolves outside workspace: ${input}`)
    }
    return resolvedExisting
  }
  return candidate
}
