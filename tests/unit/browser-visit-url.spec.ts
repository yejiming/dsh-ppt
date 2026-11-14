import { mkdir, realpath, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import { BrowserRuntime } from '../../src/browser.ts'
import { isLocalFilesystemPath } from '../../src/paths.ts'
import { SessionResourceRegistry } from '../../src/session-resources.ts'
import { createTestWorkspace, type TestWorkspace } from '../helpers/workspace.ts'

/** Local resolution failures the filesystem branch can raise, in any platform's wording. */
const FILESYSTEM_CODES = ['PPT_PATH_INVALID', 'PPT_PATH_OUTSIDE_WORKSPACE']

const workspaces: TestWorkspace[] = []

afterEach(async () => {
  await Promise.all(workspaces.splice(0).map(workspace => workspace.cleanup()))
})

function visitUrl(workspace: string, input: string): Promise<URL> {
  const runtime = new BrowserRuntime(new SessionResourceRegistry()) as unknown as { resolveVisitUrl(workspace: string, input: string): Promise<URL> }
  return runtime.resolveVisitUrl(workspace, input)
}

async function codeOf(promise: Promise<unknown>): Promise<string | undefined> {
  try {
    await promise
  } catch (error) {
    return (error as { code?: string }).code
  }
  return undefined
}

async function fixture(): Promise<{ root: string; deck: string; outside: string }> {
  const workspace = await createTestWorkspace('dsh-ppt-visit-')
  workspaces.push(workspace)
  const output = join(workspace.root, 'ppt-output', 'deck')
  await mkdir(output, { recursive: true })
  const deck = join(output, 'deck.html')
  await writeFile(deck, '<!doctype html><html><body></body></html>')
  const outside = join(workspace.root, 'outside.html')
  await writeFile(outside, '<!doctype html><html><body></body></html>')
  return { root: workspace.root, deck, outside }
}

describe('local filesystem path detection', () => {
  it('treats drive rooted, UNC, and POSIX rooted inputs as filesystem paths', () => {
    for (const input of [
      'E:\\ppt\\deck\\deck.html',
      'c:/ppt/deck/deck.html',
      '\\\\server\\share\\deck.html',
      '//server/share/deck.html',
      '/tmp/deck/deck.html',
      '\\deck.html',
    ]) {
      expect(isLocalFilesystemPath(input), input).toBe(true)
    }
  })

  it('leaves URLs and workspace relative references to the URL parser', () => {
    for (const input of [
      'file:///tmp/deck/deck.html',
      'file:///E:/ppt/deck/deck.html',
      'https://example.com/deck.html',
      'http://example.com',
      'ftp://example.com/deck.html',
      'ppt-output/deck/deck.html',
      './deck.html',
      '../deck.html',
      'deck.html',
    ]) {
      expect(isLocalFilesystemPath(input), input).toBe(false)
    }
  })
})

describe('browser visit target resolution', () => {
  it('resolves the absolute path of a deck inside plugin output', async () => {
    const { root, deck } = await fixture()
    const url = await visitUrl(root, deck)
    expect(url.protocol).toBe('file:')
    expect(fileURLToPath(url)).toBe(await realpath(deck))
  })

  it('routes a Windows drive path to the filesystem branch instead of the URL policy', async () => {
    const { root } = await fixture()
    // Regression: `new URL('E:\\ppt\\deck.html')` parses as scheme `e:` and used to be
    // rejected as a non-HTTP URL before the filesystem branch could run.
    const code = await codeOf(visitUrl(root, 'E:\\ppt\\deck\\deck.html'))
    expect(FILESYSTEM_CODES).toContain(code)
  })

  it('routes a UNC path to the filesystem branch instead of the URL policy', async () => {
    const { root } = await fixture()
    const code = await codeOf(visitUrl(root, '\\\\server\\share\\deck.html'))
    expect(FILESYSTEM_CODES).toContain(code)
  })

  it('blocks an absolute path that leaves plugin output through the filesystem branch', async () => {
    const { root, outside } = await fixture()
    await expect(visitUrl(root, outside)).rejects.toMatchObject({ code: 'BROWSER_URL_BLOCKED', message: 'local HTML is outside plugin output' })
  })

  it('keeps resolving file URLs inside plugin output and blocking ones outside', async () => {
    const { root, deck, outside } = await fixture()
    const url = await visitUrl(root, pathToFileURL(deck).href)
    expect(fileURLToPath(url)).toBe(await realpath(deck))
    await expect(visitUrl(root, pathToFileURL(outside).href)).rejects.toMatchObject({ code: 'BROWSER_URL_BLOCKED' })
  })

  it('keeps blocking non-public HTTP(S) targets', async () => {
    const { root } = await fixture()
    await expect(visitUrl(root, 'http://127.0.0.1:8080/deck.html')).rejects.toMatchObject({ code: 'BROWSER_URL_BLOCKED' })
    await expect(visitUrl(root, 'https://user:secret@example.com/deck.html')).rejects.toMatchObject({ code: 'BROWSER_URL_BLOCKED' })
    await expect(visitUrl(root, 'ftp://example.com/deck.html')).rejects.toMatchObject({ code: 'BROWSER_URL_BLOCKED' })
  })

  it('keeps resolving workspace relative paths', async () => {
    const { root, deck } = await fixture()
    const url = await visitUrl(root, 'ppt-output/deck/deck.html')
    expect(url.protocol).toBe('file:')
    expect(fileURLToPath(url)).toBe(await realpath(deck))
    await expect(visitUrl(root, 'deck/deck.html')).rejects.toMatchObject({ code: 'PPT_PATH_INVALID' })
  })
})
