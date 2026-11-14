import { stat } from 'node:fs/promises'
import { basename, extname, join } from 'node:path'
import sharp from 'sharp'
import { atomicWriteFile } from './atomic.ts'
import { safeAssetFilename, SourceManifest, type SourceAssetRecord } from './artifacts.ts'
import { validatePublicHttpUrl } from './browser-security.ts'
import { PptError, throwIfAborted } from './errors.ts'
import { DEFAULT_LIMITS, boundedInteger } from './limits.ts'
import { resolveWorkspacePath, workspaceRelative } from './paths.ts'

export type ImageOrientation = 'landscape' | 'portrait' | 'square' | 'any'
export type BuiltinImageProvider = 'openverse' | 'wikimedia-commons'
/** Built-in providers keep literal names; the open branch lets an injected backend report its own name. */
export type ImageProvider = BuiltinImageProvider | (string & {})

export interface ImageCandidate {
  image_url: string
  source_page: string
  provider: ImageProvider
  title: string
  license: string
  license_verified: false
  thumbnail_url?: string
  width?: number
  height?: number
  mime_type?: string
  author?: string
  license_url?: string
  attribution?: string
}

export interface ImageSearchResult {
  query: string
  count: number
  orientation: ImageOrientation
  cache_hit: boolean
  providers_used: ImageProvider[]
  warnings: string[]
  results: ImageCandidate[]
}

/** A zero-configuration retrieval backend. Built-in providers are always tried first. */
export interface ImageSearchBackend {
  readonly name: string
  search(query: string, amount: number, orientation: ImageOrientation, signal?: AbortSignal): Promise<ImageCandidate[]>
}

export interface ImageSearchRuntimeOptions {
  /** Extra backends appended after Openverse and Wikimedia Commons; empty by default. */
  providers?: readonly ImageSearchBackend[]
  /** Opt back into the legacy hard failure when no provider is reachable. Defaults to false. */
  strictOnUnavailable?: boolean
}

export interface ImageSearchFallback {
  id: string
  tool: string
  summary: string
}

/** Actionable alternatives carried by a degraded result when the free providers deliver nothing. */
export const IMAGE_SEARCH_FALLBACKS: readonly ImageSearchFallback[] = [
  { id: 'local-assets', tool: 'read_image', summary: 'reuse a raster already frozen under assets/images or listed in assets/source-manifest.json' },
  { id: 'browser-capture', tool: 'browser_visit', summary: 'capture the visual anchor from a public page through the browser instead of the image providers' },
  { id: 'custom-provider', tool: 'image_search', summary: 'inject an extra zero-configuration backend through the ImageSearchRuntime providers option' },
]

export type ImageSearchStatus = 'ok' | 'partial' | 'unavailable'

export interface ImageSearchDegradation {
  status: ImageSearchStatus
  requested: number
  returned: number
  failures: Array<{ provider: string; kind: string }>
  fallbacks: ImageSearchFallback[]
}

interface CacheEntry { expires: number; result: Omit<ImageSearchResult, 'cache_hit'> }

const FAILURE_KINDS = new Set(['cancelled', 'timeout', 'rate_limited', 'server_error', 'invalid_response', 'network_error'])
const DEGRADED_UNAVAILABLE = 'degraded:unavailable'
const DEGRADED_EMPTY = 'degraded:no_results'
const DEGRADED_PARTIAL = 'degraded:partial'
const PROVIDER_NAME_PATTERN = /^[a-z0-9][a-z0-9-]{0,31}$/

function fallbackWarnings(): string[] {
  return IMAGE_SEARCH_FALLBACKS.map(item => `fallback:${item.id}`)
}

/**
 * Rebuilds the structured story behind the machine-readable warnings of one search result.
 * It takes the structural subset it actually reads, so a payload derived from the tool
 * output schema (where `license_verified` widens to `boolean`) stays assignable.
 */
export function describeImageSearchDegradation(
  result: {
    readonly count: number
    readonly providers_used: readonly string[]
    readonly warnings: readonly string[]
    readonly results: readonly unknown[]
  },
): ImageSearchDegradation {
  const failures = result.warnings.flatMap((warning): Array<{ provider: string; kind: string }> => {
    const separator = warning.indexOf(':')
    if (separator <= 0) return []
    const provider = warning.slice(0, separator)
    const kind = warning.slice(separator + 1)
    if (!result.providers_used.includes(provider) || !FAILURE_KINDS.has(kind)) return []
    return [{ provider, kind }]
  })
  const unavailable = result.warnings.includes(DEGRADED_UNAVAILABLE) || result.results.length === 0
  const status: ImageSearchStatus = unavailable
    ? 'unavailable'
    : failures.length > 0 || result.results.length < result.count ? 'partial' : 'ok'
  return {
    status,
    requested: result.count,
    returned: result.results.length,
    failures,
    fallbacks: IMAGE_SEARCH_FALLBACKS.filter(item => result.warnings.includes(`fallback:${item.id}`)),
  }
}

type FetchLike = (input: string | URL | Request, init?: RequestInit) => Promise<Response>
type UrlValidator = (input: string) => Promise<URL>

const OPENVERSE_ENDPOINT = 'https://api.openverse.org/v1/images/'
const COMMONS_ENDPOINT = 'https://commons.wikimedia.org/w/api.php'
const MAX_PROVIDER_BYTES = 2 * 1024 * 1024
const ADULT_PATTERN = /(?:\bporn\b|\bnsfw\b|\bsexually explicit\b|\bnude\b|色情|成人内容|裸体)/iu

function text(value: unknown, max = 500): string | undefined {
  if (typeof value !== 'string') return undefined
  const normalized = value.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()
  return normalized.length === 0 ? undefined : normalized.slice(0, max)
}

function positiveInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isInteger(value) && value > 0 ? value : undefined
}

function matchesOrientation(width: number | undefined, height: number | undefined, orientation: ImageOrientation): boolean {
  if (orientation === 'any' || width === undefined || height === undefined) return true
  const ratio = width / height
  if (orientation === 'square') return ratio >= 0.9 && ratio <= 1.1
  return orientation === 'landscape' ? ratio > 1.1 : ratio < 0.9
}

async function readBoundedBytes(response: Response, limit: number): Promise<Uint8Array> {
  const declared = Number(response.headers.get('content-length') ?? 0)
  if (Number.isFinite(declared) && declared > limit) throw new PptError('PPT_RESOURCE_LIMIT', `response exceeds ${limit} bytes`)
  if (response.body === null) {
    const bytes = new Uint8Array(await response.arrayBuffer())
    if (bytes.byteLength > limit) throw new PptError('PPT_RESOURCE_LIMIT', `response exceeds ${limit} bytes`)
    return bytes
  }
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  try {
    while (true) {
      const chunk = await reader.read()
      if (chunk.done) break
      total += chunk.value.byteLength
      if (total > limit) {
        await reader.cancel('response size limit exceeded').catch(() => undefined)
        throw new PptError('PPT_RESOURCE_LIMIT', `response exceeds ${limit} bytes`)
      }
      chunks.push(chunk.value)
    }
  } finally {
    reader.releaseLock()
  }
  const bytes = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength }
  return bytes
}

async function readJson(response: Response): Promise<unknown> {
  const bytes = await readBoundedBytes(response, MAX_PROVIDER_BYTES)
  return JSON.parse(new TextDecoder().decode(bytes))
}

async function safeCandidate(
  candidate: ImageCandidate,
  orientation: ImageOrientation,
  validateUrl: UrlValidator,
): Promise<ImageCandidate | undefined> {
  if (ADULT_PATTERN.test(`${candidate.title} ${candidate.attribution ?? ''}`)) return undefined
  if (!matchesOrientation(candidate.width, candidate.height, orientation)) return undefined
  try {
    const [image, source] = await Promise.all([validateUrl(candidate.image_url), validateUrl(candidate.source_page)])
    let thumbnail: string | undefined
    if (candidate.thumbnail_url !== undefined) thumbnail = (await validateUrl(candidate.thumbnail_url)).href
    return { ...candidate, image_url: image.href, source_page: source.href, ...(thumbnail === undefined ? {} : { thumbnail_url: thumbnail }) }
  } catch {
    return undefined
  }
}

function upstreamCategory(error: unknown): string {
  if (error instanceof PptError && error.code === 'PPT_ABORTED') return 'cancelled'
  if (error instanceof DOMException && ['AbortError', 'TimeoutError'].includes(error.name)) return 'timeout'
  const message = error instanceof Error ? error.message : String(error)
  if (/HTTP 429/.test(message)) return 'rate_limited'
  if (/HTTP 5\d\d/.test(message)) return 'server_error'
  if (/JSON|parse/i.test(message)) return 'invalid_response'
  return 'network_error'
}

export class ImageSearchRuntime {
  private readonly cache = new Map<string, CacheEntry>()
  private readonly backends: readonly ImageSearchBackend[]
  private readonly strictOnUnavailable: boolean

  constructor(
    private readonly fetcher: FetchLike = fetch,
    private readonly validateUrl: UrlValidator = validatePublicHttpUrl,
    options: ImageSearchRuntimeOptions = {},
  ) {
    const injected = options.providers ?? []
    for (const backend of injected) {
      if (!PROVIDER_NAME_PATTERN.test(backend.name) || backend.name === 'openverse' || backend.name === 'wikimedia-commons') {
        throw new PptError('IMAGE_SEARCH_FAILED', `invalid extra image provider name: ${backend.name}`)
      }
    }
    this.backends = [
      { name: 'openverse', search: (query, amount, orientation, signal) => this.openverse(query, amount, orientation, signal) },
      { name: 'wikimedia-commons', search: (query, amount, orientation, signal) => this.commons(query, amount, orientation, signal) },
      ...injected,
    ]
    this.strictOnUnavailable = options.strictOnUnavailable === true
  }

  async search(queryInput: string, countInput = 8, orientation: ImageOrientation = 'any', signal?: AbortSignal): Promise<ImageSearchResult> {
    throwIfAborted(signal)
    const query = queryInput.normalize('NFKC').replace(/\s+/g, ' ').trim()
    if ([...query].length < 1 || [...query].length > 160) throw new PptError('IMAGE_SEARCH_FAILED', 'query must contain 1..160 Unicode code points')
    const count = boundedInteger(countInput, 'count', 1, 12)
    if (!['landscape', 'portrait', 'square', 'any'].includes(orientation)) throw new PptError('IMAGE_SEARCH_FAILED', `unsupported orientation: ${orientation}`)
    if (ADULT_PATTERN.test(query)) throw new PptError('IMAGE_SEARCH_FAILED', 'adult-content queries are blocked')
    const key = JSON.stringify([query.toLocaleLowerCase(), count, orientation])
    const cached = this.cache.get(key)
    if (cached !== undefined && cached.expires > Date.now()) return { ...structuredClone(cached.result), cache_hit: true }

    const providersUsed: ImageProvider[] = []
    const warnings: string[] = []
    const failures: Array<{ provider: string; kind: string }> = []
    const candidates: ImageCandidate[] = []
    for (const backend of this.backends) {
      if (candidates.length >= count) break
      providersUsed.push(backend.name)
      try {
        candidates.push(...await backend.search(query, Math.min(40, count * 3), orientation, signal))
      } catch (error) {
        if (signal?.aborted) throwIfAborted(signal)
        const kind = upstreamCategory(error)
        failures.push({ provider: backend.name, kind })
        warnings.push(`${backend.name}:${kind}`)
      }
    }
    const seen = new Set<string>()
    const results = candidates.filter(item => !seen.has(item.image_url) && seen.add(item.image_url)).slice(0, count)
    if (results.length === 0) {
      const unreachable = providersUsed.length > 0 && failures.length >= providersUsed.length
      warnings.push(unreachable ? DEGRADED_UNAVAILABLE : DEGRADED_EMPTY)
      warnings.push(...fallbackWarnings())
      warnings.push(`insufficient_results:0/${count}`)
      if (this.strictOnUnavailable && unreachable) {
        throw new PptError('IMAGE_SEARCH_FAILED', 'no zero-configuration image provider is reachable', {
          details: { status: 'unavailable', failures, fallbacks: IMAGE_SEARCH_FALLBACKS },
        })
      }
      // A degraded answer is never cached, so the next call really retries the providers.
      return { query, count, orientation, cache_hit: false, providers_used: providersUsed, warnings, results }
    }
    if (failures.length > 0) warnings.push(DEGRADED_PARTIAL)
    if (results.length < count) warnings.push(`insufficient_results:${results.length}/${count}`)
    const stored = { query, count, orientation, providers_used: providersUsed, warnings, results }
    this.cache.set(key, { expires: Date.now() + 10 * 60_000, result: structuredClone(stored) })
    return { ...stored, cache_hit: false }
  }

  private async request(url: URL, signal?: AbortSignal): Promise<unknown> {
    const timeout = AbortSignal.timeout(15_000)
    const combined = signal === undefined ? timeout : AbortSignal.any([signal, timeout])
    const response = await this.fetcher(url, { signal: combined, redirect: 'error', headers: { accept: 'application/json' } })
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    return readJson(response)
  }

  private async openverse(query: string, amount: number, orientation: ImageOrientation, signal?: AbortSignal): Promise<ImageCandidate[]> {
    const url = new URL(OPENVERSE_ENDPOINT)
    url.searchParams.set('q', query)
    url.searchParams.set('page_size', String(amount))
    url.searchParams.set('mature', 'false')
    const body = await this.request(url, signal) as { results?: unknown[] }
    const raw = Array.isArray(body.results) ? body.results : []
    const candidates = await Promise.all(raw.map(async (item) => {
      const row = item as Record<string, unknown>
      if (row.mature === true) return undefined
      const image = text(row.url, 2_048)
      const source = text(row.foreign_landing_url, 2_048)
      if (image === undefined || source === undefined) return undefined
      return safeCandidate({
        image_url: image, source_page: source, provider: 'openverse', title: text(row.title) ?? 'Untitled image',
        license: text(row.license) ?? 'unknown', license_verified: false,
        ...(text(row.thumbnail, 2_048) === undefined ? {} : { thumbnail_url: text(row.thumbnail, 2_048)! }),
        ...(positiveInteger(row.width) === undefined ? {} : { width: positiveInteger(row.width)! }),
        ...(positiveInteger(row.height) === undefined ? {} : { height: positiveInteger(row.height)! }),
        ...(text(row.mime_type, 100) === undefined ? {} : { mime_type: text(row.mime_type, 100)! }),
        ...(text(row.creator) === undefined ? {} : { author: text(row.creator)! }),
        ...(text(row.license_url, 2_048) === undefined ? {} : { license_url: text(row.license_url, 2_048)! }),
        ...(text(row.attribution) === undefined ? {} : { attribution: text(row.attribution)! }),
      }, orientation, this.validateUrl)
    }))
    return candidates.filter((item): item is ImageCandidate => item !== undefined)
  }

  private async commons(query: string, amount: number, orientation: ImageOrientation, signal?: AbortSignal): Promise<ImageCandidate[]> {
    const url = new URL(COMMONS_ENDPOINT)
    for (const [key, value] of Object.entries({
      action: 'query', format: 'json', origin: '*', generator: 'search', gsrnamespace: '6',
      gsrsearch: `file:${query}`, gsrlimit: String(amount), prop: 'imageinfo|info', inprop: 'url',
      iiprop: 'url|size|mime|extmetadata', iiurlwidth: '1280',
    })) url.searchParams.set(key, value)
    const body = await this.request(url, signal) as { query?: { pages?: Record<string, unknown> } }
    const pages = Object.values(body.query?.pages ?? {})
    const candidates = await Promise.all(pages.map(async (item) => {
      const page = item as Record<string, unknown>
      const info = Array.isArray(page.imageinfo) ? page.imageinfo[0] as Record<string, unknown> | undefined : undefined
      if (info === undefined) return undefined
      const meta = (info.extmetadata ?? {}) as Record<string, { value?: unknown }>
      const image = text(info.url, 2_048)
      const source = text(page.fullurl, 2_048)
      if (image === undefined || source === undefined) return undefined
      return safeCandidate({
        image_url: image, source_page: source, provider: 'wikimedia-commons',
        title: text(meta.ObjectName?.value) ?? text(page.title) ?? 'Wikimedia Commons image',
        license: text(meta.LicenseShortName?.value) ?? 'unknown', license_verified: false,
        ...(text(info.thumburl, 2_048) === undefined ? {} : { thumbnail_url: text(info.thumburl, 2_048)! }),
        ...(positiveInteger(info.width) === undefined ? {} : { width: positiveInteger(info.width)! }),
        ...(positiveInteger(info.height) === undefined ? {} : { height: positiveInteger(info.height)! }),
        ...(text(info.mime, 100) === undefined ? {} : { mime_type: text(info.mime, 100)! }),
        ...(text(meta.Artist?.value) === undefined ? {} : { author: text(meta.Artist?.value)! }),
        ...(text(meta.LicenseUrl?.value, 2_048) === undefined ? {} : { license_url: text(meta.LicenseUrl?.value, 2_048)! }),
        ...(text(meta.Attribution?.value) === undefined ? {} : { attribution: text(meta.Attribution?.value)! }),
      }, orientation, this.validateUrl)
    }))
    return candidates.filter((item): item is ImageCandidate => item !== undefined)
  }
}

export interface FrozenImageAsset {
  path: string
  width: number
  height: number
  mime_type: string
  size: number
  manifest: SourceAssetRecord
}

export async function freezeImageAsset(
  workspace: string,
  artifactRoot: string,
  candidate: ImageCandidate,
  fetcher: FetchLike = fetch,
  signal?: AbortSignal,
): Promise<FrozenImageAsset> {
  throwIfAborted(signal)
  const root = await resolveWorkspacePath(workspace, artifactRoot, { mustExist: true, kind: 'directory' })
  const images = await resolveWorkspacePath(workspace, join(root, 'assets', 'images'), { mustExist: true, kind: 'directory' })
  const [url, sourcePage] = await Promise.all([
    validatePublicHttpUrl(candidate.image_url),
    validatePublicHttpUrl(candidate.source_page),
  ])
  const response = await fetcher(url, { signal, redirect: 'error' })
  if (!response.ok) throw new PptError('IMAGE_ASSET_INVALID', `image download failed with HTTP ${response.status}`)
  const buffer = await readBoundedBytes(response, DEFAULT_LIMITS.maxImageBytes)
  const metadata = await sharp(buffer, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).metadata().catch(error => {
    throw new PptError('IMAGE_ASSET_INVALID', 'downloaded file is not a supported image', { cause: error })
  })
  const formats: Record<string, { mime: string; extension: string }> = {
    png: { mime: 'image/png', extension: '.png' }, jpeg: { mime: 'image/jpeg', extension: '.jpg' },
    webp: { mime: 'image/webp', extension: '.webp' }, gif: { mime: 'image/gif', extension: '.gif' },
  }
  const actual = metadata.format === undefined ? undefined : formats[metadata.format]
  if (actual === undefined || metadata.width === undefined || metadata.height === undefined) {
    throw new PptError('IMAGE_ASSET_INVALID', `unsupported image format: ${metadata.format ?? 'unknown'}`)
  }
  if (metadata.width * metadata.height > DEFAULT_LIMITS.maxImagePixels) throw new PptError('PPT_RESOURCE_LIMIT', 'image pixel count exceeds limit')
  const suggested = safeAssetFilename(candidate.title, url)
  const stem = basename(suggested, extname(suggested))
  let target = join(images, `${stem}${actual.extension}`)
  for (let suffix = 2; ; suffix += 1) {
    try { await stat(target); target = join(images, `${stem}-${suffix}${actual.extension}`) } catch { break }
  }
  await atomicWriteFile(target, buffer, { signal })
  const manifest = await new SourceManifest(workspace, join(root, 'assets', 'source-manifest.json')).append({
    original_url: url.href, source_page: sourcePage.href, fetched_at: new Date().toISOString(),
    author: candidate.author ?? null, license: candidate.license || 'unknown', license_url: candidate.license_url ?? null,
    localFile: target,
  })
  return {
    path: workspaceRelative(workspace, target), width: metadata.width, height: metadata.height,
    mime_type: actual.mime, size: buffer.byteLength, manifest,
  }
}
