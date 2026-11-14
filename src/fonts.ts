import { createHash } from 'node:crypto'
import { opendir, readFile } from 'node:fs/promises'
import { extname, join } from 'node:path'
import * as fontkit from 'fontkit'
import { PptError } from './errors.ts'
import type { InstalledFontFace } from './font-files.ts'
import { systemFontDirectories } from './platform.ts'

export interface FontDescriptor {
  name: string
  aliases?: readonly string[]
  language: string
  style: string
  characteristics: string
  layer: FontLayer
  platforms: readonly SupportedFontPlatform[]
  roles: readonly FontRole[]
}

export type FontLayer = 'portable' | 'system' | 'custom'
export type SupportedFontPlatform = 'darwin' | 'win32' | 'linux'
export type FontRole = 'latin-sans' | 'latin-serif' | 'cjk-sans' | 'cjk-serif' | 'display' | 'code'
export const FONT_LAYERS: readonly FontLayer[] = ['portable', 'system', 'custom']
export const FONT_ROLES: readonly FontRole[] = ['latin-sans', 'latin-serif', 'cjk-sans', 'cjk-serif', 'display', 'code']

const ALL_PLATFORMS: readonly SupportedFontPlatform[] = ['darwin', 'win32', 'linux']

export const FONT_REGISTRY: readonly FontDescriptor[] = Object.freeze([
  { name: 'Arial', aliases: ['ArialMT'], language: 'Multi-language', style: 'Sans-serif', characteristics: 'Portable Office-safe sans-serif', layer: 'portable', platforms: ALL_PLATFORMS, roles: ['latin-sans'] },
  { name: 'Times New Roman', aliases: ['TimesNewRomanPSMT'], language: 'Multi-language', style: 'Serif', characteristics: 'Portable Office-safe serif', layer: 'portable', platforms: ALL_PLATFORMS, roles: ['latin-serif'] },
  { name: 'Segoe UI', aliases: ['SegoeUI'], language: 'Western', style: 'Sans-serif', characteristics: 'Windows interface sans-serif', layer: 'system', platforms: ['win32'], roles: ['latin-sans'] },
  { name: 'Microsoft YaHei', aliases: ['Microsoft YaHei UI', 'MicrosoftYaHei', 'MicrosoftYaHeiUI'], language: 'Chinese + Western', style: 'Sans-serif', characteristics: 'Windows ClearType Chinese sans-serif', layer: 'system', platforms: ['win32'], roles: ['cjk-sans', 'latin-sans'] },
  { name: 'DengXian', aliases: ['Deng'], language: 'Chinese + Western', style: 'Sans-serif', characteristics: 'Windows modern Chinese sans-serif', layer: 'system', platforms: ['win32'], roles: ['cjk-sans', 'latin-sans'] },
  { name: 'SimSun', aliases: ['NSimSun'], language: 'Chinese + Western', style: 'Serif', characteristics: 'Windows Song-style Chinese serif', layer: 'system', platforms: ['win32'], roles: ['cjk-serif', 'latin-serif'] },
  { name: 'Helvetica Neue', aliases: ['HelveticaNeue'], language: 'Western', style: 'Sans-serif', characteristics: 'macOS system sans-serif', layer: 'system', platforms: ['darwin'], roles: ['latin-sans'] },
  { name: 'PingFang SC', aliases: ['PingFangSC'], language: 'Chinese + Western', style: 'Sans-serif', characteristics: 'macOS Simplified Chinese system sans-serif', layer: 'system', platforms: ['darwin'], roles: ['cjk-sans', 'latin-sans'] },
  { name: 'Hiragino Sans GB', aliases: ['HiraginoSansGB'], language: 'Chinese + Western', style: 'Sans-serif', characteristics: 'macOS Simplified Chinese humanist sans-serif', layer: 'system', platforms: ['darwin'], roles: ['cjk-sans', 'latin-sans'] },
  { name: 'Songti SC', aliases: ['SongtiSC'], language: 'Chinese + Western', style: 'Serif', characteristics: 'macOS Simplified Chinese Song serif', layer: 'system', platforms: ['darwin'], roles: ['cjk-serif', 'latin-serif'] },
  { name: 'Liberation Sans', aliases: ['LiberationSans'], language: 'Western', style: 'Sans-serif', characteristics: 'Linux metric-compatible Arial alternative', layer: 'system', platforms: ['linux'], roles: ['latin-sans'] },
  { name: 'Liberation Serif', aliases: ['LiberationSerif'], language: 'Western', style: 'Serif', characteristics: 'Linux metric-compatible Times alternative', layer: 'system', platforms: ['linux'], roles: ['latin-serif'] },
  { name: 'DejaVu Sans', aliases: ['DejaVuSans'], language: 'Multi-language', style: 'Sans-serif', characteristics: 'Widely available Linux sans-serif with broad glyph coverage', layer: 'system', platforms: ['linux'], roles: ['latin-sans', 'code'] },
  { name: 'DejaVu Serif', aliases: ['DejaVuSerif'], language: 'Multi-language', style: 'Serif', characteristics: 'Widely available Linux serif', layer: 'system', platforms: ['linux'], roles: ['latin-serif'] },
  { name: 'Noto Sans CJK SC', aliases: ['NotoSansCJKsc'], language: 'Chinese + Multi-language', style: 'Sans-serif', characteristics: 'Linux-oriented open CJK sans-serif', layer: 'system', platforms: ['linux'], roles: ['cjk-sans', 'latin-sans'] },
  { name: 'Noto Serif CJK SC', aliases: ['NotoSerifCJKsc'], language: 'Chinese + Multi-language', style: 'Serif', characteristics: 'Linux-oriented open CJK serif', layer: 'system', platforms: ['linux'], roles: ['cjk-serif', 'latin-serif'] },
  { name: 'Liter', language: 'English', style: 'Sans-serif', characteristics: 'Modern geometric, low contrast, balanced and rational', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-sans', 'display'] },
  { name: 'HedvigLettersSans', aliases: ['Hedvig Letters Sans'], language: 'English', style: 'Sans-serif', characteristics: 'Slightly irregular with a distinctive brand character', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-sans', 'display'] },
  { name: 'Oranienbaum', language: 'English', style: 'High-contrast serif', characteristics: 'Geometric, elegant and classical', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-serif', 'display'] },
  { name: 'QuattrocentoSans', aliases: ['Quattrocento Sans'], language: 'English', style: 'Classical sans-serif', characteristics: 'Gentle, readable and sharp at small sizes', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-sans'] },
  { name: 'SortsMillGoudy', aliases: ['Sorts Mill Goudy'], language: 'English', style: 'Serif', characteristics: 'Goudy Old Style revival with soft, legible serifs', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-serif'] },
  { name: 'Unna', language: 'English', style: 'Neoclassical serif', characteristics: 'Pronounced vertical rhythm and elegant power', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-serif', 'display'] },
  { name: 'Coda', language: 'English', style: 'Sans-serif', characteristics: 'Round, friendly and open', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['latin-sans', 'display'] },
  { name: 'Jersey15', aliases: ['Jersey 15'], language: 'English + Numbers', style: 'Pixel', characteristics: 'Sports jersey geometry with a strong grid', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['display'] },
  { name: 'Jersey20Charted', aliases: ['Jersey 20 Charted'], language: 'English + Numbers', style: 'Pixel', characteristics: 'Grid-textured sports number style', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['display'] },
  { name: 'MiSans', aliases: ['Mi Sans'], language: 'Chinese + Multi-language', style: 'Sans-serif', characteristics: 'Clean modern variable system font', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-sans', 'latin-sans', 'display'] },
  { name: 'Noto Sans SC', aliases: ['NotoSansSC'], language: 'Chinese + Multi-language', style: 'Sans-serif', characteristics: 'Neutral standardized Source Han Sans structure', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-sans', 'latin-sans'] },
  { name: 'siyuanSongti', aliases: ['Source Han Serif SC', 'Source Han Serif CN'], language: 'Chinese + Multi-language', style: 'Serif', characteristics: 'Refined Song structure with contrasting strokes', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-serif', 'latin-serif'] },
  { name: 'alimamadaoliti', language: 'Chinese', style: 'Clerical', characteristics: 'Knife-edge strokes with power and antiquity', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-serif', 'display'] },
  { name: 'alimamashuheiti', language: 'Chinese', style: 'Geometric sans-serif', characteristics: 'Orderly commercial geometry', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-sans', 'display'] },
  { name: 'zhankuwenyiti', language: 'Chinese', style: 'Handwritten', characteristics: 'Simple, fresh and lightly artistic', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-sans', 'display'] },
  { name: 'feibozhengdianti', language: 'Chinese', style: 'Brush', characteristics: 'Thick and powerful brush strokes', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-serif', 'display'] },
  { name: 'deyihei', language: 'Chinese', style: 'Sans-serif', characteristics: 'Thin slanted humanist geometry', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-sans', 'display'] },
  { name: 'jingpindianzhenTi', language: 'Chinese + Western', style: 'Pixel', characteristics: '9x9 retro-electronic bitmap style', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['display', 'code'] },
  { name: 'LXGW Bright', language: 'Chinese + Western', style: 'Song/Kai', characteristics: 'Gentle, clear and legible', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-serif', 'latin-serif', 'display'] },
  { name: 'ZCOOL KuaiLe', language: 'Chinese + Western', style: 'Display', characteristics: 'Lively, playful and youthful', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['display'] },
  { name: 'xiawuxinzhisong', language: 'Chinese', style: 'Serif', characteristics: 'Bright and elegant Mincho-derived structure', layer: 'custom', platforms: ALL_PLATFORMS, roles: ['cjk-serif', 'display'] },
])

const REGISTERED_BY_KEY = new Map<string, FontDescriptor>()
for (const font of FONT_REGISTRY) {
  for (const name of [font.name, ...(font.aliases ?? [])]) REGISTERED_BY_KEY.set(fontKey(name), font)
}
const REGISTERED_ALIASES = [...REGISTERED_BY_KEY.entries()].sort((a, b) => b[0].length - a[0].length)

export interface DiscoveredFont {
  name: string
  file: string
  sha256: string
  familyName: string
  postscriptName: string | null
  weight: string
  glyphCount: number
  supportsLatin: boolean
  supportsCjk: boolean
  codePoints: ReadonlySet<number>
}

interface FontFaceLike {
  familyName?: unknown
  postscriptName?: unknown
  subfamilyName?: unknown
  characterSet?: unknown
}

interface FontCollectionLike {
  fonts?: FontFaceLike[]
}

function fontKey(value: string): string {
  return value.toLocaleLowerCase().replace(/[\s_-]+/g, '')
}

function registeredFace(value: string): FontDescriptor | undefined {
  const key = fontKey(value)
  const exact = REGISTERED_BY_KEY.get(key)
  if (exact !== undefined) return exact
  for (const [alias, font] of REGISTERED_ALIASES) {
    if (!key.startsWith(alias)) continue
    const suffix = key.slice(alias.length)
    if (/^(?:(?:extra|ultra|semi|demi)?(?:light|bold)|thin|regular|book|medium|heavy|black|italic|oblique|w\d+)+$/u.test(suffix)) return font
  }
  return undefined
}

async function fontFiles(roots: readonly string[], maxFiles = 10_000): Promise<string[]> {
  const files: string[] = []
  const queue = [...new Set(roots)]
  while (queue.length > 0 && files.length < maxFiles) {
    const directory = queue.shift()!
    let handle
    try {
      handle = await opendir(directory)
    } catch {
      continue
    }
    for await (const entry of handle) {
      const path = join(directory, entry.name)
      if (entry.isDirectory()) queue.push(path)
      else if (entry.isFile() && /^\.(?:ttf|otf|ttc)$/i.test(extname(entry.name))) files.push(path)
      if (files.length >= maxFiles) break
    }
  }
  return files.sort()
}

function faces(value: FontFaceLike | FontCollectionLike): FontFaceLike[] {
  return Array.isArray((value as FontCollectionLike).fonts) ? (value as FontCollectionLike).fonts! : [value as FontFaceLike]
}

export async function discoverRegisteredFonts(
  extraDirs: readonly string[] = [],
  platform: SupportedFontPlatform = process.platform as SupportedFontPlatform,
): Promise<DiscoveredFont[]> {
  const discovered: DiscoveredFont[] = []
  for (const file of await fontFiles([...systemFontDirectories(platform), ...extraDirs])) {
    let opened: FontFaceLike | FontCollectionLike
    try {
      opened = await fontkit.open(file) as FontFaceLike | FontCollectionLike
    } catch {
      continue
    }
    let hash: string | undefined
    for (const face of faces(opened)) {
      const familyName = typeof face.familyName === 'string' ? face.familyName : ''
      const postscriptName = typeof face.postscriptName === 'string' ? face.postscriptName : null
      const registered = registeredFace(familyName)
        ?? (postscriptName === null ? undefined : registeredFace(postscriptName))
      if (registered === undefined) continue
      hash ??= createHash('sha256').update(await readFile(file)).digest('hex')
      const points = new Set(Array.isArray(face.characterSet) ? face.characterSet.filter((value): value is number => Number.isInteger(value)) : [])
      discovered.push({
        name: registered.name,
        file,
        sha256: hash,
        familyName,
        postscriptName,
        weight: typeof face.subfamilyName === 'string' ? face.subfamilyName : 'Regular',
        glyphCount: points.size,
        supportsLatin: [...'AaZz09'].every(character => points.has(character.codePointAt(0)!)),
        supportsCjk: points.has('中'.codePointAt(0)!) && points.has('文'.codePointAt(0)!),
        codePoints: points,
      })
    }
  }
  return discovered.sort((a, b) => a.name.localeCompare(b.name) || a.weight.localeCompare(b.weight) || a.file.localeCompare(b.file))
}

export const FONT_FALLBACKS: Readonly<Record<string, readonly string[]>> = Object.freeze({
  Liter: ['HedvigLettersSans', 'QuattrocentoSans', 'Arial'],
  HedvigLettersSans: ['Liter', 'QuattrocentoSans', 'Arial'],
  MiSans: ['alimamashuheiti', 'Noto Sans SC'],
  'Noto Sans SC': ['MiSans', 'alimamashuheiti'],
  siyuanSongti: ['xiawuxinzhisong', 'LXGW Bright'],
  xiawuxinzhisong: ['siyuanSongti', 'LXGW Bright'],
})

const PLATFORM_ROLE_FALLBACKS: Readonly<Record<SupportedFontPlatform, Readonly<Record<FontRole, readonly string[]>>>> = Object.freeze({
  darwin: Object.freeze({
    'latin-sans': ['Arial', 'Helvetica Neue', 'PingFang SC', 'Liter', 'Noto Sans SC'],
    'latin-serif': ['Times New Roman', 'Songti SC', 'siyuanSongti'],
    'cjk-sans': ['PingFang SC', 'Hiragino Sans GB', 'Noto Sans SC', 'MiSans'],
    'cjk-serif': ['Songti SC', 'siyuanSongti', 'LXGW Bright'],
    display: ['Liter', 'PingFang SC', 'Arial', 'Noto Sans SC'],
    code: ['Arial', 'Helvetica Neue', 'PingFang SC', 'Noto Sans SC'],
  }),
  win32: Object.freeze({
    'latin-sans': ['Arial', 'Segoe UI', 'Microsoft YaHei', 'Liter', 'Noto Sans SC'],
    'latin-serif': ['Times New Roman', 'SimSun', 'siyuanSongti'],
    'cjk-sans': ['Microsoft YaHei', 'DengXian', 'Noto Sans SC', 'MiSans'],
    'cjk-serif': ['SimSun', 'siyuanSongti', 'LXGW Bright'],
    display: ['Liter', 'Microsoft YaHei', 'Arial', 'Noto Sans SC'],
    code: ['Arial', 'Segoe UI', 'Microsoft YaHei', 'Noto Sans SC'],
  }),
  linux: Object.freeze({
    'latin-sans': ['Liberation Sans', 'DejaVu Sans', 'Arial', 'Liter', 'Noto Sans SC'],
    'latin-serif': ['Liberation Serif', 'DejaVu Serif', 'Times New Roman', 'Noto Serif CJK SC'],
    'cjk-sans': ['Noto Sans CJK SC', 'Noto Sans SC', 'MiSans'],
    'cjk-serif': ['Noto Serif CJK SC', 'siyuanSongti', 'LXGW Bright'],
    display: ['Liberation Sans', 'DejaVu Sans', 'Noto Sans CJK SC', 'Noto Sans SC'],
    code: ['DejaVu Sans', 'Liberation Sans', 'Noto Sans CJK SC'],
  }),
})

function supportedPlatform(platform: NodeJS.Platform): SupportedFontPlatform {
  return platform === 'win32' || platform === 'linux' ? platform : 'darwin'
}

function containsCjk(text: string): boolean {
  return /[\u3400-\u9FFF\uF900-\uFAFF]/u.test(text)
}

function semanticRole(font: FontDescriptor, text: string): FontRole {
  const serif = font.roles.includes('latin-serif') || font.roles.includes('cjk-serif')
  if (containsCjk(text)) return serif ? 'cjk-serif' : 'cjk-sans'
  if (font.roles.includes('code')) return 'code'
  if (font.roles.includes('display') && !font.roles.includes('latin-sans') && !font.roles.includes('latin-serif')) return 'display'
  return serif ? 'latin-serif' : 'latin-sans'
}

/**
 * Adapts the machine-wide font inventory onto the shape the outline resolver
 * already understands, so a deck may name any font that is actually installed
 * instead of only the families in the built-in registry. Coverage is rebuilt
 * from the face flags: a code point enters the set only when the face claims the
 * script covering it, which keeps `supportsText` meaningful without re-reading
 * every glyph table of every installed font. Pass `wanted` to restrict the
 * conversion to the families a deck actually names.
 */
export function installedFontsAsDiscovered(
  faces: readonly InstalledFontFace[],
  wanted?: ReadonlySet<string>,
): DiscoveredFont[] {
  const discovered: DiscoveredFont[] = []
  for (const face of faces) {
    if (wanted !== undefined && !wanted.has(fontKey(face.family))) continue
    const codePoints = new Set<number>()
    if (face.supportsLatin) for (let point = 0x20; point <= 0x7e; point += 1) codePoints.add(point)
    if (face.supportsCjk) for (let point = 0x4e00; point <= 0x9fa5; point += 1) codePoints.add(point)
    discovered.push({
      name: face.family,
      file: face.file,
      sha256: face.sha256,
      familyName: face.family,
      postscriptName: face.postscriptName,
      weight: face.subfamily,
      glyphCount: face.glyphCount,
      supportsLatin: face.supportsLatin,
      supportsCjk: face.supportsCjk,
      codePoints,
    })
  }
  return discovered
}

export function registeredFont(name: string): FontDescriptor | undefined {
  return REGISTERED_BY_KEY.get(fontKey(name))
}

export function fontFallbackCandidates(
  name: string,
  text: string,
  platform: NodeJS.Platform = process.platform,
): readonly string[] {
  const requested = registeredFont(name)
  if (requested === undefined) throw new PptError('PPT_OUTLINE_INVALID', `font is not registered: ${name}`)
  const role = semanticRole(requested, text)
  return [...new Set([requested.name, ...(FONT_FALLBACKS[requested.name] ?? []), ...PLATFORM_ROLE_FALLBACKS[supportedPlatform(platform)][role]])]
}

export function supportsText(font: DiscoveredFont, text: string): boolean {
  for (const character of text) {
    const point = character.codePointAt(0)!
    if (!/\s/u.test(character) && !font.codePoints.has(point)) return false
  }
  return true
}

export interface ResolvedFont {
  requested: string
  resolved: DiscoveredFont
  fallback: boolean
  warning?: string
}

/**
 * A font name this build does not know about — for example a deck authored
 * against a newer registry than the one the running host has loaded — must not
 * abort the whole pipeline. Fall back in registry order so the outcome stays
 * deterministic, and leave a warning the caller can surface.
 */
function resolveUnregisteredFont(
  name: string,
  text: string,
  discovered: readonly DiscoveredFont[],
  platform: NodeJS.Platform,
): ResolvedFont {
  const supported = supportedPlatform(platform)
  for (const candidate of FONT_REGISTRY) {
    if (!candidate.platforms.includes(supported)) continue
    const font = discovered.find(item => item.name === candidate.name && supportsText(item, text))
    if (font !== undefined) {
      return {
        requested: name,
        resolved: font,
        fallback: true,
        warning: `font ${name} is not registered in this build; replaced with ${supported} fallback ${candidate.name}`,
      }
    }
  }
  throw new PptError('PPT_DEPENDENCY_MISSING', `no installed approved font covers the requested text for ${name}`, {
    details: { requested: name, platform: supported },
  })
}

export function resolveRegisteredFont(
  name: string,
  text: string,
  discovered: readonly DiscoveredFont[],
  platform: NodeJS.Platform = process.platform,
): ResolvedFont {
  const requested = registeredFont(name)
  if (requested === undefined) return resolveUnregisteredFont(name, text, discovered, platform)
  const candidates = fontFallbackCandidates(requested.name, text, platform)
  for (const candidate of candidates) {
    const font = discovered.find(item => item.name === candidate && supportsText(item, text))
    if (font !== undefined) {
      return {
        requested: name,
        resolved: font,
        fallback: candidate !== requested.name,
        ...(candidate === requested.name ? {} : { warning: `font ${requested.name} was replaced with installed ${supportedPlatform(platform)} fallback ${candidate}` }),
      }
    }
  }
  throw new PptError('PPT_DEPENDENCY_MISSING', `no installed approved font covers the requested text for ${requested.name}`, {
    details: { requested: requested.name, platform: supportedPlatform(platform), candidates },
  })
}

export interface FontAvailabilitySummary {
  scope: 'approved_registry'
  platform: SupportedFontPlatform
  registryFamilies: number
  availableFamilies: number
  availableFaces: number
  layers: Record<FontLayer, { registered: number; available: number; families: string[] }>
  roles: Record<FontRole, { available: boolean; families: string[] }>
}

export function summarizeFontAvailability(
  discovered: readonly DiscoveredFont[],
  platform: NodeJS.Platform = process.platform,
): FontAvailabilitySummary {
  const currentPlatform = supportedPlatform(platform)
  const availableNames = new Set(discovered.map(font => font.name))
  const layers = Object.fromEntries(FONT_LAYERS.map(layer => {
    const registered = FONT_REGISTRY.filter(font => font.layer === layer)
    const families = registered.filter(font => availableNames.has(font.name)).map(font => font.name).sort()
    return [layer, { registered: registered.length, available: families.length, families }]
  })) as FontAvailabilitySummary['layers']
  const roles = Object.fromEntries(FONT_ROLES.map(role => {
    const candidates = new Set([
      ...PLATFORM_ROLE_FALLBACKS[currentPlatform][role],
      ...FONT_REGISTRY.filter(font => font.roles.includes(role)).map(font => font.name),
    ])
    const families = [...candidates].filter(name => availableNames.has(name)).sort()
    return [role, { available: families.length > 0, families }]
  })) as FontAvailabilitySummary['roles']
  return {
    scope: 'approved_registry', platform: currentPlatform, registryFamilies: FONT_REGISTRY.length,
    availableFamilies: availableNames.size, availableFaces: discovered.length, layers, roles,
  }
}

export interface FontCatalogOptions {
  text?: string
  role?: FontRole | 'all'
  layer?: FontLayer | 'all'
  includeUnavailable?: boolean
  platform?: NodeJS.Platform
}

export interface FontCatalogEntry {
  name: string
  layer: FontLayer
  platforms: SupportedFontPlatform[]
  roles: FontRole[]
  recommended_for: FontRole[]
  language: string
  style: string
  characteristics: string
  installed: boolean
  weights: string[]
  supports_latin: boolean
  supports_cjk: boolean
  covers_text?: boolean
}

export interface FontCatalog {
  scope: 'approved_registry'
  scope_note: string
  platform: SupportedFontPlatform
  registry_families: number
  available_families: number
  available_faces: number
  returned_families: number
  filters: {
    role: FontRole | 'all'
    layer: FontLayer | 'all'
    include_unavailable: boolean
    text?: string
  }
  recommendations: Record<FontRole, string[]>
  fonts: FontCatalogEntry[]
  warnings: string[]
}

export function fontRecommendations(
  discovered: readonly DiscoveredFont[],
  platform: NodeJS.Platform = process.platform,
  text?: string,
): Record<FontRole, string[]> {
  const currentPlatform = supportedPlatform(platform)
  const available = new Set(discovered.map(font => font.name))
  return Object.fromEntries(FONT_ROLES.map(role => {
    const candidates = [...new Set([
      ...PLATFORM_ROLE_FALLBACKS[currentPlatform][role],
      ...FONT_REGISTRY.filter(font => font.roles.includes(role)).map(font => font.name),
    ])]
    const matching = candidates.filter(name => available.has(name) && (
      text === undefined || discovered.some(font => font.name === name && supportsText(font, text))
    ))
    return [role, matching]
  })) as Record<FontRole, string[]>
}

export function buildFontCatalog(
  discovered: readonly DiscoveredFont[],
  options: FontCatalogOptions = {},
): FontCatalog {
  const platform = supportedPlatform(options.platform ?? process.platform)
  const role = options.role ?? 'all'
  const layer = options.layer ?? 'all'
  const includeUnavailable = options.includeUnavailable ?? false
  const text = options.text
  const availability = summarizeFontAvailability(discovered, platform)
  const recommendations = fontRecommendations(discovered, platform, text)
  const entries = FONT_REGISTRY.map((descriptor): FontCatalogEntry => {
    const faces = discovered.filter(font => font.name === descriptor.name)
    const installed = faces.length > 0
    const coversText = text === undefined ? undefined : faces.some(font => supportsText(font, text))
    const recommendedFor = FONT_ROLES.filter(candidate => recommendations[candidate].includes(descriptor.name))
    return {
      name: descriptor.name, layer: descriptor.layer, platforms: [...descriptor.platforms], roles: [...descriptor.roles],
      recommended_for: recommendedFor, language: descriptor.language, style: descriptor.style,
      characteristics: descriptor.characteristics, installed,
      weights: [...new Set(faces.map(font => font.weight))].sort(),
      supports_latin: faces.some(font => font.supportsLatin), supports_cjk: faces.some(font => font.supportsCjk),
      ...(coversText === undefined ? {} : { covers_text: coversText }),
    }
  }).filter(font => {
    if (!includeUnavailable && !font.installed) return false
    if (!includeUnavailable && text !== undefined && font.covers_text !== true) return false
    if (layer !== 'all' && font.layer !== layer) return false
    if (role !== 'all' && !font.roles.includes(role) && !font.recommended_for.includes(role)) return false
    return true
  })
  const warnings: string[] = []
  if (entries.length === 0) warnings.push('No approved font matches the requested filters on this host.')
  if (role !== 'all' && recommendations[role].length === 0) warnings.push(`No installed approved ${role} font covers the requested text.`)
  return {
    scope: 'approved_registry',
    scope_note: 'This is the installed subset of the plugin approved registry, not the host-wide font inventory.',
    platform, registry_families: availability.registryFamilies, available_families: availability.availableFamilies,
    available_faces: availability.availableFaces, returned_families: entries.length,
    filters: { role, layer, include_unavailable: includeUnavailable, ...(text === undefined ? {} : { text }) },
    recommendations, fonts: entries, warnings,
  }
}
