import { execFile } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFile, mkdir, opendir, readFile } from 'node:fs/promises'
import { homedir } from 'node:os'
import { basename, extname, join, posix, resolve, win32 } from 'node:path'
import * as fontkit from 'fontkit'
import { PptError } from './errors.ts'
import { systemFontDirectories } from './platform.ts'

/**
 * Physical font inventory, text-subset extraction, and per-user font installation.
 *
 * This module stays on plain Node built-ins plus fontkit and deliberately does
 * NOT take the DSH `subprocess` service: the three installation commands are
 * short-lived, fire-and-forget helpers, and keeping them on `node:child_process`
 * means the whole module is testable as a unit without a Cordis host.
 *
 * Every error below reuses a code that already exists in `src/errors.ts`.
 * There is no dedicated `PPT_FONT_UNAVAILABLE` code in that frozen set, so
 * "the font file cannot be parsed or addressed" is reported as
 * `PPT_DEPENDENCY_MISSING` (the same choice `src/fonts.ts` makes for
 * "no usable font") and `PPT_CAPABILITY_UNAVAILABLE` for a fontkit entry point
 * that is missing at runtime.
 */

/** Mirrors the `maxFiles` budget used by the registry scanner in `src/fonts.ts`. */
export const MAX_FONT_FILES = 10_000

export const FONT_FILE_EXTENSIONS: readonly string[] = Object.freeze(['.ttf', '.otf', '.ttc', '.otc'])

const FONT_FILE_PATTERN = /\.(?:ttf|otf|ttc|otc)$/iu

const SFNT_TRUE_TYPE = 0x00010000
const SFNT_APPLE_TRUE_TYPE = 0x74727565 // 'true'
const SFNT_CFF = 0x4f54544f // 'OTTO'
const SFNT_COLLECTION = 0x74746366 // 'ttcf'
const SFNT_HEADER_BYTES = 12
const SFNT_RECORD_BYTES = 16

/** OS/2 header layout: version, xAvgCharWidth, usWeightClass, usWidthClass, fsType, ... */
const OS2_WEIGHT_CLASS_OFFSET = 4
const OS2_FS_TYPE_OFFSET = 8
/** version + xAvgCharWidth + usWeightClass + usWidthClass + fsType + 8 int16 metrics + sFamilyClass. */
const OS2_PANOSE_OFFSET = 32
const OS2_PANOSE_BYTES = 10

/** PANOSE-1 `proportion` byte value for a monospaced design. */
const PANOSE_MONOSPACED_PROPORTION = 9
/** `post` table layout: version, italicAngle, underlinePosition, underlineThickness, isFixedPitch. */
const POST_IS_FIXED_PITCH_OFFSET = 12
const POST_MIN_BYTES = 16

const DEFAULT_WEIGHT_CLASS = 400

/** `fsType` bits, per the OpenType OS/2 specification. */
const FS_TYPE_RESTRICTED = 0x0002
const FS_TYPE_PREVIEW_PRINT = 0x0004
const FS_TYPE_EDITABLE = 0x0008
const FS_TYPE_BITMAP_ONLY = 0x0200
const FS_TYPE_EMBEDDING_MASK = 0x000e

const LATIN_FIRST = 0x0020
const LATIN_LAST = 0x007e
const CJK_FIRST = 0x4e00
const CJK_LAST = 0x9fa5
/** A face counts as CJK-capable once it covers this many of the common Han range. */
const MIN_CJK_CODEPOINTS = 200

/**
 * `lfCharSet` values, as signed bytes: GB2312_CHARSET (134), SHIFTJIS_CHARSET
 * (128), HANGUL_CHARSET (129). They are reported as hints only; a face that
 * cannot be attributed to one writing system gets ANSI_CHARSET (0).
 */
const CHARSET_ANSI = 0
const CHARSET_GB2312 = -122
const CHARSET_SHIFT_JIS = -128
const CHARSET_HANGUL = -127

const PITCH_FIXED = 1
const PITCH_VARIABLE = 2

const FAMILY_DONT_CARE = 0
const FAMILY_ROMAN = 1
const FAMILY_SWISS = 2
const FAMILY_SCRIPT = 4
const FAMILY_DECORATIVE = 5

const WINDOWS_FONTS_KEY = 'HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts'
const COMMAND_TIMEOUT_MS = 15_000

/**
 * Family names that identify the writing system a face was designed for. The
 * comparison key drops case, spaces, underscores, and hyphens, so the macOS and
 * Windows spellings of the same family collapse onto one entry.
 */
const CHINESE_FAMILY_ROOTS: readonly string[] = Object.freeze([
  'simsun', 'nsimsun', 'simhei', 'microsoftyahei', 'dengxian', 'fangsong', 'kaiti', 'microsoftjianhei',
])
const JAPANESE_FAMILY_ROOTS: readonly string[] = Object.freeze([
  'msgothic', 'msmincho', 'mspgothic', 'mspmincho', 'msuigothic', 'yugothic', 'yumincho', 'meiryo',
])
const KOREAN_FAMILY_ROOTS: readonly string[] = Object.freeze([
  'malgungothic', 'batang', 'gulim', 'dotum', 'gungsuh', 'nanumgothic',
])

export interface InstalledFontFace {
  family: string
  subfamily: string
  postscriptName: string | null
  fullName: string
  /** Absolute path of the font file on disk. */
  file: string
  /** For .ttc collections this is the index of the face inside the file, otherwise 0. */
  faceIndex: number
  format: 'truetype' | 'cff' | 'collection'
  weightClass: number
  isFixedPitch: boolean
  fsType: number
  /** True when the OS/2 fsType permission bits allow embedding. */
  embeddable: boolean
  /** 10 OS/2 PANOSE bytes as 20 uppercase hex characters, or null when the face has no OS/2 table. */
  panose: string | null
  /** Windows LOGFONT-style pitch/family byte derived from PANOSE plus the fixed-pitch flag. */
  pitchFamily: number
  /**
   * Suggested `lfCharSet` hint: GB2312 (-122), Shift-JIS (-128), Hangul (-127),
   * or ANSI (0) when the writing system cannot be established from the family
   * name. See `CHINESE_FAMILY_ROOTS` and `charsetOf` for the rule.
   */
  charset: number
  glyphCount: number
  supportsLatin: boolean
  supportsCjk: boolean
  sha256: string
}

interface SfntTableRecord {
  offset: number
  length: number
}

interface SfntDirectory {
  version: number
  tables: ReadonlyMap<string, SfntTableRecord>
}

interface FontkitGlyphLike {
  id?: unknown
}

interface FontkitFaceLike {
  familyName?: unknown
  subfamilyName?: unknown
  fullName?: unknown
  postscriptName?: unknown
  numGlyphs?: unknown
  characterSet?: unknown
  hasGlyphForCodePoint?: unknown
  glyphForCodePoint?: unknown
  createSubset?: unknown
}

interface FontkitOpenLike extends FontkitFaceLike {
  fonts?: FontkitFaceLike[]
}

interface FontkitSubsetLike {
  includeGlyph(glyph: number): number
  encode(): Uint8Array
  glyphs?: unknown
}

/**
 * Maps a PANOSE-1 `familyType` byte onto the Windows LOGFONT family nibble, then
 * packs it above the pitch nibble: `(family << 4) | pitch`, where pitch is
 * FF_FIXED_PITCH (1) when the face is monospaced and FF_VARIABLE_PITCH (2)
 * otherwise. Only `isFixedPitch` decides the pitch nibble, matching the PANOSE
 * mapping documented for `pitchFamily`. Values outside 0..11 are treated as
 * don't care (0) rather than guessed, so a face with a corrupt PANOSE header
 * never claims a serif/sans/script family it may not have.
 */
export function pitchFamilyFromPanose(panose: readonly number[] | null, isFixedPitch: boolean): number {
  const rawFamilyType = panose === null ? FAMILY_DONT_CARE : panose[0] ?? FAMILY_DONT_CARE
  return (panoseFamilyNibble(rawFamilyType) << 4) | (isFixedPitch ? PITCH_FIXED : PITCH_VARIABLE)
}

function panoseFamilyNibble(familyType: number): number {
  if (!Number.isInteger(familyType) || familyType < 0 || familyType > 11) return FAMILY_DONT_CARE
  if (familyType === 0) return FAMILY_DONT_CARE
  if (familyType === 1) return FAMILY_ROMAN
  if (familyType <= 8) return FAMILY_SWISS
  if (familyType <= 10) return FAMILY_SCRIPT
  return FAMILY_DECORATIVE
}

/**
 * Normalises a family name for comparison. This is a deliberate copy of the
 * private `fontKey` helper in `src/fonts.ts`: that function is not exported, so
 * the same rule (lowercase, drop spaces/underscores/hyphens) is restated here to
 * keep the two call sites byte-for-byte compatible.
 */
function fontKey(value: string): string {
  return value.toLocaleLowerCase().replace(/[\s_-]+/gu, '')
}

function nonEmptyString(value: unknown): string | null {
  return typeof value === 'string' && value.length > 0 ? value : null
}

function uint16(buffer: Uint8Array, offset: number): number {
  return (buffer[offset]! << 8) | buffer[offset + 1]!
}

function uint32(buffer: Uint8Array, offset: number): number {
  return ((buffer[offset]! << 24) | (buffer[offset + 1]! << 16) | (buffer[offset + 2]! << 8) | buffer[offset + 3]!) >>> 0
}

function sfntTag(buffer: Uint8Array, offset: number): string {
  return String.fromCharCode(buffer[offset]!, buffer[offset + 1]!, buffer[offset + 2]!, buffer[offset + 3]!)
}

/**
 * Reads one sfnt table directory. Offsets inside a collection are absolute file
 * offsets, so the same reader works for a bare .ttf/.otf (offset 0) and for each
 * face of a .ttc/.otc.
 */
function sfntDirectory(buffer: Uint8Array, offset: number): SfntDirectory | undefined {
  if (offset < 0 || offset + SFNT_HEADER_BYTES > buffer.length) return undefined
  const version = uint32(buffer, offset)
  const tableCount = uint16(buffer, offset + 4)
  const tables = new Map<string, SfntTableRecord>()
  for (let index = 0; index < tableCount; index++) {
    const record = offset + SFNT_HEADER_BYTES + index * SFNT_RECORD_BYTES
    if (record + SFNT_RECORD_BYTES > buffer.length) break
    tables.set(sfntTag(buffer, record), { offset: uint32(buffer, record + 8), length: uint32(buffer, record + 12) })
  }
  return { version, tables }
}

function collectionFaceOffsets(buffer: Uint8Array): number[] {
  if (buffer.length < SFNT_HEADER_BYTES || uint32(buffer, 0) !== SFNT_COLLECTION) return []
  const count = uint32(buffer, 8)
  const offsets: number[] = []
  for (let index = 0; index < count; index++) {
    const at = SFNT_HEADER_BYTES + index * 4
    if (at + 4 > buffer.length) break
    offsets.push(uint32(buffer, at))
  }
  return offsets
}

/** Table directories of every face in the file, in on-disk face order. */
function faceDirectories(buffer: Uint8Array): SfntDirectory[] {
  const offsets = collectionFaceOffsets(buffer)
  const directories = (offsets.length > 0 ? offsets : [0])
    .map(offset => sfntDirectory(buffer, offset))
    .filter((directory): directory is SfntDirectory => directory !== undefined)
  return directories
}

/**
 * PANOSE bytes are read straight out of the file at the fixed OS/2 offset
 * instead of through fontkit. fontkit 2.0.4 exposes tables only as minified
 * prototype getters (`font['OS/2']`, `font.post`, ...) and its bundled type
 * declaration in `src/types/fontkit.d.ts` declares none of them, so the byte
 * path is both more reliable across fontkit builds and type-safe without
 * touching a file this module does not own.
 */
function panoseBytes(buffer: Uint8Array, directory: SfntDirectory | undefined): number[] | null {
  const table = directory?.tables.get('OS/2')
  if (table === undefined) return null
  const end = table.offset + OS2_PANOSE_OFFSET + OS2_PANOSE_BYTES
  if (table.offset + OS2_PANOSE_OFFSET < 0 || end > buffer.length) return null
  return [...buffer.subarray(table.offset + OS2_PANOSE_OFFSET, end)]
}

function panoseHex(panose: readonly number[] | null): string | null {
  return panose === null ? null : panose.map(byte => byte.toString(16).padStart(2, '0').toUpperCase()).join('')
}

function fsTypeOf(buffer: Uint8Array, directory: SfntDirectory | undefined): number {
  const table = directory?.tables.get('OS/2')
  if (table === undefined || table.offset + OS2_FS_TYPE_OFFSET + 2 > buffer.length) return 0
  return uint16(buffer, table.offset + OS2_FS_TYPE_OFFSET)
}

/**
 * Installable embedding is `fsType` bits 1..3 all clear; the restricted-license
 * bit (0x0002) forbids document embedding outright, while preview/print
 * (0x0004) and editable (0x0008) both permit it. Bitmap-only embedding (0x0200)
 * forbids embedding the outline glyphs, which is what this flag reports on.
 */
function isEmbeddable(fsType: number): boolean {
  if ((fsType & FS_TYPE_BITMAP_ONLY) !== 0) return false
  if ((fsType & FS_TYPE_EMBEDDING_MASK) === 0) return true
  return (fsType & FS_TYPE_RESTRICTED) === 0 && (fsType & (FS_TYPE_PREVIEW_PRINT | FS_TYPE_EDITABLE)) !== 0
}

function isFixedPitchOf(buffer: Uint8Array, directory: SfntDirectory | undefined, panose: readonly number[] | null): boolean {
  const post = directory?.tables.get('post')
  const postFixed = post !== undefined && post.length >= POST_MIN_BYTES && uint32(buffer, post.offset + POST_IS_FIXED_PITCH_OFFSET) !== 0
  return postFixed || panose?.[3] === PANOSE_MONOSPACED_PROPORTION
}

function weightClassOf(buffer: Uint8Array, directory: SfntDirectory | undefined): number {
  const table = directory?.tables.get('OS/2')
  if (table === undefined || table.offset + OS2_WEIGHT_CLASS_OFFSET + 2 > buffer.length) return DEFAULT_WEIGHT_CLASS
  const value = uint16(buffer, table.offset + OS2_WEIGHT_CLASS_OFFSET)
  return value === 0 || value > 1000 ? DEFAULT_WEIGHT_CLASS : value
}

function formatOf(buffer: Uint8Array, directory: SfntDirectory | undefined): InstalledFontFace['format'] {
  if (collectionFaceOffsets(buffer).length > 0) return 'collection'
  return directory?.version === SFNT_CFF ? 'cff' : 'truetype'
}

/**
 * Glyph coverage probe. `hasGlyphForCodePoint` is preferred because it resolves
 * through the cmap without materialising the full character set, which matters
 * for CJK faces carrying 30k+ code points; the documented `characterSet` array
 * is the fallback for any fontkit face that does not expose the method.
 */
function coverageProbe(face: FontkitFaceLike): (codePoint: number) => boolean {
  const hasGlyph = face.hasGlyphForCodePoint
  if (typeof hasGlyph === 'function') {
    const probe = hasGlyph as (this: FontkitFaceLike, codePoint: number) => unknown
    return codePoint => probe.call(face, codePoint) === true
  }
  const points = new Set(Array.isArray(face.characterSet) ? face.characterSet.filter((value): value is number => Number.isInteger(value)) : [])
  return codePoint => points.has(codePoint)
}

function coversRange(probe: (codePoint: number) => boolean, first: number, last: number, minimum: number): boolean {
  let found = 0
  for (let codePoint = first; codePoint <= last; codePoint++) {
    if (!probe(codePoint)) continue
    found++
    if (found >= minimum) return true
  }
  return false
}

function charsetOf(names: readonly (string | null)[], supportsCjk: boolean): number {
  if (!supportsCjk) return CHARSET_ANSI
  const keys = names.filter((name): name is string => name !== null).map(fontKey)
  const hits = (roots: readonly string[]): boolean => keys.some(key => roots.some(root => key.startsWith(root)))
  if (hits(CHINESE_FAMILY_ROOTS)) return CHARSET_GB2312
  if (hits(JAPANESE_FAMILY_ROOTS)) return CHARSET_SHIFT_JIS
  if (hits(KOREAN_FAMILY_ROOTS)) return CHARSET_HANGUL
  return CHARSET_ANSI
}

/** Recursive font-file scan, same shape and budget as the scanner in `src/fonts.ts`. */
async function fontFiles(roots: readonly string[], maxFiles = MAX_FONT_FILES): Promise<string[]> {
  const files: string[] = []
  const queue = [...new Set(roots.map(root => resolve(root)))]
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
      else if (entry.isFile() && FONT_FILE_PATTERN.test(extname(entry.name))) files.push(path)
      if (files.length >= maxFiles) break
    }
  }
  return files.sort()
}

function faceEntry(
  file: string,
  faceIndex: number,
  face: FontkitFaceLike,
  directory: SfntDirectory | undefined,
  buffer: Uint8Array,
  sha256: string,
): InstalledFontFace | undefined {
  const postscriptName = nonEmptyString(face.postscriptName)
  const family = nonEmptyString(face.familyName) ?? postscriptName ?? basename(file)
  const subfamily = nonEmptyString(face.subfamilyName) ?? 'Regular'
  const probe = coverageProbe(face)
  const panose = panoseBytes(buffer, directory)
  const isFixedPitch = isFixedPitchOf(buffer, directory, panose)
  const supportsLatin = coversRange(probe, LATIN_FIRST, LATIN_LAST, LATIN_LAST - LATIN_FIRST + 1)
  const supportsCjk = coversRange(probe, CJK_FIRST, CJK_LAST, MIN_CJK_CODEPOINTS)
  const fsType = fsTypeOf(buffer, directory)
  return {
    family,
    subfamily,
    postscriptName,
    fullName: nonEmptyString(face.fullName) ?? `${family} ${subfamily}`,
    file,
    faceIndex,
    format: formatOf(buffer, directory),
    weightClass: weightClassOf(buffer, directory),
    isFixedPitch,
    fsType,
    embeddable: isEmbeddable(fsType),
    panose: panoseHex(panose),
    pitchFamily: pitchFamilyFromPanose(panose, isFixedPitch),
    charset: charsetOf([family, postscriptName], supportsCjk),
    glyphCount: Number.isInteger(face.numGlyphs) ? face.numGlyphs as number : 0,
    supportsLatin,
    supportsCjk,
    sha256,
  }
}

/**
 * Enumerates every face installed for the requested platform plus `extraDirs`.
 *
 * A file that cannot be read, cannot be parsed as sfnt, or that fontkit refuses
 * is skipped: one damaged file in a font directory must not fail the whole scan.
 * File contents are read exactly once per file, so all faces of a collection
 * share one buffer read and one sha256.
 */
export async function listInstalledFonts(
  extraDirs: readonly string[] = [],
  platform: NodeJS.Platform = process.platform,
): Promise<InstalledFontFace[]> {
  const faces: InstalledFontFace[] = []
  for (const file of await fontFiles([...systemFontDirectories(platform), ...extraDirs])) {
    let buffer: Uint8Array
    try {
      buffer = await readFile(file)
    } catch {
      continue // the file vanished between the directory walk and the read
    }
    const directories = faceDirectories(buffer)
    if (directories.length === 0) continue
    let opened: FontkitOpenLike
    try {
      opened = await fontkit.open(file) as unknown as FontkitOpenLike
    } catch {
      continue
    }
    const declared = Array.isArray(opened.fonts) ? opened.fonts : [opened]
    const sha256 = createHash('sha256').update(buffer).digest('hex')
    declared.forEach((face, faceIndex) => {
      const entry = faceEntry(file, faceIndex, face, directories[faceIndex], buffer, sha256)
      if (entry !== undefined) faces.push(entry)
    })
  }
  return faces.sort((a, b) =>
    a.family.localeCompare(b.family)
    || a.weightClass - b.weightClass
    || a.file.localeCompare(b.file)
    || a.faceIndex - b.faceIndex)
}

export interface FontLookupResult {
  family: string
  matched: InstalledFontFace | undefined
  faces: InstalledFontFace[]
  /** True when the family was found on this machine. */
  installed: boolean
}

function byWeight(a: InstalledFontFace, b: InstalledFontFace): number {
  return a.weightClass - b.weightClass || a.subfamily.localeCompare(b.subfamily) || a.file.localeCompare(b.file) || a.faceIndex - b.faceIndex
}

/**
 * Case-insensitive family lookup. Spaces, underscores, and hyphens are ignored,
 * so "Microsoft YaHei", "microsoftyahei", and "Microsoft_YaHei" all resolve to
 * the same family. When no family matches exactly, the PostScript name is
 * retried, which is how `.ttc` faces whose family name is localised still answer.
 */
export function fontLookup(family: string, faces: readonly InstalledFontFace[]): FontLookupResult {
  const key = fontKey(family)
  const exact = faces.filter(face => fontKey(face.family) === key)
  const matches = exact.length > 0
    ? exact
    : faces.filter(face => face.postscriptName !== null && fontKey(face.postscriptName) === key)
  const sorted = [...matches].sort(byWeight)
  const matched = sorted.find(face => face.weightClass === DEFAULT_WEIGHT_CLASS) ?? sorted[0]
  return { family, matched, faces: sorted, installed: sorted.length > 0 }
}

export interface FontSubsetResult {
  data: Uint8Array
  /** Glyphs in the emitted font, including `.notdef` and any compound components fontkit pulled in. */
  glyphCount: number
  /** Code points from the request that the face cannot render. */
  missingCodePoints: string[]
}

async function openFace(face: InstalledFontFace): Promise<FontkitFaceLike> {
  let opened: FontkitOpenLike
  try {
    opened = await fontkit.open(face.file) as unknown as FontkitOpenLike
  } catch (error) {
    throw new PptError('PPT_DEPENDENCY_MISSING', `font file is not parseable: ${face.file}`, {
      cause: error,
      details: { file: face.file, family: face.family },
    })
  }
  const declared = Array.isArray(opened.fonts) ? opened.fonts : [opened]
  const target = declared[face.faceIndex]
  if (target === undefined) {
    throw new PptError('PPT_DEPENDENCY_MISSING', `font file has no face at index ${face.faceIndex}: ${face.file}`, {
      details: { file: face.file, faceIndex: face.faceIndex, faces: declared.length },
    })
  }
  return target
}

function glyphIdFor(face: FontkitFaceLike, codePoint: number): number {
  const glyphForCodePoint = face.glyphForCodePoint
  if (typeof glyphForCodePoint !== 'function') return 0
  const glyph = (glyphForCodePoint as (this: FontkitFaceLike, codePoint: number) => FontkitGlyphLike | null | undefined).call(face, codePoint)
  const id = glyph?.id
  return typeof id === 'number' && Number.isInteger(id) && id > 0 ? id : 0
}

function formatCodePoint(codePoint: number): string {
  return `U+${codePoint.toString(16).toUpperCase().padStart(4, '0')}`
}

/**
 * Extracts a standalone subset containing `.notdef` and every glyph the text
 * needs.
 *
 * `keepLayoutTables` is accepted but cannot be honoured by fontkit 2.0.4: both
 * `TTFSubset.encode()` and `CFFSubset.encode()` hard-code the table list they
 * write and expose no hook for GSUB/GPOS/GDEF, so the option is documented as a
 * no-op rather than silently failing or throwing. Callers that need layout
 * features must re-attach them from the original file.
 *
 * Two further fontkit behaviours are recorded here because they are observable
 * in the returned bytes: a TrueType subset is a complete sfnt whose version tag
 * is Apple's 'true' (0x74727565), while a CFF/CID subset is a bare CFF table
 * stream with no sfnt wrapper.
 */
export async function subsetFontForText(
  face: InstalledFontFace,
  text: string,
  options: { keepLayoutTables?: boolean } = {},
): Promise<FontSubsetResult> {
  if (typeof text !== 'string' || text.length === 0) {
    throw new PptError('PPT_CREATE_INPUT_INVALID', 'subsetFontForText requires non-empty text', {
      details: { family: face.family, file: face.file },
    })
  }
  const target = await openFace(face)
  const createSubset = target.createSubset
  if (typeof createSubset !== 'function') {
    throw new PptError('PPT_CAPABILITY_UNAVAILABLE', `fontkit cannot subset this face: ${face.file}`, {
      details: { file: face.file, family: face.family, keepLayoutTables: options.keepLayoutTables ?? true },
    })
  }
  const subset = (createSubset as (this: FontkitFaceLike) => FontkitSubsetLike).call(target)
  subset.includeGlyph(0) // `.notdef` first; the base constructor already does this, kept explicit

  const requested = new Set<number>()
  const missing: number[] = []
  for (const character of text) {
    const codePoint = character.codePointAt(0)!
    if (requested.has(codePoint)) continue
    requested.add(codePoint)
    const glyphId = glyphIdFor(target, codePoint)
    if (glyphId === 0) {
      missing.push(codePoint)
      continue
    }
    subset.includeGlyph(glyphId)
  }

  const data = subset.encode()
  if (!(data instanceof Uint8Array)) {
    throw new PptError('PPT_CAPABILITY_UNAVAILABLE', 'fontkit returned a non-binary subset', {
      details: { file: face.file, family: face.family },
    })
  }
  const included = subset.glyphs
  return {
    data,
    glyphCount: Array.isArray(included) ? included.length : requested.size + 1,
    missingCodePoints: missing.sort((a, b) => a - b).map(formatCodePoint),
  }
}

export interface FontInstallResult {
  family: string
  installedPath: string
  platform: NodeJS.Platform
  scope: 'user'
  registered: boolean
  /** Command that undoes the install, or null when a manual removal is required. */
  uninstallHint: string | null
}

/**
 * `platform`, `home`, and `env` are injectable for the same reason every helper
 * in `src/platform.ts` takes them: the per-user targets differ per OS and the
 * unit tests must be able to exercise the macOS/Linux branches on any host
 * without writing into a real home directory. They are optional overrides, so
 * the documented `{ family, dryRun }` call shape is unchanged.
 */
export interface FontInstallOptions {
  family?: string
  dryRun?: boolean
  platform?: NodeJS.Platform
  home?: string
  env?: NodeJS.ProcessEnv
}

interface CommandResult {
  ok: boolean
  missing: boolean
  stderr: string
}

/** `execFile` only — never a shell string — so a font family name can never be interpreted as a command. */
async function runCommand(file: string, args: readonly string[]): Promise<CommandResult> {
  return new Promise<CommandResult>(settle => {
    execFile(file, [...args], { timeout: COMMAND_TIMEOUT_MS, windowsHide: true }, (error, _stdout, stderr) => {
      settle({
        ok: error === null,
        missing: error !== null && (error as NodeJS.ErrnoException).code === 'ENOENT',
        stderr: String(stderr ?? '').trim(),
      })
    })
  })
}

async function sha256OfFile(path: string): Promise<string | undefined> {
  try {
    return createHash('sha256').update(await readFile(path)).digest('hex')
  } catch {
    return undefined
  }
}

function userFontDirectory(platform: NodeJS.Platform, home: string, env: NodeJS.ProcessEnv): string {
  if (platform === 'win32') {
    const local = env.LOCALAPPDATA ?? win32.join(home, 'AppData', 'Local')
    return win32.join(local, 'Microsoft', 'Windows', 'Fonts')
  }
  if (platform === 'darwin') return posix.join(home, 'Library', 'Fonts')
  if (platform === 'linux') return posix.join(home, '.local', 'share', 'fonts')
  throw new PptError('PPT_PLATFORM_UNSUPPORTED', `font installation is not supported on ${platform}`, {
    details: { platform },
  })
}

function installPath(platform: NodeJS.Platform, directory: string, fileName: string): string {
  return platform === 'win32' ? win32.join(directory, fileName) : posix.join(directory, fileName)
}

async function familyOfInstalledSource(source: string): Promise<string> {
  let opened: FontkitOpenLike
  try {
    opened = await fontkit.open(source) as unknown as FontkitOpenLike
  } catch (error) {
    const code = (error as NodeJS.ErrnoException | undefined)?.code
    if (typeof code === 'string') {
      throw new PptError('PPT_PATH_INVALID', `font source is not readable: ${source}`, { cause: error, details: { code } })
    }
    throw new PptError('PPT_DEPENDENCY_MISSING', `font source is not a parseable font: ${source}`, { cause: error })
  }
  const first = (Array.isArray(opened.fonts) ? opened.fonts[0] : opened) ?? {}
  const family = nonEmptyString(first.familyName) ?? nonEmptyString(first.postscriptName)
  if (family === null) {
    throw new PptError('PPT_DEPENDENCY_MISSING', `font source exposes no family name: ${source}`)
  }
  return family
}

function summarizeCommand(command: string, result: CommandResult): string {
  return `${command} exited with an error${result.stderr.length === 0 ? '' : `: ${result.stderr.split(/\r?\n/u)[0]}`}`
}

/**
 * Installs a font for the current user only, with no elevation on any platform:
 * `%LOCALAPPDATA%\Microsoft\Windows\Fonts` plus an `HKCU` registry value on
 * Windows, `~/Library/Fonts` on macOS, `~/.local/share/fonts` plus a best-effort
 * `fc-cache -f` on Linux. Registry writes go through `reg.exe` so no new
 * dependency (winreg, PowerShell COM) is introduced.
 *
 * The call is idempotent: when the destination already holds the same bytes it
 * is left untouched. `dryRun` resolves the family and the paths only.
 */
export async function installFontFile(sourcePath: string, options: FontInstallOptions = {}): Promise<FontInstallResult> {
  const platform = options.platform ?? process.platform
  const home = options.home ?? homedir()
  const env = options.env ?? process.env
  const directory = userFontDirectory(platform, home, env)
  const source = resolve(sourcePath)
  const family = nonEmptyString(options.family) ?? await familyOfInstalledSource(source)
  const fileName = basename(source)
  const destination = installPath(platform, directory, fileName)

  if (options.dryRun === true) {
    return { family, installedPath: destination, platform, scope: 'user', registered: false, uninstallHint: null }
  }

  await mkdir(directory, { recursive: true })
  const existing = await sha256OfFile(destination)
  if (existing === undefined || existing !== await sha256OfFile(source)) {
    try {
      await copyFile(source, destination)
    } catch (error) {
      throw new PptError('PPT_CREATE_WRITE_FAILED', `failed to copy the font to ${destination}`, {
        cause: error,
        details: { source, destination },
      })
    }
  }

  if (platform === 'win32') {
    const value = `${family} (TrueType)`
    const result = await runCommand('reg.exe', ['add', WINDOWS_FONTS_KEY, '/v', value, '/t', 'REG_SZ', '/d', fileName, '/f'])
    if (!result.ok) {
      throw new PptError('PPT_CREATE_WRITE_FAILED', summarizeCommand('reg.exe add', result), {
        details: { key: WINDOWS_FONTS_KEY, value, file: destination, missing: result.missing },
      })
    }
    return {
      family,
      installedPath: destination,
      platform,
      scope: 'user',
      registered: true,
      uninstallHint: `reg delete "${WINDOWS_FONTS_KEY}" /v "${value}" /f && del "${destination}"`,
    }
  }

  if (platform === 'linux') {
    // A missing or failing fc-cache is not an installation failure: the file is
    // already in a scanned directory and the next cache rebuild picks it up.
    await runCommand('fc-cache', ['-f'])
    return {
      family,
      installedPath: destination,
      platform,
      scope: 'user',
      registered: false,
      uninstallHint: `rm "${destination}" && fc-cache -f`,
    }
  }

  // macOS reads ~/Library/Fonts directly; there is no index to update.
  return {
    family,
    installedPath: destination,
    platform,
    scope: 'user',
    registered: false,
    uninstallHint: `rm "${destination}"`,
  }
}
