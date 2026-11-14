import { createHash } from 'node:crypto'
import { access, copyFile, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import {
  fontLookup, installFontFile, listInstalledFonts, pitchFamilyFromPanose, subsetFontForText,
  type InstalledFontFace,
} from '../../src/font-files.ts'
import { systemFontDirectories } from '../../src/platform.ts'

const HEX_64 = /^[0-9a-f]{64}$/u
const PANOSE_HEX = /^[0-9A-F]{20}$/u

/** Slash-normalised comparison, because the simulated macOS/Linux branches build posix paths on any host. */
function slashes(value: string): string {
  return value.replace(/\\/gu, '/')
}

/**
 * Real system font files are copied into a synthetic, writable directory so the
 * scanner is exercised without touching the host font directories. On this
 * Windows host that is `C:\Windows\Fonts`; on any other host the same lookup
 * runs against that platform's first system font directory.
 */
const systemFontDir = systemFontDirectories(process.platform)[0]!
const systemFontNames: string[] = await readdir(systemFontDir).catch(() => [])

function pickFontFile(pattern: RegExp, preferred?: string): string | undefined {
  const chosen = preferred !== undefined && systemFontNames.includes(preferred)
    ? preferred
    : systemFontNames.find(name => pattern.test(name))
  return chosen === undefined ? undefined : join(systemFontDir, chosen)
}

const trueTypeSource = pickFontFile(/\.ttf$/iu, 'consola.ttf')
const collectionSource = pickFontFile(/\.ttc$/iu, 'simsun.ttc')

const fonts = await mkdtemp(join(tmpdir(), 'dsh-ppt-font-files-'))
const installHome = await mkdtemp(join(tmpdir(), 'dsh-ppt-font-home-'))

if (trueTypeSource !== undefined) await copyFile(trueTypeSource, join(fonts, basename(trueTypeSource)))
if (collectionSource !== undefined) await copyFile(collectionSource, join(fonts, basename(collectionSource)))
await writeFile(join(fonts, 'broken.ttf'), Buffer.from('not a font'))

/** A synthetic face, so lookup assertions never depend on the host inventory. */
function faceOf(overrides: Partial<InstalledFontFace> & { family: string }): InstalledFontFace {
  return {
    subfamily: 'Regular',
    postscriptName: null,
    fullName: overrides.family,
    file: `/fonts/${overrides.family}.ttf`,
    faceIndex: 0,
    format: 'truetype',
    weightClass: 400,
    isFixedPitch: false,
    fsType: 0,
    embeddable: true,
    panose: null,
    pitchFamily: 2,
    charset: 0,
    glyphCount: 0,
    supportsLatin: true,
    supportsCjk: false,
    sha256: 'a'.repeat(64),
    ...overrides,
  }
}

async function sha256Of(path: string): Promise<string> {
  return createHash('sha256').update(await readFile(path)).digest('hex')
}

/**
 * Scans the synthetic directory only: a platform whose font directories do not
 * exist on this host keeps the host inventory out of the result.
 */
async function scan(): Promise<InstalledFontFace[]> {
  return listInstalledFonts([fonts], 'linux')
}

afterAll(async () => {
  await rm(fonts, { recursive: true, force: true })
  await rm(installHome, { recursive: true, force: true })
})

describe('pitchFamilyFromPanose', () => {
  it('sets the pitch nibble to fixed only for a fixed-pitch face', () => {
    const panose = [2, 11, 6, 9, 2, 2, 4, 3, 2, 4]
    expect(pitchFamilyFromPanose(panose, true) & 0x0f).toBe(1)
    expect(pitchFamilyFromPanose(panose, false) & 0x0f).toBe(2)
    expect(pitchFamilyFromPanose(null, true) & 0x0f).toBe(1)
    expect(pitchFamilyFromPanose(null, false) & 0x0f).toBe(2)
  })

  it('maps the PANOSE family type onto the LOGFONT family nibble', () => {
    const familyNibble = (familyType: number): number =>
      pitchFamilyFromPanose([familyType, 0, 0, 0, 0, 0, 0, 0, 0, 0], false) >> 4
    expect(familyNibble(0)).toBe(0) // don't care
    expect(familyNibble(1)).toBe(1) // roman / serif
    expect(familyNibble(2)).toBe(2) // swiss / sans
    expect(familyNibble(3)).toBe(2)
    expect(familyNibble(8)).toBe(2)
    expect(familyNibble(9)).toBe(4) // script
    expect(familyNibble(10)).toBe(4)
    expect(familyNibble(11)).toBe(5) // decorative
  })

  it('falls back to the don\'t-care family for unknown family types instead of guessing', () => {
    for (const familyType of [12, 255, -1, 1.5, Number.NaN]) {
      expect(pitchFamilyFromPanose([familyType, 0, 0, 0, 0, 0, 0, 0, 0, 0], true)).toBe(0x01)
    }
    expect(pitchFamilyFromPanose([], true)).toBe(0x01)
  })
})

describe('fontLookup', () => {
  it('matches a family case-insensitively and ignoring spaces, underscores, and hyphens', () => {
    const faces = [
      faceOf({ family: 'Microsoft YaHei', subfamily: 'Regular', weightClass: 400 }),
      faceOf({ family: 'Microsoft YaHei', subfamily: 'Bold', weightClass: 700 }),
    ]
    for (const query of ['Microsoft YaHei', 'microsoftyahei', 'Microsoft_YaHei', 'MICROSOFT-YA-HEI']) {
      const result = fontLookup(query, faces)
      expect(result.installed).toBe(true)
      expect(result.faces).toHaveLength(2)
      expect(result.matched).toMatchObject({ family: 'Microsoft YaHei', subfamily: 'Regular' })
    }
  })

  it('prefers the regular weight and returns an empty result for an unknown family', () => {
    const faces = [
      faceOf({ family: 'Consolas', subfamily: 'Bold', weightClass: 700 }),
      faceOf({ family: 'Consolas', subfamily: 'Light', weightClass: 300 }),
    ]
    expect(fontLookup('consolas', faces).matched).toMatchObject({ subfamily: 'Light' })

    const missing = fontLookup('Not Installed Anywhere', faces)
    expect(missing.installed).toBe(false)
    expect(missing.matched).toBeUndefined()
    expect(missing.faces).toEqual([])
  })
})

describe('listInstalledFonts', () => {
  it.skipIf(trueTypeSource === undefined)('describes every face of the copied system fonts', async () => {
    const faces = await scan()
    expect(faces.length).toBeGreaterThanOrEqual(1)
    for (const face of faces) {
      expect(face.family.length).toBeGreaterThan(0)
      expect(face.fullName.length).toBeGreaterThan(0)
      expect(face.panose === null || PANOSE_HEX.test(face.panose)).toBe(true)
      expect(face.sha256).toMatch(HEX_64)
      expect(['truetype', 'cff', 'collection']).toContain(face.format)
      // The pitch nibble follows the fixed-pitch flag, mirrored by PANOSE proportion 9.
      expect(face.pitchFamily & 0x0f).toBe(face.isFixedPitch ? 1 : 2)
      expect(face.pitchFamily >> 4).toBeGreaterThanOrEqual(0)
      expect(face.pitchFamily >> 4).toBeLessThanOrEqual(5)
      expect(Number.isInteger(face.glyphCount)).toBe(true)
      expect(face.glyphCount).toBeGreaterThan(0)
      // fsType embedding rules: bitmap-only and restricted licence both deny embedding.
      const permitted = (face.fsType & 0x0200) === 0
        && ((face.fsType & 0x000e) === 0 || ((face.fsType & 0x0002) === 0 && (face.fsType & 0x000c) !== 0))
      expect(face.embeddable).toBe(permitted)
    }
    const consolas = faces.find(face => face.family === 'Consolas')
    if (consolas !== undefined) {
      // Consolas is a monospaced text face: PANOSE family type 2, proportion 9.
      expect(consolas.panose?.slice(0, 2)).toBe('02')
      expect(consolas.panose?.slice(6, 8)).toBe('09')
      expect(consolas.isFixedPitch).toBe(true)
      expect(consolas.supportsLatin).toBe(true)
      expect(consolas.supportsCjk).toBe(false)
      expect(consolas.charset).toBe(0)
      expect(consolas.format).toBe('truetype')
    }
  })

  it.skipIf(collectionSource === undefined)('.ttc faces of one file share a single read and one sha256', async () => {
    const faces = await scan()
    const collected = new Set(faces.filter(face => face.format === 'collection').map(face => face.file))
    expect(collected.size).toBeGreaterThan(0)
    for (const file of collected) {
      const group = faces.filter(face => face.file === file)
      expect(group.length).toBeGreaterThan(1)
      expect(new Set(group.map(face => face.sha256)).size).toBe(1)
      // Every face of the collection is enumerated once; the global family sort
      // may interleave them, so only the set of indices is compared.
      expect(group.map(face => face.faceIndex).sort((a, b) => a - b)).toEqual(group.map((_, index) => index))
      expect(await sha256Of(file)).toBe(group[0]!.sha256)
    }
  })

  it('skips a damaged file instead of failing the whole scan', async () => {
    const faces = await scan()
    expect(faces.some(face => face.file.endsWith('broken.ttf'))).toBe(false)
    if (trueTypeSource !== undefined) expect(faces.some(face => basename(face.file) === basename(trueTypeSource))).toBe(true)
  })

  it.skipIf(trueTypeSource === undefined)('is deterministic across runs and ordered by family', async () => {
    const first = await scan()
    const second = await scan()
    expect(second.map(face => `${face.family}/${face.subfamily}/${face.faceIndex}`))
      .toEqual(first.map(face => `${face.family}/${face.subfamily}/${face.faceIndex}`))
    const families = first.map(face => face.family)
    expect([...families].sort((a, b) => a.localeCompare(b))).toEqual(families)
    for (const file of new Set(first.map(face => face.file))) {
      const weights = first.filter(face => face.file === file).map(face => face.weightClass)
      expect([...weights].sort((a, b) => a - b)).toEqual(weights)
    }
  })
})

describe('subsetFontForText', () => {
  it.skipIf(trueTypeSource === undefined)('produces sfnt data and reports unrenderable code points', async () => {
    const faces = await scan()
    const consolas = faces.find(face => face.family === 'Consolas')
    const target = consolas ?? faces[0]!
    const result = await subsetFontForText(target, 'Hello 世界')
    const magic = Buffer.from(result.data.subarray(0, 4)).toString('hex')
    // fontkit 2.0.4 emits Apple's 'true' (0x74727565) for TrueType subsets and a
    // bare CFF stream for CFF faces; every sfnt variant is accepted here.
    expect(['00010000', '74727565', '4f54544f']).toContain(magic)
    expect(result.data.byteLength).toBeGreaterThan(0)
    expect(result.glyphCount).toBeGreaterThanOrEqual(1)
    if (consolas !== undefined) {
      expect(result.missingCodePoints).toEqual(['U+4E16', 'U+754C'])
    } else {
      expect(Array.isArray(result.missingCodePoints)).toBe(true)
    }
  })

  it.skipIf(trueTypeSource === undefined)('rejects empty text and an unaddressable face', async () => {
    const faces = await scan()
    const target = faces[0]!
    await expect(subsetFontForText(target, '')).rejects.toMatchObject({ code: 'PPT_CREATE_INPUT_INVALID' })
    await expect(subsetFontForText({ ...target, faceIndex: 999 }, 'Hi')).rejects.toMatchObject({ code: 'PPT_DEPENDENCY_MISSING' })
    await expect(subsetFontForText({ ...target, file: join(fonts, 'broken.ttf') }, 'Hi')).rejects.toMatchObject({ code: 'PPT_DEPENDENCY_MISSING' })
  })
})

describe('installFontFile', () => {
  it.skipIf(trueTypeSource === undefined)('resolves the family without touching the disk on dryRun', async () => {
    const home = join(installHome, 'dry-run-home')
    const result = await installFontFile(trueTypeSource!, { platform: 'linux', home, env: {}, dryRun: true })
    expect(result.family).toBe('Consolas')
    expect(result.platform).toBe('linux')
    expect(result.scope).toBe('user')
    expect(result.registered).toBe(false)
    expect(result.uninstallHint).toBeNull()
    expect(slashes(result.installedPath)).toBe(`${slashes(home)}/.local/share/fonts/${basename(trueTypeSource!)}`)
    // Nothing was created anywhere below the sandbox home.
    await expect(access(home)).rejects.toThrow()
    await expect(access(result.installedPath)).rejects.toThrow()
  })

  it.skipIf(trueTypeSource === undefined)('installs into a simulated macOS home and is idempotent', async () => {
    const home = join(installHome, 'darwin-home')
    const result = await installFontFile(trueTypeSource!, { platform: 'darwin', home, env: {} })
    expect(slashes(result.installedPath)).toBe(`${slashes(home)}/Library/Fonts/${basename(trueTypeSource!)}`)
    expect(result.registered).toBe(false)
    expect(result.uninstallHint).toContain('rm ')
    const installed = await stat(result.installedPath)
    expect(await sha256Of(result.installedPath)).toBe(await sha256Of(trueTypeSource!))

    await new Promise(resolvePromise => setTimeout(resolvePromise, 20))
    const again = await installFontFile(trueTypeSource!, { platform: 'darwin', home, env: {} })
    expect(again.installedPath).toBe(result.installedPath)
    // copyFile would bump mtime; an unchanged one proves the bytes were not rewritten.
    expect((await stat(result.installedPath)).mtimeMs).toBe(installed.mtimeMs)
  })

  it.skipIf(trueTypeSource === undefined)('installs into a simulated Linux home under .local/share/fonts', async () => {
    const home = join(installHome, 'linux-home')
    const result = await installFontFile(trueTypeSource!, { platform: 'linux', home, env: {} })
    expect(slashes(result.installedPath)).toBe(`${slashes(home)}/.local/share/fonts/${basename(trueTypeSource!)}`)
    expect(result.family).toBe('Consolas')
    expect(typeof result.registered).toBe('boolean')
    expect(result.uninstallHint).toContain('fc-cache -f')
    expect(await sha256Of(result.installedPath)).toBe(await sha256Of(trueTypeSource!))
  })

  it('overrides the detected family and rejects an unsupported platform', async () => {
    const home = join(installHome, 'override-home')
    const source = trueTypeSource ?? join(fonts, 'broken.ttf')
    const result = await installFontFile(source, { platform: 'linux', home, env: {}, family: 'Custom Family', dryRun: true })
    expect(result.family).toBe('Custom Family')
    await expect(installFontFile(source, { platform: 'aix', home, env: {}, dryRun: true }))
      .rejects.toMatchObject({ code: 'PPT_PLATFORM_UNSUPPORTED' })
  })

  it.skipIf(trueTypeSource === undefined)('computes the Windows per-user path without elevating', async () => {
    const result = await installFontFile(trueTypeSource!, {
      platform: 'win32',
      dryRun: true,
      home: 'C:\\Users\\tester',
      env: { LOCALAPPDATA: 'C:\\Users\\tester\\AppData\\Local' },
    })
    expect(result.installedPath).toBe('C:\\Users\\tester\\AppData\\Local\\Microsoft\\Windows\\Fonts\\consola.ttf')
    await expect(access(result.installedPath)).rejects.toThrow()
  })

  // Deliberately skipped: the suite must never write the developer's real HKCU
  // registry or install a font for the whole machine. The Windows branch is
  // covered by the dryRun path assertion above.
  it.skip('registers the face under HKCU\\...\\Fonts through reg.exe (skipped: never touches the real registry)', () => {})
})
