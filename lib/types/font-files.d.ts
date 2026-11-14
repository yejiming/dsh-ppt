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
export declare const MAX_FONT_FILES = 10000;
export declare const FONT_FILE_EXTENSIONS: readonly string[];
export interface InstalledFontFace {
    family: string;
    subfamily: string;
    postscriptName: string | null;
    fullName: string;
    /** Absolute path of the font file on disk. */
    file: string;
    /** For .ttc collections this is the index of the face inside the file, otherwise 0. */
    faceIndex: number;
    format: 'truetype' | 'cff' | 'collection';
    weightClass: number;
    isFixedPitch: boolean;
    fsType: number;
    /** True when the OS/2 fsType permission bits allow embedding. */
    embeddable: boolean;
    /** 10 OS/2 PANOSE bytes as 20 uppercase hex characters, or null when the face has no OS/2 table. */
    panose: string | null;
    /** Windows LOGFONT-style pitch/family byte derived from PANOSE plus the fixed-pitch flag. */
    pitchFamily: number;
    /**
     * Suggested `lfCharSet` hint: GB2312 (-122), Shift-JIS (-128), Hangul (-127),
     * or ANSI (0) when the writing system cannot be established from the family
     * name. See `CHINESE_FAMILY_ROOTS` and `charsetOf` for the rule.
     */
    charset: number;
    glyphCount: number;
    supportsLatin: boolean;
    supportsCjk: boolean;
    sha256: string;
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
export declare function pitchFamilyFromPanose(panose: readonly number[] | null, isFixedPitch: boolean): number;
/**
 * Enumerates every face installed for the requested platform plus `extraDirs`.
 *
 * A file that cannot be read, cannot be parsed as sfnt, or that fontkit refuses
 * is skipped: one damaged file in a font directory must not fail the whole scan.
 * File contents are read exactly once per file, so all faces of a collection
 * share one buffer read and one sha256.
 */
export declare function listInstalledFonts(extraDirs?: readonly string[], platform?: NodeJS.Platform): Promise<InstalledFontFace[]>;
export interface FontLookupResult {
    family: string;
    matched: InstalledFontFace | undefined;
    faces: InstalledFontFace[];
    /** True when the family was found on this machine. */
    installed: boolean;
}
/**
 * Case-insensitive family lookup. Spaces, underscores, and hyphens are ignored,
 * so "Microsoft YaHei", "microsoftyahei", and "Microsoft_YaHei" all resolve to
 * the same family. When no family matches exactly, the PostScript name is
 * retried, which is how `.ttc` faces whose family name is localised still answer.
 */
export declare function fontLookup(family: string, faces: readonly InstalledFontFace[]): FontLookupResult;
export interface FontSubsetResult {
    data: Uint8Array;
    /** Glyphs in the emitted font, including `.notdef` and any compound components fontkit pulled in. */
    glyphCount: number;
    /** Code points from the request that the face cannot render. */
    missingCodePoints: string[];
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
export declare function subsetFontForText(face: InstalledFontFace, text: string, options?: {
    keepLayoutTables?: boolean;
}): Promise<FontSubsetResult>;
export interface FontInstallResult {
    family: string;
    installedPath: string;
    platform: NodeJS.Platform;
    scope: 'user';
    registered: boolean;
    /** Command that undoes the install, or null when a manual removal is required. */
    uninstallHint: string | null;
}
/**
 * `platform`, `home`, and `env` are injectable for the same reason every helper
 * in `src/platform.ts` takes them: the per-user targets differ per OS and the
 * unit tests must be able to exercise the macOS/Linux branches on any host
 * without writing into a real home directory. They are optional overrides, so
 * the documented `{ family, dryRun }` call shape is unchanged.
 */
export interface FontInstallOptions {
    family?: string;
    dryRun?: boolean;
    platform?: NodeJS.Platform;
    home?: string;
    env?: NodeJS.ProcessEnv;
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
export declare function installFontFile(sourcePath: string, options?: FontInstallOptions): Promise<FontInstallResult>;
