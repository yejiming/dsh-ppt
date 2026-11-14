import { ART_BACKGROUNDS, ART_COMPOSITIONS, ART_DENSITIES, ART_FRAME_POLICIES, ART_TITLE_TREATMENTS, type ArtDirection } from './art-direction.ts';
import { SLIDE_TYPES } from './slide-taxonomy.ts';
/**
 * A deck theme is a reusable design recipe: colours, type roles, a page rhythm
 * and the decoration vocabulary that holds them together. The plugin shipped
 * without one, so every deck was designed from nothing and the quality ceiling
 * was whatever the model improvised that turn. A theme does not replace the
 * agent's design pass — it gives the pass a vetted starting point and a
 * verifiable palette, the way a template does for a human designer.
 */
export type ArtComposition = typeof ART_COMPOSITIONS[number];
export type ArtBackground = typeof ART_BACKGROUNDS[number];
export type ArtDensity = typeof ART_DENSITIES[number];
export type ArtTitleTreatment = typeof ART_TITLE_TREATMENTS[number];
export type ArtFramePolicy = typeof ART_FRAME_POLICIES[number];
export interface ThemeFontRole {
    family: string;
    weight: number;
}
export interface PptTheme {
    id: string;
    /** Display name shown to the user when the theme is offered. */
    name: string;
    concept: string;
    audience_effect: string;
    scenes: readonly string[];
    /**
     * Where the palette values come from. Every built-in theme traces to a
     * published design system so a colour can be audited instead of trusted; a
     * theme authored from a brand book or a public palette site records that
     * source here too.
     */
    palette_source: string;
    palette: {
        background: readonly string[];
        surface: readonly string[];
        accent: string;
        /**
         * Accent for the inverted tonal group. A single accent cannot clear 4.5:1 on
         * both a light and a dark field, so every theme carries the pair and the plan
         * tells the model which one each page uses.
         */
        accent_inverted: string;
        text: readonly string[];
    };
    typography: {
        display: ThemeFontRole;
        body: ThemeFontRole;
        latin: ThemeFontRole;
        code: ThemeFontRole;
    };
    /** Composition preference, repeated across the deck in page order. */
    composition_cycle: readonly ArtComposition[];
    /** Background rhythm preference, repeated across the deck in page order. */
    background_cycle: readonly ArtBackground[];
    /** Rhythm limits the theme's own decoration was designed against. */
    rhythm_limits: {
        max_grouped_frame_slides: number;
        max_same_composition_run: number;
    };
    /** Decoration moves a page may use; every entry must survive the HTML whitelist. */
    decoration: readonly string[];
    /** What each composition means inside this theme. */
    layout_notes: Readonly<Partial<Record<ArtComposition, string>>>;
}
/** WCAG 2.1 relative luminance of an #RRGGBB colour. */
export declare function relativeLuminance(color: string): number;
/** WCAG 2.1 contrast ratio between two #RRGGBB colours, from 1 to 21. */
export declare function contrastRatio(a: string, b: string): number;
export interface ThemeFinding {
    code: string;
    severity: 'warning' | 'error';
    message: string;
}
/**
 * Contrast is the one part of "looks good" that can be checked by machine, so
 * every theme is held to it at load time rather than discovered in review.
 *
 * A palette is a set of tonal groups, not a cross product: `background[i]`,
 * `surface[i]` and `text[i]` belong together, so text is paired by index. The
 * accent is a single value that must stay readable on the primary group; on an
 * inverted group it is expected to be swapped for the theme's inverted accent,
 * which is reported as a warning rather than a failure.
 */
export declare function themeFindings(theme: PptTheme): ThemeFinding[];
export declare function validateTheme(theme: PptTheme): PptTheme;
export interface ThemeSummary {
    id: string;
    name: string;
    concept: string;
    /** Copied out of the theme so the catalogue is plain mutable JSON for the tool wire. */
    scenes: string[];
    palette_source: string;
    accent: string;
    accent_inverted: string;
    display_font: string;
    body_font: string;
    signature: string;
}
export interface ThemeCatalog {
    themes: ThemeSummary[];
    usage: string;
    warnings: string[];
}
export declare function themeSummary(theme: PptTheme): ThemeSummary;
export declare function listThemes(themes?: readonly PptTheme[], scene?: string): ThemeCatalog;
export declare function findTheme(id: string, themes?: readonly PptTheme[]): PptTheme;
export interface ThemePagePlan {
    page: number;
    type: typeof SLIDE_TYPES[number];
    composition: ArtComposition;
    density: ArtDensity;
    background_role: ArtBackground;
    title_treatment: ArtTitleTreatment;
    frame_policy: ArtFramePolicy;
}
/**
 * Expand the theme's cycles into one visual plan per page. The caller keeps the
 * content decisions (job, takeaway, visual anchor), so this returns only the
 * structural half of an Art Direction plan.
 */
export declare function planThemePages(theme: PptTheme, types: readonly (typeof SLIDE_TYPES[number])[]): ThemePagePlan[];
/**
 * The theme fields an Art Direction plan inherits verbatim. Everything else in
 * the plan stays the agent's own composition work.
 */
export declare function themePaletteAndType(theme: PptTheme): Pick<ArtDirection, 'palette' | 'typography'>;
/**
 * Confirm an authored plan still belongs to the theme it claims. Palette drift
 * is the failure this catches: a deck that names a theme and then uses unrelated
 * colours reads as neither.
 */
export declare function themeFindingsForPlan(theme: PptTheme, plan: ArtDirection): ThemeFinding[];
/**
 * Twelve built-in themes. Every palette clears WCAG AA (4.5:1) for body text
 * against every background and surface it defines, and for the accent as well,
 * because the accent carries emphasis text and not only rules. The palettes come
 * from Carbon, Tailwind, Radix, Fluent 2 and Open Color values; the typography is
 * restricted to families this build can actually render (see FONT_REGISTRY).
 */
export declare const PPT_THEMES: readonly PptTheme[];
