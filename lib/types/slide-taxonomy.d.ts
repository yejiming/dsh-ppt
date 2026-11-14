/**
 * Slide taxonomy shared by the outline schema, the art direction schema and the
 * built-in theme library. It lives in its own module because all three need it:
 * keeping it in `outline.ts` made `outline.ts -> themes.ts -> outline.ts` a
 * cycle, which breaks the moment any of those modules is loaded first.
 */
export declare const SLIDE_TYPES: readonly ["cover", "agenda", "section", "content", "comparison", "timeline", "process", "data", "quote", "summary", "ending"];
export declare const SLIDE_LAYOUTS: readonly ["cover", "center", "title-content", "split", "two-column", "three-column", "grid", "hero-image", "image-left", "image-right", "timeline-horizontal", "timeline-vertical", "process-horizontal", "process-vertical", "chart-focus", "quote-focus", "full-bleed", "closing"];
export type SlideType = typeof SLIDE_TYPES[number];
export type SlideLayout = typeof SLIDE_LAYOUTS[number];
/**
 * Which layouts can carry which slide roles. A layout that appears here for the
 * wrong role is rejected before anything is rendered, so this table is the
 * single authority behind both the schema and the model-facing layout guide.
 */
export declare const SLIDE_TYPE_LAYOUTS: Record<SlideLayout, readonly SlideType[]>;
/** Slide roles that may legitimately carry no visible content item. */
export declare const CONTENT_OPTIONAL_TYPES: readonly SlideType[];
/** Slide roles whose layout must be compatible with an ordered point sequence. */
export declare const SEQUENCED_TYPES: readonly SlideType[];
