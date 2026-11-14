/**
 * Text and shape animation lives in the slide's `p:timing` timeline, which
 * pptxgenjs does not emit at all. A deck opts in by declaring, per page, which
 * element plays which entrance effect; the finished package is then rewritten so
 * only `ppt/slides/slideN.xml` changes.
 *
 * The effect catalogue and every literal in the emitted XML below were taken from
 * timelines that PowerPoint 16.0 wrote itself (probed through COM), not from the
 * prose of the specification, because the specification deliberately leaves the
 * `presetID` numbering unspecified:
 *
 * - `presetClass` partitions the `presetID` numbering space. `presetID="10"` is
 *   Fade only under `presetClass="entr"`; the same number means something else
 *   under `emph`, `exit`, or `path`. This module therefore always writes both.
 * - `presetSubtype` is a direction bitmask, not a per-effect enum:
 *   1 = from top, 2 = from right, 4 = from bottom, 8 = from left, 16 = in,
 *   32 = out. Derived values are the bitwise or: 10 = 2|8 = horizontal,
 *   5 = 1|4 = vertical, 9 = top-left, 3 = top-right, 6 = bottom-right,
 *   12 = bottom-left. Confirmed against the defaults PowerPoint emits
 *   (Fade 0, Blinds 10, Box 16, Strips 12, Wheel 1, Zoom 16, Split 21).
 * - The renderer obeys the effect *nodes*, not the preset triple: `p:animEffect`
 *   carries a filter string and `p:anim` carries explicit from/to values, while
 *   the preset triple only drives the UI's effect list. That is why a direction
 *   change is expressed by rewriting the node, and why `peek` can legitimately
 *   pair `presetSubtype="4"` with `filter="wipe(up)"` (PowerPoint emits exactly
 *   that: peek's filter runs opposite to its subtype).
 * - Every entrance effect PowerPoint emitted carried a `p:set` that forces
 *   `style.visibility` to `visible`; without it the shape can stay hidden.
 */
export declare const TEXT_ANIMATION_EFFECTS: readonly ["appear", "flash-once", "fade", "dissolve", "wedge", "wipe", "blinds", "checkerboard", "random-bars", "box", "circle", "diamond", "plus", "split", "strips", "wheel", "zoom", "fly-in", "crawl", "peek", "stretch", "swivel", "spiral", "bounce", "credits", "float-in", "grow-turn", "rise-up", "unfold"];
export type TextAnimationEffect = typeof TEXT_ANIMATION_EFFECTS[number];
export declare const TEXT_ANIMATION_DIRECTIONS: readonly ["left", "right", "up", "down"];
export type TextAnimationDirection = typeof TEXT_ANIMATION_DIRECTIONS[number];
export declare const TEXT_ANIMATION_STARTS: readonly ["on-click", "with-previous", "after-previous"];
export type TextAnimationStart = typeof TEXT_ANIMATION_STARTS[number];
export interface TextAnimation {
    /** Element id from the deck IR, which pptxgenjs writes as `p:cNvPr/@name`. */
    target: string;
    effect: TextAnimationEffect;
    direction?: TextAnimationDirection;
    /** Build the effect one paragraph at a time instead of animating the whole box. */
    byParagraph?: boolean;
    start?: TextAnimationStart;
    durationMs?: number;
}
export interface SlideAnimationPlanEntry {
    page: number;
    animations: readonly TextAnimation[];
}
export type AnimationPlan = ReadonlyMap<number, readonly TextAnimation[]>;
export declare const TEXT_ANIMATION_DEFAULT_DURATION_MS = 500;
export declare const TEXT_ANIMATION_MAX_PER_PAGE = 24;
export declare function isTextAnimationEffect(value: string): value is TextAnimationEffect;
export declare function isTextAnimationDirection(value: string): value is TextAnimationDirection;
export declare function normalizeTextAnimation(animation: TextAnimation): TextAnimation;
export declare function planSlideAnimations(entries: readonly SlideAnimationPlanEntry[]): AnimationPlan;
interface ShapeIndex {
    idByName: Map<string, string>;
    paragraphCountByName: Map<string, number>;
}
export declare function animationTimingXml(animations: readonly TextAnimation[], shape: ShapeIndex): string;
export declare function injectSlideAnimation(xml: string, animations: readonly TextAnimation[]): string;
export declare function rewritePptxAnimations(data: Uint8Array, plan: AnimationPlan): Uint8Array;
export {};
