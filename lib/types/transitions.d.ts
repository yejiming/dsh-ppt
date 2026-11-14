/**
 * Slide transitions live outside the pptxgenjs feature set, so a deck opts in by
 * declaring a per-page plan and letting the PPTX writer rewrite only the slide XML
 * parts after pptxgenjs produced the package. Every other zip entry is copied
 * verbatim so the result stays byte-comparable with the untouched archive.
 */
export declare const SLIDE_TRANSITION_TYPES: readonly ["cut", "fade", "dissolve", "push", "wipe", "cover", "pull"];
export type SlideTransitionType = typeof SLIDE_TRANSITION_TYPES[number];
export declare const SLIDE_TRANSITION_DIRECTIONS: readonly ["left", "right", "up", "down"];
export type SlideTransitionDirection = typeof SLIDE_TRANSITION_DIRECTIONS[number];
export declare const SLIDE_TRANSITION_SPEEDS: readonly ["slow", "med", "fast"];
export type SlideTransitionSpeed = typeof SLIDE_TRANSITION_SPEEDS[number];
export interface SlideTransition {
    type: SlideTransitionType;
    direction?: SlideTransitionDirection;
    speed?: SlideTransitionSpeed;
    advanceAfterMs?: number;
    advanceOnClick?: boolean;
}
export interface SlideRhythmPage {
    page: number;
    type: string;
}
export type SlideTransitionPlan = ReadonlyMap<number, SlideTransition>;
export { inspectZipEntries, type ZipEntryInfo } from './zip.ts';
export declare function isSlideTransitionType(value: string): value is SlideTransitionType;
export declare function normalizeSlideTransition(transition: SlideTransition): SlideTransition;
export declare function transitionElementXml(transition: SlideTransition): string;
export declare function injectSlideTransition(xml: string, transition: SlideTransition): string;
export declare function planSlideTransitions(pages: readonly SlideRhythmPage[]): SlideTransitionPlan;
export declare function rewritePptxTransitions(data: Uint8Array, plan: SlideTransitionPlan): Uint8Array;
