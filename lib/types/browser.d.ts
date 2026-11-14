import type { SessionOwner } from './session-resources.ts';
import { SessionResourceRegistry } from './session-resources.ts';
import type { DeckIR } from './ir.ts';
export interface BrowserElementRef {
    ref: string;
    tag: string;
    text: string;
    href?: string;
    clickable: boolean;
}
export interface BrowserPageResult {
    url: string;
    title: string;
    text: string;
    page_version: number;
    content_is_untrusted: true;
    elements?: BrowserElementRef[];
}
export interface HtmlPreviewResult {
    previews: string[];
    fonts: string[];
    warnings: string[];
    designPages: Array<{
        page: number;
        anchorAreaRatio?: number;
        frameCount: number;
        occupancy: number[];
        roleStyles: Array<{
            role: string;
            fontFamily: string;
            fontWeight: number;
        }>;
    }>;
}
/** A leaf rectangle expressed relative to the slide origin. */
export interface SlideLeafBox {
    left: number;
    top: number;
    right: number;
    bottom: number;
}
export interface SlideSize {
    width: number;
    height: number;
}
/**
 * 4x4 occupancy silhouette of one slide, used to detect adjacent pages with the same layout.
 * Full-bleed layers are skipped: a page background leaf is identical on every page and would
 * light all sixteen cells everywhere, so every adjacent pair would look repeated.
 */
export declare function occupancySilhouette(leafBoxes: readonly SlideLeafBox[], slide?: SlideSize): number[];
export declare class BrowserRuntime {
    private readonly resources;
    private readonly configuredExecutable?;
    private readonly outputRoot;
    private browser?;
    private launching?;
    private readonly states;
    constructor(resources: SessionResourceRegistry, configuredExecutable?: string | undefined, outputRoot?: string);
    visit(owner: SessionOwner, workspace: string, input: string, signal?: AbortSignal): Promise<BrowserPageResult>;
    find(owner: SessionOwner, query: string, signal?: AbortSignal): Promise<BrowserPageResult>;
    click(owner: SessionOwner, ref: string, signal?: AbortSignal): Promise<BrowserPageResult>;
    scroll(owner: SessionOwner, direction: 'up' | 'down', amount?: number, signal?: AbortSignal): Promise<BrowserPageResult>;
    renderHtmlPreview(owner: SessionOwner, workspace: string, htmlPath: string, previewDirectory: string, pageCount: number, allowedFonts: readonly string[], signal?: AbortSignal): Promise<HtmlPreviewResult>;
    extractDeckIr(owner: SessionOwner, workspace: string, htmlPath: string, pageCount: number, signal?: AbortSignal): Promise<DeckIR>;
    rasterizeElement(owner: SessionOwner, workspace: string, htmlPath: string, elementId: string, targetInput: string, signal?: AbortSignal): Promise<string>;
    dispose(): Promise<void>;
    private launch;
    private requireState;
    private requireExisting;
    private resolveVisitUrl;
    private resolveLocalFileUrl;
    private validateCurrentUrl;
    private bumpVersion;
    private refreshMutationVersion;
    private pageResult;
    private cancellable;
    private abortOwnerIfRequested;
}
