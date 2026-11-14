import { type SourceAssetRecord } from './artifacts.ts';
export type ImageOrientation = 'landscape' | 'portrait' | 'square' | 'any';
export type BuiltinImageProvider = 'openverse' | 'wikimedia-commons';
/** Built-in providers keep literal names; the open branch lets an injected backend report its own name. */
export type ImageProvider = BuiltinImageProvider | (string & {});
export interface ImageCandidate {
    image_url: string;
    source_page: string;
    provider: ImageProvider;
    title: string;
    license: string;
    license_verified: false;
    thumbnail_url?: string;
    width?: number;
    height?: number;
    mime_type?: string;
    author?: string;
    license_url?: string;
    attribution?: string;
}
export interface ImageSearchResult {
    query: string;
    count: number;
    orientation: ImageOrientation;
    cache_hit: boolean;
    providers_used: ImageProvider[];
    warnings: string[];
    results: ImageCandidate[];
}
/** A zero-configuration retrieval backend. Built-in providers are always tried first. */
export interface ImageSearchBackend {
    readonly name: string;
    search(query: string, amount: number, orientation: ImageOrientation, signal?: AbortSignal): Promise<ImageCandidate[]>;
}
export interface ImageSearchRuntimeOptions {
    /** Extra backends appended after Openverse and Wikimedia Commons; empty by default. */
    providers?: readonly ImageSearchBackend[];
    /** Opt back into the legacy hard failure when no provider is reachable. Defaults to false. */
    strictOnUnavailable?: boolean;
}
export interface ImageSearchFallback {
    id: string;
    tool: string;
    summary: string;
}
/** Actionable alternatives carried by a degraded result when the free providers deliver nothing. */
export declare const IMAGE_SEARCH_FALLBACKS: readonly ImageSearchFallback[];
export type ImageSearchStatus = 'ok' | 'partial' | 'unavailable';
export interface ImageSearchDegradation {
    status: ImageSearchStatus;
    requested: number;
    returned: number;
    failures: Array<{
        provider: string;
        kind: string;
    }>;
    fallbacks: ImageSearchFallback[];
}
/**
 * Rebuilds the structured story behind the machine-readable warnings of one search result.
 * It takes the structural subset it actually reads, so a payload derived from the tool
 * output schema (where `license_verified` widens to `boolean`) stays assignable.
 */
export declare function describeImageSearchDegradation(result: {
    readonly count: number;
    readonly providers_used: readonly string[];
    readonly warnings: readonly string[];
    readonly results: readonly unknown[];
}): ImageSearchDegradation;
type FetchLike = (input: string | URL | Request, init?: RequestInit) => Promise<Response>;
type UrlValidator = (input: string) => Promise<URL>;
export declare class ImageSearchRuntime {
    private readonly fetcher;
    private readonly validateUrl;
    private readonly cache;
    private readonly backends;
    private readonly strictOnUnavailable;
    constructor(fetcher?: FetchLike, validateUrl?: UrlValidator, options?: ImageSearchRuntimeOptions);
    search(queryInput: string, countInput?: number, orientation?: ImageOrientation, signal?: AbortSignal): Promise<ImageSearchResult>;
    private request;
    private openverse;
    private commons;
}
export interface FrozenImageAsset {
    path: string;
    width: number;
    height: number;
    mime_type: string;
    size: number;
    manifest: SourceAssetRecord;
}
export declare function freezeImageAsset(workspace: string, artifactRoot: string, candidate: ImageCandidate, fetcher?: FetchLike, signal?: AbortSignal): Promise<FrozenImageAsset>;
export {};
