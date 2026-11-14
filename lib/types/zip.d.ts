/**
 * OOXML parts that pptxgenjs cannot produce — slide transitions, animation
 * timelines, embedded font parts — are injected by rewriting the finished
 * package. Untouched entries keep their original local header bytes, so a
 * package that is not targeted stays byte-comparable with what pptxgenjs wrote.
 */
export interface ZipEntryInfo {
    name: string;
    method: number;
    compressed_size: number;
    uncompressed_size: number;
    local_header_offset: number;
}
export interface ZipEntry extends ZipEntryInfo {
    flags: number;
    version_needed: number;
    time: number;
    date: number;
    crc: number;
    local_extra: Uint8Array;
    data_offset: number;
    header: Uint8Array;
    central: Uint8Array;
}
export interface ZipDirectory {
    entries: ZipEntry[];
    comment: Uint8Array;
}
export declare const ZIP_METHOD_STORE = 0;
export declare const ZIP_METHOD_DEFLATE = 8;
export declare function crc32(data: Uint8Array): number;
export declare function readZipDirectory(data: Uint8Array): ZipDirectory;
export declare function readEntryBytes(entry: ZipEntry, data: Uint8Array): Uint8Array;
export declare function inspectZipEntries(data: Uint8Array): ZipEntryInfo[];
export interface ZipRewritePlan {
    /**
     * Called once per entry with a lazy reader. Returning replacement bytes swaps
     * the part, `undefined` keeps it byte-identical, and `null` drops it. Entries
     * whose content is never read stay compressed on disk exactly as they were.
     */
    transform?: (name: string, read: () => Uint8Array) => Uint8Array | null | undefined;
    /** Parts to append; an existing entry with the same name is an error. */
    additions?: ReadonlyMap<string, Uint8Array>;
}
export declare function rewriteZip(data: Uint8Array, plan: ZipRewritePlan): Uint8Array;
