/**
 * AltChunk collection module for managing alternative format content parts.
 *
 * @module
 */

/**
 * Stores alternative format chunk data for later serialization by the compiler.
 */
export interface IAltChunkData {
    /** Unique key for this alt chunk (e.g., relId) */
    readonly key: string;
    /** Raw content data */
    readonly data: Uint8Array;
    /** Part sub-path within word/ (e.g., "afchunks/afchunk1.html") */
    readonly path: string;
    /** File extension (e.g., "html", "rtf", "txt") */
    readonly extension: string;
    /** MIME content type (e.g., "text/html", "application/rtf") */
    readonly contentType: string;
}

/**
 * Manages alternative format chunk parts in a document.
 *
 * Stores external content (HTML, RTF, plain text) that will be
 * serialized into separate parts in the DOCX package.
 */
export class AltChunkCollection {
    private readonly map: Map<string, IAltChunkData>;

    public constructor() {
        this.map = new Map<string, IAltChunkData>();
    }

    public addAltChunk(key: string, data: IAltChunkData): void {
        this.map.set(key, data);
    }

    public get Array(): readonly IAltChunkData[] {
        return [...this.map.values()];
    }
}
