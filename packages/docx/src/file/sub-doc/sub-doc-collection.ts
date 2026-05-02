/**
 * Sub-document collection module for managing sub-document parts.
 *
 * @module
 */

/**
 * Stores sub-document data for later serialization by the compiler.
 */
export interface ISubDocData {
    /** Raw document data (.docx bytes) */
    readonly data: Uint8Array;
    /** Part sub-path within word/ (e.g., "subdocs/subdoc1.docx") */
    readonly path: string;
}

/**
 * Manages sub-document parts in a document.
 */
export class SubDocCollection {
    private readonly map: Map<string, ISubDocData>;

    public constructor() {
        this.map = new Map<string, ISubDocData>();
    }

    public addSubDoc(key: string, data: ISubDocData): void {
        this.map.set(key, data);
    }

    public get Array(): readonly ISubDocData[] {
        return [...this.map.values()];
    }
}
