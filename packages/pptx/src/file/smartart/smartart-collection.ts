/**
 * SmartArt collection module for managing diagram data parts.
 *
 * @module
 */
import type { DataModel } from "./data-model/data-model";

/**
 * Stores SmartArt data for later serialization by the compiler.
 */
export interface ISmartArtData {
    /** Unique key for this SmartArt */
    readonly key: string;
    /** The DataModel XML component */
    readonly dataModel: DataModel;
    /** Layout ID (e.g. "default", "process1", "hierarchy1") */
    readonly layout: string;
    /** Quick style ID (e.g. "simple1", "moderate1") */
    readonly style: string;
    /** Color transform ID (e.g. "accent1_2", "colorful1") */
    readonly color: string;
}

/**
 * Manages SmartArt parts in a document.
 */
export class SmartArtCollection {
    private readonly map: Map<string, ISmartArtData>;

    public constructor() {
        this.map = new Map<string, ISmartArtData>();
    }

    public addSmartArt(key: string, data: ISmartArtData): void {
        this.map.set(key, data);
    }

    public get Array(): readonly ISmartArtData[] {
        return [...this.map.values()];
    }
}
