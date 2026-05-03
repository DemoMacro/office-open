import type { DataModel } from "./data-model/data-model";

export interface ISmartArtData {
    readonly key: string;
    readonly dataModel: DataModel;
    readonly layout: string;
    readonly style: string;
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
