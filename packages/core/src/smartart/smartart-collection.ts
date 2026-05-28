import type { DataModel } from "./data-model/data-model";

export interface SmartArtData {
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
  private readonly map: Map<string, SmartArtData>;

  public constructor() {
    this.map = new Map<string, SmartArtData>();
  }

  public addSmartArt(key: string, data: SmartArtData): void {
    this.map.set(key, data);
  }

  public get array(): readonly SmartArtData[] {
    return [...this.map.values()];
  }
}
