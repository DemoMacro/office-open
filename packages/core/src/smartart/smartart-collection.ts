/**
 * SmartArt data and collection for document generation.
 *
 * @module
 */

export interface SmartArtData {
  key: string;
  dataModelXml: string;
  layout: string;
  style: string;
  color: string;
}

/**
 * Manages SmartArt parts in a document.
 */
export class SmartArtCollection {
  private map: Map<string, SmartArtData>;

  public constructor() {
    this.map = new Map<string, SmartArtData>();
  }

  public addSmartArt(key: string, data: SmartArtData): void {
    this.map.set(key, data);
  }

  public get array(): SmartArtData[] {
    return [...this.map.values()];
  }
}
