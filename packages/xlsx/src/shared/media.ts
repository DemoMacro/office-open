/**
 * Media collection for XLSX files — stores image binary data.
 *
 * @module
 */

export interface MediaData {
  fileName: string;
  type: string;
  data: Uint8Array;
  width: number;
  height: number;
}

export class Media {
  private map = new Map<string, MediaData>();

  public addImage(key: string, data: MediaData): void {
    this.map.set(key, data);
  }

  public get array(): MediaData[] {
    return [...this.map.values()];
  }
}
