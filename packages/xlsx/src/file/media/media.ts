/**
 * Media collection for XLSX files — stores image binary data.
 *
 * @module
 */

export interface MediaData {
  readonly fileName: string;
  readonly type: string;
  readonly data: Uint8Array;
  readonly width: number;
  readonly height: number;
}

export class Media {
  private readonly map = new Map<string, MediaData>();

  public addImage(key: string, data: MediaData): void {
    this.map.set(key, data);
  }

  public get array(): readonly MediaData[] {
    return [...this.map.values()];
  }
}
