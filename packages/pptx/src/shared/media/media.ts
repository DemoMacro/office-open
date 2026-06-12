/**
 * Media module for PresentationML documents.
 *
 * @module
 */
import { convertPixelsToEmu } from "@office-open/core";

import type { MediaDataTransformation } from "./data";
import type { MediaData } from "./data";

export interface MediaTransformation {
  offset?: {
    top?: number;
    left?: number;
  };
  width: number;
  height: number;
  flip?: {
    vertical?: boolean;
    horizontal?: boolean;
  };
  rotation?: number;
}

export const createTransformation = (options: MediaTransformation): MediaDataTransformation => ({
  emus: {
    x: convertPixelsToEmu(options.width),
    y: convertPixelsToEmu(options.height),
  },
  pixels: {
    x: Math.round(options.width),
    y: Math.round(options.height),
  },
  flip: options.flip,
  rotation: options.rotation ? options.rotation * 60_000 : undefined,
});

export class Media {
  private map: Map<string, MediaData>;

  public constructor() {
    this.map = new Map<string, MediaData>();
  }

  public addImage(key: string, mediaData: MediaData): void {
    this.map.set(key, mediaData);
  }

  public addMedia(key: string, mediaData: MediaData): void {
    this.map.set(key, mediaData);
  }

  public get array(): MediaData[] {
    return [...this.map.values()];
  }
}
