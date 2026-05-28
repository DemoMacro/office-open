/**
 * Media module for PresentationML documents.
 *
 * @module
 */
import { convertPixelsToEmu } from "@office-open/core";

import type { MediaDataTransformation } from "./data";
import type { IMediaData } from "./data";

export interface MediaTransformation {
  readonly offset?: {
    readonly top?: number;
    readonly left?: number;
  };
  readonly width: number;
  readonly height: number;
  readonly flip?: {
    readonly vertical?: boolean;
    readonly horizontal?: boolean;
  };
  readonly rotation?: number;
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
  private readonly map: Map<string, IMediaData>;

  public constructor() {
    this.map = new Map<string, IMediaData>();
  }

  public addImage(key: string, mediaData: IMediaData): void {
    this.map.set(key, mediaData);
  }

  public addMedia(key: string, mediaData: IMediaData): void {
    this.map.set(key, mediaData);
  }

  public get array(): readonly IMediaData[] {
    return [...this.map.values()];
  }
}
