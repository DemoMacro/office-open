import { MediaFrameBase, type MediaFrameBaseOptions } from "./media-frame-base";

const MEDIA_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";

export type AudioType = "mp3" | "wav" | "wma" | "aac";

export interface AudioFrameOptions extends MediaFrameBaseOptions {
  readonly type: AudioType;
}

/**
 * p:pic — An audio frame on a slide.
 *
 * Uses a media relationship for the audio file (via {media:fileName} placeholder).
 */
export class AudioFrame extends MediaFrameBase {
  private static nextId = 200;

  public constructor(options: AudioFrameOptions) {
    const id = AudioFrame.nextId++;
    const name = options.name ?? `Audio ${id}`;
    const mediaFileName = `${name.replace(/\s+/g, "_")}.${options.type}`;

    super(options, id, mediaFileName, {
      extUri: MEDIA_EXT_URI,
      cNvPrPrefix: "p",
    });
  }
}
