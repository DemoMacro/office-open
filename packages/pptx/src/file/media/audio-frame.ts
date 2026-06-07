import type { MediaFrameBaseOptions } from "./media-frame-base";

export type AudioType = "mp3" | "wav" | "wma" | "aac";

export interface AudioFrameOptions extends MediaFrameBaseOptions {
  readonly type: AudioType;
}
