import type { MediaFrameBaseOptions } from "./media-frame-base";

export type VideoType = "mp4" | "mov" | "wmv" | "avi";
export type PosterType = "png" | "jpg";

export interface VideoFrameOptions extends MediaFrameBaseOptions {
  type: VideoType;
  poster?: Uint8Array;
  posterType?: PosterType;
}
