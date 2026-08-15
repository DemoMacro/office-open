import type { DataType } from "@office-open/core";

import type { MediaFrameBaseOptions } from "./media-frame-base";

export type VideoType = "mp4" | "mov" | "wmv" | "avi";
export type PosterType = "png" | "jpg";

export interface VideoFrameOptions extends MediaFrameBaseOptions {
  type: VideoType;
  poster?: DataType;
  posterType?: PosterType;
  /** MIME content type of the linked video (CT_VideoFile @contentType) */
  contentType?: string;
}
