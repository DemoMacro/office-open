import type { BlipCompression, DataType } from "@office-open/core";

import type { MediaFrameBaseOptions } from "./media-frame-base";

/** Video container format of the media file. */
export type VideoType = "mp4" | "mov" | "wmv" | "avi";
/** Poster-frame image format. */
export type PosterType = "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";

export interface VideoFrameOptions extends MediaFrameBaseOptions {
  type: VideoType;
  poster?: DataType;
  posterType?: PosterType;
  /** Compression state of the poster blip; absent = attribute omitted. */
  posterCompression?: BlipCompression;
  /** Source poster media file name (round-trip only). */
  posterFileName?: string;
  /** MIME content type of the linked video (CT_VideoFile `@contentType`) */
  contentType?: string;
}
