import { BuilderElement } from "@file/xml-components";

import { MediaFrameBase, type MediaFrameBaseOptions } from "./media-frame-base";

const MEDIA_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

/** Minimal 1x1 transparent PNG (67 bytes). */
const MINIMAL_PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0,
  0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 8, 215, 99, 24, 5, 163, 0, 0, 0, 2, 0, 1, 226,
  33, 188, 51, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

export type VideoType = "mp4" | "mov" | "wmv" | "avi";
export type PosterType = "png" | "jpg";

export interface VideoFrameOptions extends MediaFrameBaseOptions {
  readonly type: VideoType;
  readonly poster?: Uint8Array;
  readonly posterType?: PosterType;
}

/**
 * p:pic — A video frame on a slide.
 *
 * Uses three relationships:
 * - Image relationship for poster frame (via {posterFileName} placeholder)
 * - Video relationship for video file (via {video:fileName} placeholder in a:videoFile r:link)
 * - Media relationship for video file (via {media:fileName} placeholder in p14:media r:embed)
 */
export class VideoFrame extends MediaFrameBase {
  private static nextId = 100;

  public constructor(options: VideoFrameOptions) {
    const id = VideoFrame.nextId++;
    const name = options.name ?? `Video ${id}`;
    const mediaFileName = `${name.replace(/\s+/g, "_")}.${options.type}`;
    const posterBytes = options.poster ?? MINIMAL_PNG;
    const posterType = options.posterType ?? "png";
    const posterFileName = `${name.replace(/\s+/g, "_")}_poster.${posterType}`;

    super(options, id, mediaFileName, {
      extUri: MEDIA_EXT_URI,
      cNvPrPrefix: "p",
      nvPrExtraChildren: [
        new BuilderElement({
          name: "a:videoFile",
          attributes: {
            "r:link": { key: "r:link", value: `{video:${mediaFileName}}` },
          },
        }),
      ],
      posterBytes,
      posterType,
      posterFileName,
    });
  }
}
