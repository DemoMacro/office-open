import type {
  DataType,
  NonVisualDrawingPropertiesOptions,
  PictureLockingOptions,
  SourceRectangleOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { MediaData } from "@shared/media/data";

/**
 * Common options for all media frames. The cNvPr fields
 * (name/description/title/hidden) come from
 * {@link NonVisualDrawingPropertiesOptions}.
 * @internal
 */
export interface MediaFrameBaseOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  data: DataType;
  type: MediaData["type"];
  /**
   * Media file name inside the package (ppt/media/<fileName>). Round-trip
   * keeps the source name; fresh generation derives it from the frame name.
   */
  fileName?: string;
  /**
   * Click-to-play hyperlink (a:hlinkClick `action="ppaction://media"` on
   * p:cNvPr) — Office emits it on media frames.
   */
  mediaAction?: boolean;
  /** Picture locks (a:picLocks inside p:cNvPicPr). */
  locking?: PictureLockingOptions;
  /** Resize relative to the original media rather than the frame (p:cNvPicPr attribute). */
  preferRelativeResize?: boolean;
  /**
   * Play window trim of the p14:media extension copy (p14:trim, seconds).
   * undefined = no trim child.
   */
  trim?: MediaTrimOptions;
  /** Player settings carried by ODF `draw:param`; OOXML writers ignore unsupported values. */
  playback?: MediaPlaybackOptions;
  /** Bare `<a:stretch/>` (no a:fillRect child) on the blip fill — round-trip. */
  stretchBare?: boolean;
  /** Rotate image fill with the shape frame (p:blipFill @rotWithShape). */
  rotWithShape?: boolean;
  /** Crop rectangle (p:blipFill/a:srcRect), integer percent. */
  sourceRectangle?: SourceRectangleOptions;
}

/** Media player settings expressible by ODF draw:param. */
export interface MediaPlaybackOptions {
  loop?: boolean;
  mute?: boolean;
  /** Playback volume in decibels (0 = full scale). */
  volume?: number;
  /** ODF zoom setting: a percentage such as "100%" or "fit". */
  zoom?: string;
}

/** Play window trim (p14:trim @st/@end, seconds). */
export interface MediaTrimOptions {
  /** Trim start (p14:trim `@st`). */
  start?: number;
  /** Trim end (p14:trim `@end`). */
  end?: number;
}
