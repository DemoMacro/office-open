import type {
  DataType,
  NonVisualDrawingPropertiesOptions,
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
  /**
   * Play window trim of the p14:media extension copy (p14:trim, seconds).
   * undefined = no trim child.
   */
  trim?: MediaTrimOptions;
}

/** Play window trim (p14:trim @st/@end, seconds). */
export interface MediaTrimOptions {
  /** Trim start (p14:trim `@st`). */
  start?: number;
  /** Trim end (p14:trim `@end`). */
  end?: number;
}
