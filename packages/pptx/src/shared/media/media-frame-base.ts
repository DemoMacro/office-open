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
   * Click-to-play hyperlink (a:hlinkClick `action="ppaction://media"` on
   * p:cNvPr) — Office emits it on media frames.
   */
  mediaAction?: boolean;
}
