import type { DataType, UniversalMeasure } from "@office-open/core";
import type { AnimationOptions } from "@shared/animation/types";
import type { MediaData } from "@shared/media/data";

/**
 * Common options for all media frames.
 * @internal
 */
export interface MediaFrameBaseOptions {
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  data: DataType;
  type: MediaData["type"];
  name?: string;
  animation?: AnimationOptions;
}
