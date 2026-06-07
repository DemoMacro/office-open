import type { AnimationOptions } from "@file/animation/types";
import type { IMediaData } from "@file/media/data";

/**
 * Common options for all media frames.
 * @internal
 */
export interface MediaFrameBaseOptions {
  id?: number;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  data: Uint8Array;
  type: IMediaData["type"];
  name?: string;
  animation?: AnimationOptions;
}
