import type { AnimationOptions } from "@file/animation/types";
import type { IMediaData } from "@file/media/data";

/**
 * Common options for all media frames.
 * @internal
 */
export interface MediaFrameBaseOptions {
  readonly id?: number;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly data: Uint8Array;
  readonly type: IMediaData["type"];
  readonly name?: string;
  readonly animation?: AnimationOptions;
}
