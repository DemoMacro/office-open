import type { BlackWhiteMode, EffectListOptions } from "@office-open/core/drawing";
import type { FillOptions } from "@shared/drawing/fill";

export interface BackgroundOptions {
  fill?: FillOptions;
  effects?: EffectListOptions;
  shadeToTitle?: boolean;
  /** `@bwMode` (ST_BlackWhiteMode) — unprefixed attribute on `p:bg`. */
  blackWhiteMode?: BlackWhiteMode;
}
