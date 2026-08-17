import type { BlackWhiteMode, EffectListOptions } from "@office-open/core/drawing";
import type { StyleMatrixReferenceOptions } from "@office-open/core/drawing";
import type { FillOptions } from "@shared/drawing/fill";

export type { StyleMatrixReferenceOptions };

export interface BackgroundOptions {
  /** Style matrix reference (p:bgRef) — the common "inherit theme background style" form; mutually exclusive with fill. */
  reference?: StyleMatrixReferenceOptions;
  fill?: FillOptions;
  effects?: EffectListOptions;
  shadeToTitle?: boolean;
  /** `@bwMode` (ST_BlackWhiteMode) — unprefixed attribute on `p:bg`. */
  blackWhiteMode?: BlackWhiteMode;
}

/** MS Office default master/notes background — the first theme background style. */
export const DEFAULT_BACKGROUND_REFERENCE: StyleMatrixReferenceOptions = {
  index: 1001,
  color: { value: "bg1" },
};
