import type { BaseGroupOptions, UniversalMeasure } from "@office-open/core";
import type { BlackWhiteMode, EffectListOptions, FillOptions } from "@office-open/core/drawing";
import type { SlideChild } from "@parts/slide/slide-child";

/**
 * Group shape options for pptx slides (p:grpSp). The cNvPr fields
 * (name/description/title/hidden) come from {@link BaseGroupOptions}; the rest
 * is the pptx flat positioning model plus the group's children. The single
 * source of truth for both the public slide-child entry and the descriptor.
 */
export interface GroupOptions extends BaseGroupOptions {
  /** Group id (p:cNvPr @id). Auto-generated if omitted. */
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Rotation angle in degrees (e.g., 45 = 45°). */
  rotation?: number;
  flipHorizontal?: boolean;
  /** Child coordinate system offset (a:chOff). Defaults to {x,y} when omitted. */
  childOffset?: { x: number | UniversalMeasure; y: number | UniversalMeasure };
  /** Child coordinate system extent (a:chExt). Defaults to {width,height} when omitted. */
  childExtent?: { cx: number | UniversalMeasure; cy: number | UniversalMeasure };
  /** Group-level fill (EG_FillProperties on grpSpPr). */
  fill?: FillOptions;
  /** Group-level effects (EG_EffectProperties on grpSpPr). */
  effects?: EffectListOptions;
  /** @bwMode container attribute (ST_BlackWhiteMode) on p:grpSpPr. */
  bwMode?: BlackWhiteMode;
  children: SlideChild[];
}
