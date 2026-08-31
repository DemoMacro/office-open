import type { BaseGroupOptions } from "@office-open/core";
import type {
  BlackWhiteMode,
  EffectListOptions,
  FillOptions,
  GroupTransform2DOptions,
} from "@office-open/core/drawing";
import type { SlideChild } from "@parts/slide/slide-child";

/**
 * Group shape (p:grpSp) for pptx slides. cNvPr fields from BaseGroupOptions;
 * transform fields (incl. the child coordinate system, a:chOff/a:chExt) from
 * GroupTransform2DOptions — an unset chOff/chExt stringifies to the group
 * offset/extent. The rest is the group-level paint plus the children.
 */
export interface GroupOptions extends BaseGroupOptions, GroupTransform2DOptions {
  /** Group id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  /** Group-level fill (EG_FillProperties on grpSpPr). */
  fill?: FillOptions;
  /** Group-level effects (EG_EffectProperties on grpSpPr). */
  effects?: EffectListOptions;
  /** `@bwMode` (ST_BlackWhiteMode) on `p:grpSpPr` — how the group renders in black-and-white mode. */
  blackWhiteMode?: BlackWhiteMode;
  children: SlideChild[];
}
