import { pickNonVisualDrawingProperties } from "../drawing";
import type { NonVisualDrawingPropertiesOptions } from "../drawing";

/**
 * Base group options — the shared shape across pptx (p:grpSp), xlsx (xdr:grpSp),
 * and docx (wpg:wgp) groups.
 *
 * Carries only the non-visual drawing properties (name/description/title/hidden
 * that mirror a:CT_NonVisualDrawingProps) — the part of a group that round-trips
 * losslessly across all three formats. The group transform (a:xfrm), fill,
 * effects, and children stay package-side: positioning models differ (pptx flat
 * x/y/w/h, xlsx cell anchor, docx transformation) and children types are
 * fundamentally different (LegacySlideChild vs shapes+connectors vs
 * GroupChildMediaData), so they cannot be unified without distorting each
 * package's model. docx bridges its cNvPr through altText (DocPropertiesOptions),
 * mirroring the picture base pattern.
 */
export interface BaseGroupOptions extends NonVisualDrawingPropertiesOptions {
  /**
   * Reserved for group-specific shared fields (e.g. a:grpSpLocks) should a
   * future phase promote them across all three packages. Empty for now — see
   * the module docstring for why grpSpPr content and children stay per-package.
   */
}

/**
 * Pick the group base fields (cNvPr) actually set on `opts`, dropping undefined.
 * Used when bridging a package's group options onto the shared base during
 * cross-format conversion.
 */
export function pickGroupBase(opts: BaseGroupOptions | undefined): Partial<BaseGroupOptions> {
  if (!opts) return {};
  return { ...pickNonVisualDrawingProperties(opts) };
}
