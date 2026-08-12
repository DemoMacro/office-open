/**
 * Group shape properties descriptor for DrawingML CT_GroupShapeProperties.
 *
 * Assembles the XSD-ordered children of a group's properties element
 * (a:xfrm → EG_FillProperties → EG_EffectProperties → a:scene3d). Returns the
 * inner content only; the caller wraps the container tag (p:grpSpPr /
 * wpg:grpSpPr / xdr:grpSpPr / a:grpSpPr) and any container-only attributes
 * such as bwMode.
 *
 * Unlike CT_ShapeProperties, a group carries no geometry (prstGeom/custGeom),
 * no outline (a:ln), and no shape 3D (a:sp3d) — only the transform, fill,
 * effects, and scene 3D that apply to the group as a whole. The transform is
 * always a group transform (CT_GroupTransform2D), so the fields flatten in
 * directly via `extends GroupTransform2DOptions` (mirroring the flat style of
 * ShapePropertiesOptions, without its geometry/outline/shape3d extras).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GroupShapeProperties
 *
 * @module
 */

import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import { parse } from "../descriptor";
import { createEffectDag } from "./effects/effect-dag";
import type { EffectDagOptions } from "./effects/effect-dag";
import { effectListDesc } from "./effects/effect-descriptors";
import type { EffectListOptions } from "./effects/effect-list";
import { fillDesc } from "./fill/fill-descriptors";
import type { FillOptions } from "./fill/fill-options";
import type { Scene3DOptions } from "./three-d/scene-3d";
import { scene3DDesc } from "./three-d/three-d-descriptors";
import type { GroupTransform2DOptions } from "./transform";
import { groupTransform2DDesc } from "./transform-descriptors";

// ── Types ──

/**
 * Group shape properties (CT_GroupShapeProperties). Field order follows the XSD
 * content model. The transform fields come from {@link GroupTransform2DOptions}
 * (flat x/y/width/height/flip/rotation/childOffset/childExtent); the serializer
 * emits them as a single a:xfrm group transform. The container-only @bwMode
 * attribute and the wrapping tag are the caller's responsibility (mirrors
 * shapePropertiesDesc).
 */
export interface GroupShapePropertiesOptions extends GroupTransform2DOptions {
  /** EG_FillProperties. */
  fill?: FillOptions;
  /** EG_EffectProperties (choice: a:effectDag | a:effectLst). */
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
  /** a:scene3d. */
  scene3d?: Scene3DOptions;
}

// ── Descriptor ──

export const groupShapePropertiesDesc: CustomDescriptor<GroupShapePropertiesOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // a:xfrm — group transform (CT_GroupTransform2D). Emit when any transform
    // field is set. A group always uses a group transform (no plain/sp transform).
    const hasTransform =
      opts.x !== undefined ||
      opts.y !== undefined ||
      opts.width !== undefined ||
      opts.height !== undefined ||
      opts.flipHorizontal !== undefined ||
      opts.flipVertical !== undefined ||
      opts.rotation !== undefined ||
      opts.childOffsetX !== undefined ||
      opts.childOffsetY !== undefined ||
      opts.childExtentWidth !== undefined ||
      opts.childExtentHeight !== undefined;
    if (hasTransform) {
      const xfrm = groupTransform2DDesc.stringify(opts, ctx);
      if (xfrm) parts.push(xfrm);
    }

    // EG_FillProperties
    if (opts.fill !== undefined) {
      const f = fillDesc.stringify(opts.fill, ctx);
      if (f) parts.push(f);
    }

    // EG_EffectProperties: effectDag | effectLst (mutually exclusive).
    if (opts.effectDag) {
      parts.push(createEffectDag(opts.effectDag));
    } else if (opts.effects) {
      const e = effectListDesc.stringify(opts.effects, ctx);
      if (e) parts.push(e);
    }

    // a:scene3d
    if (opts.scene3d) {
      const s = scene3DDesc.stringify(opts.scene3d, ctx);
      if (s) parts.push(s);
    }

    return parts.length > 0 ? parts.join("") : undefined;
  },

  parse(el, ctx) {
    const result: Partial<GroupShapePropertiesOptions> = {};

    // a:xfrm — group transform.
    const xfrm = findChild(el, "a:xfrm");
    if (xfrm) {
      Object.assign(result, groupTransform2DDesc.parse(xfrm, ctx));
    }

    // EG_FillProperties — fillDesc always returns a value (defaults to none),
    // so probe for a fill child first to avoid synthesizing a spurious fill.
    const fillChild =
      findChild(el, "a:noFill") ??
      findChild(el, "a:solidFill") ??
      findChild(el, "a:gradFill") ??
      findChild(el, "a:pattFill") ??
      findChild(el, "a:grpFill") ??
      findChild(el, "a:blipFill");
    if (fillChild) {
      result.fill = parse(fillDesc, el, ctx);
    }

    // EG_EffectProperties — effectLst. (effectDag parse is not modeled here;
    // matches shapePropertiesDesc behavior. effectDag still round-trips via
    // stringify when the caller supplies it.)
    const effectLst = findChild(el, "a:effectLst");
    if (effectLst) result.effects = effectListDesc.parse(effectLst, ctx);

    // a:scene3d
    const scene3d = findChild(el, "a:scene3d");
    if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, ctx);

    return result as GroupShapePropertiesOptions;
  },
};
