/**
 * Shape properties descriptor for DrawingML CT_ShapeProperties.
 *
 * Assembles the XSD-ordered children of a shape/picture/group properties
 * element (a:xfrm → geometry → fill → a:ln → effects → a:scene3d → a:sp3d).
 * Returns the inner content only; the caller wraps the container tag
 * (p:spPr / wps:spPr / pic:spPr / xdr:spPr) and any container-only attributes
 * such as bwMode.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ShapeProperties / EG_ShapeProperties
 *
 * @module
 */

import { escapeXml, findChild, stringifyElement } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import { parse } from "../descriptor";
import type { UniversalMeasure } from "../util/values";
import { createEffectDag } from "./effects/effect-dag";
import type { EffectDagOptions } from "./effects/effect-dag";
import { effectListDesc } from "./effects/effect-descriptors";
import type { EffectListOptions } from "./effects/effect-list";
import { fillDesc, findFillChild } from "./fill/fill-descriptors";
import type { FillOptions } from "./fill/fill-options";
import type { CustomGeometryOptions } from "./geometry/custom-geometry";
import { presetGeometryDesc, customGeometryDesc } from "./geometry/geometry-descriptors";
import type { PresetGeometryOptions } from "./geometry/preset-geometry";
import type { OutlineOptions } from "./outline/outline";
import { outlineDesc, stringifyLineProperties } from "./outline/outline-descriptors";
import type { Scene3DOptions } from "./three-d/scene-3d";
import type { Shape3DOptions } from "./three-d/shape-3d";
import { scene3DDesc, shape3DDesc } from "./three-d/three-d-descriptors";
import type { Transform2DOptions, GroupTransform2DOptions } from "./transform";
import { transform2DDesc, groupTransform2DDesc } from "./transform-descriptors";

// ── Types ──

/** A shape-property extension (a:ext). Known Office extensions are structured;
 * unknown payloads remain verbatim in `content`. */
export interface ShapePropertiesExtensionOptions {
  uri: string;
  hiddenLine?: OutlineOptions;
  content?: string;
}

/**
 * Shape properties (CT_ShapeProperties children). Field order follows the XSD
 * content model; the serializer emits in that order. Container-only attributes
 * (bwMode, rotWithShape) and the wrapping tag are the caller's responsibility.
 */
export interface ShapePropertiesOptions {
  // CT_Transform2D (a:xfrm)
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  /** Rotation angle in degrees (e.g., 45 = 45°). */
  rotation?: number;
  // CT_GroupTransform2D extras (a:chOff / a:chExt). When any is set, the
  // descriptor emits a group transform (groupTransform2DDesc) instead of a
  // plain CT_Transform2D.
  childOffsetX?: number | UniversalMeasure;
  childOffsetY?: number | UniversalMeasure;
  childExtentWidth?: number | UniversalMeasure;
  childExtentHeight?: number | UniversalMeasure;
  // EG_Geometry (choice: a:custGeom | a:prstGeom). customGeometry wins; a bare
  // string geometry is shorthand for { preset: "<name>" }.
  geometry?: string | PresetGeometryOptions;
  customGeometry?: CustomGeometryOptions;
  // EG_FillProperties
  fill?: FillOptions;
  // a:ln
  outline?: OutlineOptions;
  // EG_EffectProperties (choice: a:effectDag | a:effectLst)
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
  // 3D
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
  /** Shape-property extensions (a:extLst/a:ext). */
  extensions?: ShapePropertiesExtensionOptions[];
  /**
   * Raw a:extLst inner XML. Prefer `extensions` for structured shape-property
   * extensions; retained for existing callers and unmodeled payloads.
   */
  ext?: string;
}

// ── Descriptor ──

export const shapePropertiesDesc: CustomDescriptor<ShapePropertiesOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // a:xfrm — emit only when any transform field is set. Group-only child
    // fields (chOff/chExt) trigger a group transform.
    const hasGroupChild =
      opts.childOffsetX !== undefined ||
      opts.childOffsetY !== undefined ||
      opts.childExtentWidth !== undefined ||
      opts.childExtentHeight !== undefined;
    const hasTransform =
      opts.x !== undefined ||
      opts.y !== undefined ||
      opts.width !== undefined ||
      opts.height !== undefined ||
      opts.flipHorizontal !== undefined ||
      opts.flipVertical !== undefined ||
      opts.rotation !== undefined ||
      hasGroupChild;
    if (hasTransform) {
      const transformOpts = {
        x: opts.x,
        y: opts.y,
        width: opts.width,
        height: opts.height,
        flipHorizontal: opts.flipHorizontal,
        flipVertical: opts.flipVertical,
        rotation: opts.rotation,
        childOffsetX: opts.childOffsetX,
        childOffsetY: opts.childOffsetY,
        childExtentWidth: opts.childExtentWidth,
        childExtentHeight: opts.childExtentHeight,
      };
      const xfrm = hasGroupChild
        ? groupTransform2DDesc.stringify(transformOpts as GroupTransform2DOptions, ctx)
        : transform2DDesc.stringify(transformOpts as Transform2DOptions, ctx);
      if (xfrm) parts.push(xfrm);
    }

    // EG_Geometry: customGeometry takes precedence (matches pptx/docx).
    if (opts.customGeometry) {
      const g = customGeometryDesc.stringify(opts.customGeometry, ctx);
      if (g) parts.push(g);
    } else if (opts.geometry !== undefined) {
      const geom: PresetGeometryOptions =
        typeof opts.geometry === "string" ? { preset: opts.geometry } : opts.geometry;
      const g = presetGeometryDesc.stringify(geom, ctx);
      if (g) parts.push(g);
    }

    // EG_FillProperties
    if (opts.fill !== undefined) {
      const f = fillDesc.stringify(opts.fill, ctx);
      if (f) parts.push(f);
    }

    // a:ln
    if (opts.outline) {
      const ln = outlineDesc.stringify(opts.outline, ctx);
      if (ln) parts.push(ln);
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

    // a:sp3d
    if (opts.shape3d) {
      const sp = shape3DDesc.stringify(opts.shape3d, ctx);
      if (sp) parts.push(sp);
    }

    // a:extLst — last child per CT_ShapeProperties sequence. Structured
    // extensions take precedence over the legacy raw inner XML field.
    if (opts.extensions !== undefined) {
      const extensions = opts.extensions
        .map((extension) => {
          let content = extension.content ?? "";
          if (extension.hiddenLine) {
            const hiddenLine =
              stringifyLineProperties("a14:hiddenLine", extension.hiddenLine, ctx) ?? "";
            content += hiddenLine.replace(
              "<a14:hiddenLine",
              '<a14:hiddenLine xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"',
            );
          }
          return `<a:ext uri="${escapeXml(extension.uri)}">${content}</a:ext>`;
        })
        .join("");
      parts.push(`<a:extLst>${extensions}</a:extLst>`);
    } else if (opts.ext !== undefined) {
      parts.push(`<a:extLst>${opts.ext}</a:extLst>`);
    }

    return parts.length > 0 ? parts.join("") : undefined;
  },

  parse(el, ctx) {
    const result: Partial<ShapePropertiesOptions> = {};

    // a:xfrm — group transform when chOff/chExt present, else plain transform.
    const xfrm = findChild(el, "a:xfrm");
    if (xfrm) {
      const hasGroupChild = !!(findChild(xfrm, "a:chOff") ?? findChild(xfrm, "a:chExt"));
      Object.assign(
        result,
        hasGroupChild ? groupTransform2DDesc.parse(xfrm, ctx) : transform2DDesc.parse(xfrm, ctx),
      );
    }

    // EG_Geometry
    const custGeom = findChild(el, "a:custGeom");
    if (custGeom) {
      result.customGeometry = customGeometryDesc.parse(custGeom, ctx);
    } else {
      const prstGeom = findChild(el, "a:prstGeom");
      if (prstGeom) result.geometry = presetGeometryDesc.parse(prstGeom, ctx);
    }

    // EG_FillProperties — fillDesc always returns a value (defaults to none),
    // so probe for a fill child first to avoid synthesizing a spurious fill.
    const fillChild = findFillChild(el);
    if (fillChild) {
      result.fill = parse(fillDesc, fillChild, ctx);
    }

    // a:ln
    const ln = findChild(el, "a:ln");
    if (ln) result.outline = outlineDesc.parse(ln, ctx);

    // EG_EffectProperties — effectLst. (effectDag parse is not modeled here;
    // matches current per-package behavior. effectDag still round-trips via
    // stringify when the caller supplies it.)
    const effectLst = findChild(el, "a:effectLst");
    if (effectLst) result.effects = effectListDesc.parse(effectLst, ctx);

    // a:scene3d
    const scene3d = findChild(el, "a:scene3d");
    if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, ctx);

    // a:sp3d
    const sp3d = findChild(el, "a:sp3d");
    if (sp3d) result.shape3d = shape3DDesc.parse(sp3d, ctx);

    // a:extLst — structure known Office extensions. Lists with no known
    // payload stay on the legacy raw field so existing public APIs keep their
    // established shape; mixed lists retain unknown payloads per extension.
    const extLst = findChild(el, "a:extLst");
    if (extLst) {
      const extensions = (extLst.elements ?? [])
        .filter((child) => child.type === "element" && child.name === "a:ext")
        .map((extension) => {
          const parsed: ShapePropertiesExtensionOptions = {
            uri: String(extension.attributes?.["uri"] ?? ""),
          };
          const unknown: string[] = [];
          for (const child of extension.elements ?? []) {
            if (child.type === "element" && child.name === "a14:hiddenLine") {
              parsed.hiddenLine = outlineDesc.parse(child, ctx);
            } else {
              unknown.push(stringifyElement(child));
            }
          }
          if (unknown.length > 0) parsed.content = unknown.join("");
          return parsed;
        });
      if (extensions.some((extension) => extension.hiddenLine !== undefined)) {
        result.extensions = extensions;
      } else {
        result.ext = (extLst.elements ?? []).map(stringifyElement).join("");
      }
    }

    return result as ShapePropertiesOptions;
  },
};
