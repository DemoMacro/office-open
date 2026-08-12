/**
 * Cross-format shape conversion.
 *
 * Shape options convert between docx (wps), pptx (p:sp), and xlsx (xdr:sp).
 * The shape body — geometry/fill/outline/effects/3D — is already core
 * DrawingML in all three packages, so it round-trips near-losslessly. The
 * lossy legs are positioning (pptx/docx absolute EMU ↔ xlsx heuristic cell
 * anchors — see ./position) and text (docx w:p ↔ DrawingML a:p — see ./text).
 *
 * docx shapes carry text as w:p paragraphs; pptx/xlsx carry it as a core
 * TextBody (a:p). bodyProperties (a:bodyPr) round-trips verbatim. Geometry
 * adapts to docx's stricter API (presetGeometry rejects the bare-string
 * shorthand pptx/xlsx accept).
 *
 * @module
 */

import { pickNonVisualDrawingProperties } from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions } from "@office-open/core";
import type {
  FillOptions,
  OutlineOptions,
  EffectListOptions,
  EffectDagOptions,
  Scene3DOptions,
  Shape3DOptions,
  PresetGeometryOptions,
  ParagraphDescriptorOptions,
  TextBodyOptions,
  ShapePropertiesOptions,
} from "@office-open/core/drawingml";
import type {
  ShapeOptions as DocxShapeRunOptions,
  ShapeCoreOptions,
  MediaTransformation,
  ParagraphOptions as DocxParagraph,
} from "@office-open/docx";
import type { ShapeOptions as PptxShapeOptions } from "@office-open/pptx";
import type { ShapeOptions as XlsxShapeOptions } from "@office-open/xlsx";

import {
  boxFromPptx,
  boxFromXlsxAnchor,
  boxFromDocx,
  boxToPptx,
  boxToXlsx,
  boxToDocx,
} from "./position";
import { fromDrawingParagraph, toDrawingParagraph } from "./text";

/** docx shape input = ShapeOptions (core fields + transformation). */
export type DocxShapeOptions = DocxShapeRunOptions;

/** The five-plus shape-content fields shared verbatim across all three packages. */
export interface ShapeContent {
  fill?: FillOptions;
  outline?: OutlineOptions;
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
}

/** Copy the shared shape-content fields that are present on the source. */
export function pickContent<T extends ShapeContent>(source: T): ShapeContent {
  const out: ShapeContent = {};
  if (source.fill !== undefined) out.fill = source.fill;
  if (source.outline !== undefined) out.outline = source.outline;
  if (source.effects !== undefined) out.effects = source.effects;
  if (source.effectDag !== undefined) out.effectDag = source.effectDag;
  if (source.scene3d !== undefined) out.scene3d = source.scene3d;
  if (source.shape3d !== undefined) out.shape3d = source.shape3d;
  return out;
}

/** pptx/xlsx geometry shorthand (string | PresetGeometryOptions) → docx preset. */
export function toPresetGeometry(
  g: string | PresetGeometryOptions | undefined,
): PresetGeometryOptions | undefined {
  if (g === undefined) return undefined;
  return typeof g === "string" ? { preset: g } : g;
}

// ── text bridge ──

/** DrawingML text body (a:p) → docx w:p children. */
export function textBodyToDocxChildren(textBody: TextBodyOptions): DocxParagraph[] | string[] {
  const paragraphs = textBody.paragraphs ?? (textBody.text !== undefined ? [textBody.text] : []);
  const out: (DocxParagraph | string)[] = [];
  for (const p of paragraphs) {
    if (typeof p === "string") out.push(p);
    else out.push(fromDrawingParagraph(p));
  }
  return out as DocxParagraph[] | string[];
}

/** docx w:p children + bodyProperties → DrawingML text body (a:p), or undefined when empty. */
export function docxToTextBody(
  children: (DocxParagraph | string)[] | undefined,
  bodyProperties: TextBodyOptions["bodyProperties"],
): TextBodyOptions | undefined {
  const paragraphs: (ParagraphDescriptorOptions | string)[] = [];
  for (const c of children ?? []) {
    if (typeof c === "string") paragraphs.push(c);
    else paragraphs.push(toDrawingParagraph(c));
  }
  if (paragraphs.length === 0 && bodyProperties === undefined) return undefined;
  const out: TextBodyOptions = {};
  if (paragraphs.length > 0) out.paragraphs = paragraphs;
  if (bodyProperties !== undefined) out.bodyProperties = bodyProperties;
  return out;
}

// ── → docx ──

/** docx shape split: the wps core (data) + position (transformation). */
export interface DocxShapeParts {
  data: ShapeCoreOptions;
  transformation: MediaTransformation;
}

/**
 * Build docx nonVisualProperties from a source's cNvPr — all authored fields,
 * not just name. name defaults to "Shape" (docx requires it).
 */
const docxNonVisual = (
  source: NonVisualDrawingPropertiesOptions,
): { nonVisualProperties: NonVisualDrawingPropertiesOptions } => {
  const picked = pickNonVisualDrawingProperties(source);
  return { nonVisualProperties: { name: picked.name ?? "Shape", ...picked } };
};

/** True when a source carries at least one authored cNvPr field. */
const hasCnvPr = (source: NonVisualDrawingPropertiesOptions): boolean => {
  const picked = pickNonVisualDrawingProperties(source);
  return (
    picked.name !== undefined ||
    picked.description !== undefined ||
    picked.title !== undefined ||
    picked.hidden !== undefined
  );
};

/**
 * Build the docx wps core + position from a pptx or xlsx shape. Group
 * conversion reuses this to embed shapes as wpg children (the child carries a
 * full MediaDataTransformation; the caller runs it through createTransformation).
 */
export function toDocxShapeParts(source: PptxShapeOptions | XlsxShapeOptions): DocxShapeParts {
  if ("spPr" in source) {
    // xlsx → docx
    const spPr = source.spPr;
    const box = boxFromXlsxAnchor(
      source,
      spPr.width,
      spPr.height,
      spPr.rotation,
      spPr.flipHorizontal,
      spPr.flipVertical,
    );
    const preset = toPresetGeometry(spPr.geometry);
    return {
      data: {
        children: source.textBody ? textBodyToDocxChildren(source.textBody) : [],
        ...pickContent(spPr),
        ...(spPr.customGeometry !== undefined ? { customGeometry: spPr.customGeometry } : {}),
        ...(preset !== undefined ? { presetGeometry: preset } : {}),
        ...(hasCnvPr(source) ? docxNonVisual(source) : {}),
      },
      transformation: boxToDocx(box),
    };
  }
  // pptx → docx
  const box = boxFromPptx(
    source.x,
    source.y,
    source.width,
    source.height,
    source.rotation,
    source.flipHorizontal,
  );
  const preset = toPresetGeometry(source.geometry);
  return {
    data: {
      children: source.textBody ? textBodyToDocxChildren(source.textBody) : [],
      ...pickContent(source),
      ...(source.customGeometry !== undefined ? { customGeometry: source.customGeometry } : {}),
      ...(preset !== undefined ? { presetGeometry: preset } : {}),
      ...(hasCnvPr(source) ? docxNonVisual(source) : {}),
    },
    transformation: boxToDocx(box),
  };
}

/** Convert a pptx shape to a docx wps shape. */
export function toDocxShape(source: PptxShapeOptions): DocxShapeOptions;
/** Convert an xlsx shape to a docx wps shape. */
export function toDocxShape(source: XlsxShapeOptions): DocxShapeOptions;
export function toDocxShape(source: PptxShapeOptions | XlsxShapeOptions): DocxShapeOptions {
  const { data, transformation } = toDocxShapeParts(source);
  return { ...data, transformation };
}

// ── → pptx ──

/** Convert a docx wps shape to a pptx shape. */
export function toPptxShape(source: DocxShapeOptions): PptxShapeOptions;
/** Convert an xlsx shape to a pptx shape. */
export function toPptxShape(source: XlsxShapeOptions): PptxShapeOptions;
export function toPptxShape(source: DocxShapeOptions | XlsxShapeOptions): PptxShapeOptions {
  if ("spPr" in source) {
    // xlsx → pptx
    const spPr = source.spPr;
    const box = boxFromXlsxAnchor(
      source,
      spPr.width,
      spPr.height,
      spPr.rotation,
      spPr.flipHorizontal,
      spPr.flipVertical,
    );
    const result: PptxShapeOptions = {
      ...boxToPptx(box),
      ...pickContent(spPr),
      ...(spPr.geometry !== undefined ? { geometry: spPr.geometry } : {}),
      ...(spPr.customGeometry !== undefined ? { customGeometry: spPr.customGeometry } : {}),
      ...(source.textBody ? { textBody: source.textBody } : {}),
      ...pickNonVisualDrawingProperties(source),
    };
    return result;
  }
  // docx → pptx
  const box = boxFromDocx(source.transformation);
  const textBody = docxToTextBody(source.children, source.bodyProperties);
  const result: PptxShapeOptions = {
    ...boxToPptx(box),
    ...pickContent(source),
    ...(source.presetGeometry !== undefined
      ? { geometry: source.presetGeometry }
      : source.customGeometry !== undefined
        ? { customGeometry: source.customGeometry }
        : {}),
    ...(textBody ? { textBody } : {}),
    ...pickNonVisualDrawingProperties(source.nonVisualProperties),
  };
  return result;
}

// ── → xlsx ──

/** Convert a docx wps shape to an xlsx shape. */
export function toXlsxShape(source: DocxShapeOptions): XlsxShapeOptions;
/** Convert a pptx shape to an xlsx shape. */
export function toXlsxShape(source: PptxShapeOptions): XlsxShapeOptions;
export function toXlsxShape(source: DocxShapeOptions | PptxShapeOptions): XlsxShapeOptions {
  if ("transformation" in source) {
    // docx → xlsx
    const box = boxFromDocx(source.transformation);
    const pos = boxToXlsx(box);
    const textBody = docxToTextBody(source.children, source.bodyProperties);
    const spPr: ShapePropertiesOptions = {
      x: pos.xfrmX,
      y: pos.xfrmY,
      width: box.width,
      height: box.height,
      ...pickContent(source),
      ...(source.presetGeometry !== undefined
        ? { geometry: source.presetGeometry }
        : source.customGeometry !== undefined
          ? { customGeometry: source.customGeometry }
          : {}),
      ...(box.rotation !== undefined ? { rotation: box.rotation } : {}),
      ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
      ...(box.flipVertical ? { flipVertical: true } : {}),
    };
    return {
      ...pos.anchor,
      spPr,
      ...(textBody ? { textBody } : {}),
      ...pickNonVisualDrawingProperties(source.nonVisualProperties),
    };
  }
  // pptx → xlsx
  const box = boxFromPptx(
    source.x,
    source.y,
    source.width,
    source.height,
    source.rotation,
    source.flipHorizontal,
  );
  const pos = boxToXlsx(box);
  const spPr: ShapePropertiesOptions = {
    x: pos.xfrmX,
    y: pos.xfrmY,
    width: box.width,
    height: box.height,
    ...pickContent(source),
    ...(source.geometry !== undefined ? { geometry: source.geometry } : {}),
    ...(source.customGeometry !== undefined ? { customGeometry: source.customGeometry } : {}),
    ...(box.rotation !== undefined ? { rotation: box.rotation } : {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
  };
  return {
    ...pos.anchor,
    spPr,
    ...(source.textBody ? { textBody: source.textBody } : {}),
    ...pickNonVisualDrawingProperties(source),
  };
}
