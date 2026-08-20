/**
 * Cross-format picture conversion.
 *
 * Each package's PictureOptions extends (or, for docx, is bridged onto) the
 * core BasePictureOptions: the binary payload (data/type) plus the non-visual
 * drawing properties (name/description/title/hidden) that mirror
 * a:CT_NonVisualDrawingProps. Those cNvPr fields pass straight through every
 * conversion leg via pickNonVisualDrawingProperties, so alt text survives a
 * cross-format copy instead of being dropped.
 *
 * Position mapping is heuristic where the target has no matching coordinate
 * model: pptx/docx use absolute EMU coordinates; xlsx uses 1-based cell anchors
 * with no size on the public input. EMU↔cell converts via a default cell size
 * (8.43-char column × 15pt row); size is lost on the xlsx leg. This matches MS
 * Office paste behavior between apps.
 *
 * docx is the odd one out: its PictureOptions is a format discriminated union
 * (regular raster vs. SVG-with-fallback), so it does not extend BasePictureOptions.
 * The cNvPr fields live on its structured altText field instead, and SVG falls
 * back to its raster payload when targeting pptx/xlsx (no vector support).
 *
 * @module
 */
import { pickNonVisualDrawingProperties } from "@office-open/core";
import type { BasePictureOptions, NonVisualDrawingPropertiesOptions } from "@office-open/core";
import type { PictureOptions as DocxPictureOptions } from "@office-open/docx";
import type { PictureOptions as PptxPictureOptions } from "@office-open/pptx";
import type { PictureOptions as XlsxPictureOptions } from "@office-open/xlsx";

import { DEFAULT_COL_EMU, DEFAULT_ROW_EMU, emuToCell, toEmu } from "./position";

// ── base readers: each package → shared BasePictureOptions ──

/** Project a pptx picture onto the shared base (data/type + cNvPr). */
const baseFromPptx = (p: PptxPictureOptions): BasePictureOptions => ({
  data: p.data,
  type: p.type,
  ...(p.sourceUrl !== undefined ? { sourceUrl: p.sourceUrl } : {}),
  ...pickNonVisualDrawingProperties(p),
});

/** Project an xlsx picture onto the shared base. */
const baseFromXlsx = (x: XlsxPictureOptions): BasePictureOptions => ({
  data: x.data,
  type: x.type,
  ...(x.sourceUrl !== undefined ? { sourceUrl: x.sourceUrl } : {}),
  ...pickNonVisualDrawingProperties(x),
});

/**
 * Project a docx picture onto the shared base. docx does not extend
 * BasePictureOptions (its PictureOptions is a format discriminated union), so
 * the cNvPr fields are read from the structured altText. SVG falls back to its
 * raster payload since pptx/xlsx have no vector picture support.
 */
const baseFromDocx = (d: DocxPictureOptions): BasePictureOptions => {
  const cNvPr = pickNonVisualDrawingProperties(d.altText);
  if (d.type === "svg") {
    return { data: d.fallback.data, type: d.fallback.type, ...cNvPr };
  }
  return {
    data: d.data,
    type: d.type,
    ...(d.sourceUrl !== undefined ? { sourceUrl: d.sourceUrl } : {}),
    ...cNvPr,
  };
};

// ── type narrowing ──

type DocxRasterType = "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
const DOCX_RASTER_TYPES: readonly DocxRasterType[] = [
  "jpg",
  "png",
  "gif",
  "bmp",
  "tif",
  "ico",
  "emf",
  "wmf",
];

/** Narrow an image type to docx's raster set (pptx/xlsx sources are never svg). */
const docxType = (type: string): DocxRasterType =>
  (DOCX_RASTER_TYPES as readonly string[]).includes(type) ? (type as DocxRasterType) : "png";

const PPTX_TYPES = ["png", "jpg", "gif", "bmp", "emf", "wmf"] as const;
/** Narrow an image type to pptx's supported set, falling back to png. */
const pptxType = (type: string): PptxPictureOptions["type"] =>
  (PPTX_TYPES as readonly string[]).includes(type) ? (type as PptxPictureOptions["type"]) : "png";

/** Narrow an image type to xlsx's png/jpg set. */
const xlsxType = (type: string): "png" | "jpg" =>
  type === "jpg" || type === "jpeg" ? "jpg" : "png";

/**
 * Build the docx altText (wp:docPr) from the shared base. Only emitted when at
 * least one cNvPr field is authored; name defaults to "Picture" since docx
 * requires it. Structurally compatible with docx's DocPropertiesOptions without
 * importing that internal type.
 */
const altTextFromBase = (
  base: BasePictureOptions,
): { altText?: NonVisualDrawingPropertiesOptions & { name: string } } => {
  const picked = pickNonVisualDrawingProperties(base);
  if (
    picked.name === undefined &&
    picked.description === undefined &&
    picked.title === undefined &&
    picked.hidden === undefined
  ) {
    return {};
  }
  return { altText: { name: picked.name ?? "Picture", ...picked } };
};

// ── → docx ──

/** Convert a pptx picture to a docx inline image. */
export function toDocxPicture(source: PptxPictureOptions): DocxPictureOptions;
/** Convert an xlsx image to a docx inline image (size defaults to 0; xlsx carries no size). */
export function toDocxPicture(source: XlsxPictureOptions): DocxPictureOptions;
export function toDocxPicture(source: PptxPictureOptions | XlsxPictureOptions): DocxPictureOptions {
  // pptx → docx: absolute x/y → offset, width/height → transformation.
  if ("width" in source || "height" in source) {
    const p = source as PptxPictureOptions;
    const base = baseFromPptx(p);
    return {
      type: docxType(base.type),
      data: base.data,
      ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
      transformation: {
        width: p.width ?? 0,
        height: p.height ?? 0,
        ...(p.x !== undefined || p.y !== undefined
          ? { offset: { left: p.x ?? 0, top: p.y ?? 0 } }
          : {}),
      },
      ...altTextFromBase(base),
    };
  }
  // xlsx → docx: cell anchor → offset EMU; size unknown.
  const x = source as XlsxPictureOptions;
  const base = baseFromXlsx(x);
  return {
    type: docxType(base.type),
    data: base.data,
    ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
    transformation: {
      width: 0,
      height: 0,
      offset: { left: (x.col - 1) * DEFAULT_COL_EMU, top: (x.row - 1) * DEFAULT_ROW_EMU },
    },
    ...altTextFromBase(base),
  };
}

// ── → pptx ──

/** Convert a docx image to a pptx picture. */
export function toPptxPicture(source: DocxPictureOptions): PptxPictureOptions;
/** Convert an xlsx image to a pptx picture (size defaults to 0; xlsx carries no size). */
export function toPptxPicture(source: XlsxPictureOptions): PptxPictureOptions;
export function toPptxPicture(source: DocxPictureOptions | XlsxPictureOptions): PptxPictureOptions {
  // docx → pptx: transformation → absolute x/y + width/height.
  if ("transformation" in source) {
    const d = source as DocxPictureOptions;
    const base = baseFromDocx(d);
    const t = d.transformation;
    return {
      type: pptxType(base.type),
      data: base.data,
      ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
      width: t.width,
      height: t.height,
      ...(t.offset ? { x: t.offset.left, y: t.offset.top } : {}),
      ...pickNonVisualDrawingProperties(base),
    };
  }
  // xlsx → pptx: cell anchor → absolute EMU; size unknown.
  const x = source as XlsxPictureOptions;
  const base = baseFromXlsx(x);
  return {
    type: pptxType(base.type),
    data: base.data,
    ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
    x: (x.col - 1) * DEFAULT_COL_EMU,
    y: (x.row - 1) * DEFAULT_ROW_EMU,
    width: 0,
    height: 0,
    ...pickNonVisualDrawingProperties(base),
  };
}

// ── → xlsx ──

/** Convert a docx image to an xlsx picture (position mapped to cell anchor; size lost). */
export function toXlsxPicture(source: DocxPictureOptions): XlsxPictureOptions;
/** Convert a pptx picture to an xlsx picture (position mapped to cell anchor; size lost). */
export function toXlsxPicture(source: PptxPictureOptions): XlsxPictureOptions;
export function toXlsxPicture(source: DocxPictureOptions | PptxPictureOptions): XlsxPictureOptions {
  // docx → xlsx: offset EMU → cell anchor.
  if ("transformation" in source) {
    const d = source as DocxPictureOptions;
    const base = baseFromDocx(d);
    const left = toEmu(d.transformation.offset?.left);
    const top = toEmu(d.transformation.offset?.top);
    return {
      data: base.data,
      type: xlsxType(base.type),
      ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
      col: emuToCell(left, DEFAULT_COL_EMU),
      row: emuToCell(top, DEFAULT_ROW_EMU),
      ...pickNonVisualDrawingProperties(base),
    };
  }
  // pptx → xlsx: absolute EMU → cell anchor.
  const p = source as PptxPictureOptions;
  const base = baseFromPptx(p);
  return {
    data: base.data,
    type: xlsxType(base.type),
    ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
    col: emuToCell(toEmu(p.x), DEFAULT_COL_EMU),
    row: emuToCell(toEmu(p.y), DEFAULT_ROW_EMU),
    ...pickNonVisualDrawingProperties(base),
  };
}
