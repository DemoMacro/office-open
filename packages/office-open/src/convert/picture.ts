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
import type {
  BasePictureOptions,
  NonVisualDrawingPropertiesOptions,
  PictureLockingOptions,
} from "@office-open/core";
import type { BlipCompression, BlackWhiteMode } from "@office-open/core";
import type { BlipEffectsOptions, SourceRectangleOptions } from "@office-open/core/drawing";
import type { PictureOptions as DocxPictureOptions } from "@office-open/docx";
import type { PictureOptions as PptxPictureOptions } from "@office-open/pptx";
import type { PictureOptions as XlsxPictureOptions } from "@office-open/xlsx";

import {
  boxFromDocx,
  boxToPptx,
  boxToDocxTransformation,
  DEFAULT_COL_EMU,
  DEFAULT_ROW_EMU,
  emuToCell,
  toEmu,
} from "./position";

interface PictureAppearanceOptions {
  sourceRectangle?: SourceRectangleOptions;
  blipEffects?: BlipEffectsOptions;
  compression?: BlipCompression;
  blackWhiteMode?: BlackWhiteMode;
  rotWithShape?: boolean;
  blipExt?: string;
}

function pickPictureAppearance(source: PictureAppearanceOptions): PictureAppearanceOptions {
  return {
    ...(source.sourceRectangle !== undefined ? { sourceRectangle: source.sourceRectangle } : {}),
    ...(source.blipEffects !== undefined ? { blipEffects: source.blipEffects } : {}),
    ...(source.compression !== undefined ? { compression: source.compression } : {}),
    ...(source.blackWhiteMode !== undefined ? { blackWhiteMode: source.blackWhiteMode } : {}),
    ...(source.rotWithShape !== undefined ? { rotWithShape: source.rotWithShape } : {}),
    ...(source.blipExt !== undefined ? { blipExt: source.blipExt } : {}),
  };
}

function pickPptxXlsxPictureOptions(source: {
  locking?: PictureLockingOptions;
  preferRelativeResize?: boolean;
}): Pick<PptxPictureOptions, "locking" | "preferRelativeResize"> {
  return {
    ...(source.locking !== undefined ? { locking: source.locking } : {}),
    ...(source.preferRelativeResize !== undefined
      ? { preferRelativeResize: source.preferRelativeResize }
      : {}),
  };
}

function docxPictureFill(
  fillRectangle: XlsxPictureOptions["fillRectangle"],
): Pick<DocxPictureOptions, "fillRectangle" | "stretch"> {
  if (fillRectangle === false) return { stretch: false };
  if (fillRectangle === true) return { fillRectangle: false };
  return fillRectangle === undefined ? {} : { fillRectangle };
}

function docxPictureFillFromPptx(
  fillRectangle: PptxPictureOptions["fillRectangle"],
): Pick<DocxPictureOptions, "fillRectangle"> {
  return fillRectangle === undefined ? {} : { fillRectangle };
}

function xlsxPictureFill(
  fillRectangle: DocxPictureOptions["fillRectangle"],
  stretch?: boolean,
): XlsxPictureOptions["fillRectangle"] {
  if (stretch === false) return false;
  if (fillRectangle === false) return true;
  return fillRectangle;
}

function pptxPictureFill(
  fillRectangle: XlsxPictureOptions["fillRectangle"],
): PptxPictureOptions["fillRectangle"] {
  if (fillRectangle === true) return false;
  return fillRectangle === false ? undefined : fillRectangle;
}

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
const docxType = (type: string): DocxRasterType => {
  if ((DOCX_RASTER_TYPES as readonly string[]).includes(type)) return type as DocxRasterType;
  throw new Error(`Cannot convert picture type "${type}" to DOCX without transcoding`);
};

const PPTX_TYPES: readonly PptxPictureOptions["type"][] = [
  "png",
  "jpg",
  "gif",
  "bmp",
  "tif",
  "ico",
  "emf",
  "wmf",
];
/** Narrow an image type to pptx's supported set without changing its bytes. */
const pptxType = (type: string): PptxPictureOptions["type"] => {
  if ((PPTX_TYPES as readonly string[]).includes(type)) return type as PptxPictureOptions["type"];
  throw new Error(`Cannot convert picture type "${type}" to PPTX without transcoding`);
};

/** Narrow an image type to xlsx's supported set without changing its bytes. */
const xlsxType = (type: string): XlsxPictureOptions["type"] => {
  if (type === "jpeg") return "jpg";
  if (["png", "jpg", "gif", "webp", "wmf", "emf"].includes(type)) {
    return type as XlsxPictureOptions["type"];
  }
  throw new Error(`Cannot convert picture type "${type}" to XLSX without transcoding`);
};

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
  if (!("col" in source)) {
    const p = source;
    const base = baseFromPptx(p);
    return {
      type: docxType(base.type),
      data: base.data,
      ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
      ...pickPictureAppearance(source),
      ...docxPictureFillFromPptx(source.fillRectangle),
      ...(source.useLocalDpi !== undefined
        ? { useLocalDpi: source.useLocalDpi === "default" ? false : source.useLocalDpi }
        : {}),
      transformation: boxToDocxTransformation({
        x: toEmu(p.x),
        y: toEmu(p.y),
        width: toEmu(p.width),
        height: toEmu(p.height),
        ...(p.rotation !== undefined ? { rotation: p.rotation } : {}),
        ...(p.flipHorizontal ? { flipHorizontal: true } : {}),
        ...(p.flipVertical ? { flipVertical: true } : {}),
      }),
      ...altTextFromBase(base),
    };
  }
  // xlsx → docx: cell anchor → offset EMU; size unknown.
  const x = source;
  const base = baseFromXlsx(x);
  return {
    type: docxType(base.type),
    data: base.data,
    ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
    ...pickPictureAppearance(x),
    ...(x.dpi !== undefined ? { dpi: x.dpi } : {}),
    ...docxPictureFill(x.fillRectangle),
    ...(x.useLocalDpi !== undefined
      ? { useLocalDpi: x.useLocalDpi === "default" ? false : x.useLocalDpi }
      : {}),
    transformation: boxToDocxTransformation({
      x: (x.col - 1) * DEFAULT_COL_EMU,
      y: (x.row - 1) * DEFAULT_ROW_EMU,
      width: 0,
      height: 0,
    }),
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
    const box = boxFromDocx(d.transformation);
    return {
      type: pptxType(base.type),
      data: base.data,
      ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
      ...boxToPptx(box),
      ...pickPictureAppearance(d),
      ...(d.fillRectangle !== undefined ? { fillRectangle: d.fillRectangle } : {}),
      ...(d.useLocalDpi !== undefined ? { useLocalDpi: d.useLocalDpi } : {}),
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
    ...pickPictureAppearance(x),
    ...pptxPictureFill(x.fillRectangle),
    ...pickPptxXlsxPictureOptions(x),
    ...(x.useLocalDpi !== undefined ? { useLocalDpi: x.useLocalDpi } : {}),
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
    const anchorBox = boxFromDocx(d.transformation);
    return {
      data: base.data,
      type: xlsxType(base.type),
      ...(base.sourceUrl !== undefined ? { sourceUrl: base.sourceUrl } : {}),
      col: emuToCell(anchorBox.x, DEFAULT_COL_EMU),
      row: emuToCell(anchorBox.y, DEFAULT_ROW_EMU),
      ...pickPictureAppearance(d),
      ...(d.dpi !== undefined ? { dpi: d.dpi } : {}),
      fillRectangle: xlsxPictureFill(d.fillRectangle, d.stretch),
      ...(d.useLocalDpi !== undefined ? { useLocalDpi: d.useLocalDpi } : {}),
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
    ...pickPictureAppearance(p),
    fillRectangle: p.fillRectangle === false ? true : p.fillRectangle,
    ...pickPptxXlsxPictureOptions(p),
    ...(p.useLocalDpi !== undefined ? { useLocalDpi: p.useLocalDpi } : {}),
    ...pickNonVisualDrawingProperties(base),
  };
}
