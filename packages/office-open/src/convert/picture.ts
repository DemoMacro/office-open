/**
 * Cross-format picture conversion.
 *
 * Picture/image options convert between docx, pptx, and xlsx. The output
 * carries the binary `data`; each package's own generate registers the media,
 * so these are pure functions with no context injection.
 *
 * Position mapping is heuristic where the target has no matching coordinate
 * model:
 * - pptx/docx use absolute EMU coordinates; xlsx uses 1-based cell anchors with
 *   no size on the public input (`WorksheetImageOptions` carries only `col/row`).
 * - EMU→cell divides by a default cell size (8.43-char column × 15pt row);
 *   cell→EMU multiplies back. Size is lost on the xlsx leg.
 *
 * The loss is documented per function and matches MS Office paste behavior.
 *
 * @module
 */
import { convertToEmu, type UniversalMeasure } from "@office-open/core";
import type { ImageOptions } from "@office-open/docx";
import type { PictureOptions } from "@office-open/pptx";
import type { WorksheetImageOptions } from "@office-open/xlsx";

/** Heuristic default column width in EMU (8.43 chars ≈ 64 px at 96 DPI). */
const DEFAULT_COL_EMU = 609600;
/** Heuristic default row height in EMU (15 pt). */
const DEFAULT_ROW_EMU = 190500;

/** Coerce a coordinate (EMU number or universal measure) to raw EMU. */
function toEmu(value: number | UniversalMeasure | undefined, fallback = 0): number {
  return value === undefined ? fallback : convertToEmu(value);
}

/** Convert a raw EMU offset to a 1-based cell index. */
function emuToCell(emus: number, cellEmu: number): number {
  return Math.floor(emus / cellEmu) + 1;
}

// ── → docx ──

/** Convert a pptx picture to a docx inline image. */
export function toDocxPicture(source: PictureOptions): ImageOptions;
/** Convert an xlsx image to a docx inline image (size defaults to 0; xlsx carries no size). */
export function toDocxPicture(source: WorksheetImageOptions): ImageOptions;
export function toDocxPicture(source: PictureOptions | WorksheetImageOptions): ImageOptions {
  // pptx → docx: absolute x/y → offset, width/height → transformation.
  if ("width" in source || "height" in source) {
    const p = source as PictureOptions;
    return {
      type: p.type,
      data: p.data,
      transformation: {
        width: p.width ?? 0,
        height: p.height ?? 0,
        ...(p.x !== undefined || p.y !== undefined
          ? { offset: { left: p.x ?? 0, top: p.y ?? 0 } }
          : {}),
      },
    };
  }
  // xlsx → docx: cell anchor → offset EMU; size unknown.
  const x = source as WorksheetImageOptions;
  return {
    type: x.type,
    data: x.data,
    transformation: {
      width: 0,
      height: 0,
      offset: { left: (x.col - 1) * DEFAULT_COL_EMU, top: (x.row - 1) * DEFAULT_ROW_EMU },
    },
  };
}

// ── → pptx ──

/** Convert a docx image to a pptx picture. */
export function toPptxPicture(source: ImageOptions): PictureOptions;
/** Convert an xlsx image to a pptx picture (size defaults to 0; xlsx carries no size). */
export function toPptxPicture(source: WorksheetImageOptions): PictureOptions;
export function toPptxPicture(source: ImageOptions | WorksheetImageOptions): PictureOptions {
  // docx → pptx: transformation → absolute x/y + width/height.
  if ("transformation" in source) {
    const d = source as ImageOptions;
    const t = d.transformation;
    return {
      type: d.type as PictureOptions["type"],
      data: d.data,
      width: t.width,
      height: t.height,
      ...(t.offset ? { x: t.offset.left, y: t.offset.top } : {}),
    };
  }
  // xlsx → pptx: cell anchor → absolute EMU; size unknown.
  const x = source as WorksheetImageOptions;
  return {
    type: x.type as PictureOptions["type"],
    data: x.data,
    x: (x.col - 1) * DEFAULT_COL_EMU,
    y: (x.row - 1) * DEFAULT_ROW_EMU,
    width: 0,
    height: 0,
  };
}

// ── → xlsx ──

/** Convert a docx image to an xlsx image (position mapped to cell anchor; size lost). */
export function toXlsxImage(source: ImageOptions): WorksheetImageOptions;
/** Convert a pptx picture to an xlsx image (position mapped to cell anchor; size lost). */
export function toXlsxImage(source: PictureOptions): WorksheetImageOptions;
export function toXlsxImage(source: ImageOptions | PictureOptions): WorksheetImageOptions {
  // docx → xlsx: offset EMU → cell anchor.
  if ("transformation" in source) {
    const d = source as ImageOptions;
    const left = toEmu(d.transformation.offset?.left);
    const top = toEmu(d.transformation.offset?.top);
    return {
      data: d.data,
      type: xlsxType(d.type),
      col: emuToCell(left, DEFAULT_COL_EMU),
      row: emuToCell(top, DEFAULT_ROW_EMU),
    };
  }
  // pptx → xlsx: absolute EMU → cell anchor.
  const p = source as PictureOptions;
  return {
    data: p.data,
    type: xlsxType(p.type),
    col: emuToCell(toEmu(p.x), DEFAULT_COL_EMU),
    row: emuToCell(toEmu(p.y), DEFAULT_ROW_EMU),
  };
}

/** Narrow an image type to xlsx's png/jpg set (xlsx public input only accepts these). */
function xlsxType(type: string): "png" | "jpg" {
  return type === "jpg" || type === "jpeg" ? "jpg" : "png";
}
