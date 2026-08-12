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
 *   no size on the public input (`XlsxPictureOptions` carries only `col/row`).
 * - EMU→cell divides by a default cell size (8.43-char column × 15pt row);
 *   cell→EMU multiplies back. Size is lost on the xlsx leg.
 *
 * The loss is documented per function and matches MS Office paste behavior.
 *
 * @module
 */
import type { PictureOptions as DocxPictureOptions } from "@office-open/docx";
import type { PictureOptions as PptxPictureOptions } from "@office-open/pptx";
import type { PictureOptions as XlsxPictureOptions } from "@office-open/xlsx";

import { DEFAULT_COL_EMU, DEFAULT_ROW_EMU, emuToCell, toEmu } from "./position";

// ── → docx ──

/** Convert a pptx picture to a docx inline image. */
export function toDocxPicture(source: PptxPictureOptions): DocxPictureOptions;
/** Convert an xlsx image to a docx inline image (size defaults to 0; xlsx carries no size). */
export function toDocxPicture(source: XlsxPictureOptions): DocxPictureOptions;
export function toDocxPicture(source: PptxPictureOptions | XlsxPictureOptions): DocxPictureOptions {
  // pptx → docx: absolute x/y → offset, width/height → transformation.
  if ("width" in source || "height" in source) {
    const p = source as PptxPictureOptions;
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
  const x = source as XlsxPictureOptions;
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
export function toPptxPicture(source: DocxPictureOptions): PptxPictureOptions;
/** Convert an xlsx image to a pptx picture (size defaults to 0; xlsx carries no size). */
export function toPptxPicture(source: XlsxPictureOptions): PptxPictureOptions;
export function toPptxPicture(source: DocxPictureOptions | XlsxPictureOptions): PptxPictureOptions {
  // docx → pptx: transformation → absolute x/y + width/height.
  if ("transformation" in source) {
    const d = source as DocxPictureOptions;
    const t = d.transformation;
    return {
      type: d.type as PptxPictureOptions["type"],
      data: d.data,
      width: t.width,
      height: t.height,
      ...(t.offset ? { x: t.offset.left, y: t.offset.top } : {}),
    };
  }
  // xlsx → pptx: cell anchor → absolute EMU; size unknown.
  const x = source as XlsxPictureOptions;
  return {
    type: x.type as PptxPictureOptions["type"],
    data: x.data,
    x: (x.col - 1) * DEFAULT_COL_EMU,
    y: (x.row - 1) * DEFAULT_ROW_EMU,
    width: 0,
    height: 0,
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
  const p = source as PptxPictureOptions;
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
