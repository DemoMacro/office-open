/**
 * Cross-format position helpers shared by the picture/shape/connector/group
 * converters. Each format uses a different coordinate model; this module
 * translates between them via an absolute EMU bounding box.
 *
 * - pptx: absolute EMU coordinates as top-level x/y/w/h.
 * - docx: {@link MediaTransformation} (offset left/top + width/height).
 * - xlsx: 1-based cell anchors. Column/row sizes are heuristic (8.43-char
 *   column × 15pt row), so xlsx ↔ {pptx,docx} loses precise positioning — the
 *   same loss MS Office paste incurs between apps.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { MediaTransformation } from "@office-open/docx";
import type { DrawingAnchorOptions } from "@office-open/xlsx";

/** Heuristic default column width in EMU (8.43 chars ≈ 64 px at 96 DPI). */
export const DEFAULT_COL_EMU = 609600;
/** Heuristic default row height in EMU (15 pt). */
export const DEFAULT_ROW_EMU = 190500;

/** Coerce a coordinate (EMU number or universal measure) to raw EMU. */
export function toEmu(value: number | UniversalMeasure | undefined, fallback = 0): number {
  return value === undefined ? fallback : convertToEmu(value);
}

/** Convert a raw EMU offset to a 1-based cell index. */
export function emuToCell(emus: number, cellEmu: number): number {
  return Math.floor(emus / cellEmu) + 1;
}

/** Absolute EMU bounding box (top-left + size + optional rotation/flip). */
export interface AbsoluteBox {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}

// ── → box ──

/** Build a box from pptx top-level position fields. */
export function boxFromPptx(
  x: number | UniversalMeasure | undefined,
  y: number | UniversalMeasure | undefined,
  width: number | UniversalMeasure | undefined,
  height: number | UniversalMeasure | undefined,
  rotation?: number,
  flipHorizontal?: boolean,
): AbsoluteBox {
  return {
    x: toEmu(x),
    y: toEmu(y),
    width: toEmu(width),
    height: toEmu(height),
    ...(rotation !== undefined ? { rotation } : {}),
    ...(flipHorizontal ? { flipHorizontal: true } : {}),
  };
}

/**
 * Build a box from a core spPr transform (off/ext + rotation/flip). Used for
 * group children, which position via spPr.xfrm with no cell anchor.
 */
export function boxFromSpPr(spPr: {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  rotation?: number;
  flipHorizontal?: boolean;
  flipVertical?: boolean;
}): AbsoluteBox {
  return {
    x: toEmu(spPr.x),
    y: toEmu(spPr.y),
    width: toEmu(spPr.width),
    height: toEmu(spPr.height),
    ...(spPr.rotation !== undefined ? { rotation: spPr.rotation } : {}),
    ...(spPr.flipHorizontal ? { flipHorizontal: true } : {}),
    ...(spPr.flipVertical ? { flipVertical: true } : {}),
  };
}

/**
 * Build a box from an xlsx cell anchor. The absolute top-left comes from the
 * from-marker (col/row + offsets); the size comes from the caller (spPr.xfrm
 * extent or the to-marker, depending on the source).
 */
export function boxFromXlsxAnchor(
  anchor: DrawingAnchorOptions,
  width: number | UniversalMeasure | undefined,
  height: number | UniversalMeasure | undefined,
  rotation?: number,
  flipHorizontal?: boolean,
  flipVertical?: boolean,
): AbsoluteBox {
  const x = (anchor.col - 1) * DEFAULT_COL_EMU + toEmu(anchor.colOffset);
  const y = (anchor.row - 1) * DEFAULT_ROW_EMU + toEmu(anchor.rowOffset);
  return {
    x,
    y,
    width: toEmu(width),
    height: toEmu(height),
    ...(rotation !== undefined ? { rotation } : {}),
    ...(flipHorizontal ? { flipHorizontal: true } : {}),
    ...(flipVertical ? { flipVertical: true } : {}),
  };
}

/** Build a box from a docx MediaTransformation. */
export function boxFromDocx(transformation: MediaTransformation): AbsoluteBox {
  return {
    x: toEmu(transformation.offset?.left),
    y: toEmu(transformation.offset?.top),
    width: toEmu(transformation.width),
    height: toEmu(transformation.height),
    ...(transformation.rotation !== undefined ? { rotation: transformation.rotation } : {}),
    ...(transformation.flip?.horizontal ? { flipHorizontal: true } : {}),
    ...(transformation.flip?.vertical ? { flipVertical: true } : {}),
  };
}

// ── box → ──

/** Pptx top-level position fields derived from a box. */
export interface PptxPosition {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  flipHorizontal?: boolean;
}

/** Emit pptx top-level position fields from a box. */
export function boxToPptx(box: AbsoluteBox): PptxPosition {
  return {
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    ...(box.rotation !== undefined ? { rotation: box.rotation } : {}),
    ...(box.flipHorizontal ? { flipHorizontal: true } : {}),
  };
}

/** xlsx position: cell anchor plus the matching spPr.xfrm offset/extent. */
export interface XlsxPosition {
  anchor: DrawingAnchorOptions;
  /** spPr.xfrm.off.x — the in-cell horizontal offset (= anchor colOffset). */
  xfrmX: number;
  /** spPr.xfrm.off.y — the in-cell vertical offset (= anchor rowOffset). */
  xfrmY: number;
}

/**
 * Emit an xlsx position from a box. The from-marker locates the cell, the
 * to-marker carries the size (twoCellAnchor), and the xfrm offset mirrors the
 * from-marker offset so the anchor and spPr agree.
 */
export function boxToXlsx(box: AbsoluteBox): XlsxPosition {
  const col = emuToCell(box.x, DEFAULT_COL_EMU);
  const row = emuToCell(box.y, DEFAULT_ROW_EMU);
  const colOffset = box.x - (col - 1) * DEFAULT_COL_EMU;
  const rowOffset = box.y - (row - 1) * DEFAULT_ROW_EMU;
  return {
    anchor: {
      col,
      row,
      colOffset,
      rowOffset,
      toCol: emuToCell(box.x + box.width, DEFAULT_COL_EMU),
      toRow: emuToCell(box.y + box.height, DEFAULT_ROW_EMU),
    },
    xfrmX: colOffset,
    xfrmY: rowOffset,
  };
}

/** Emit a docx MediaTransformation from a box. */
export function boxToDocx(box: AbsoluteBox): MediaTransformation {
  return {
    offset: { left: box.x, top: box.y },
    width: box.width,
    height: box.height,
    ...(box.rotation !== undefined ? { rotation: box.rotation } : {}),
    ...(box.flipHorizontal || box.flipVertical
      ? {
          flip: {
            ...(box.flipHorizontal ? { horizontal: true } : {}),
            ...(box.flipVertical ? { vertical: true } : {}),
          },
        }
      : {}),
  };
}
