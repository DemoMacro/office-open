import { decimalNumber } from "@office-open/core";
import { element } from "@office-open/xml";

/**
 * Line pitch behavior of the East Asian document grid (ST_DocGrid): "default"
 * no grid, "lines" snap lines to the grid, "linesAndChars" snap lines and
 * characters, "snapToChars" snap characters only.
 */
export const DocumentGridType = {
  DEFAULT: "default",
  LINES: "lines",
  LINES_AND_CHARS: "linesAndChars",
  SNAP_TO_CHARS: "snapToChars",
} as const;

export interface DocGridProperties {
  /** Grid behavior — see {@link DocumentGridType} values. */
  type?: (typeof DocumentGridType)[keyof typeof DocumentGridType];
  /** Line pitch in twentieths of a point (684 = 34.2 pt). Paragraphs opt out via snapToGrid or exact lineRule spacing. */
  linePitch: number;
  /** Character-pitch delta over the Normal font, encoded ×4096 (40960 = +10 pt over the Normal pitch). Runs opt out via snapToGrid. */
  charSpace?: number;
}

/** Document grid (w:docGrid) — line/character pitch grid for East Asian text (CT_DocGrid). */
export const createDocumentGrid = ({ type, linePitch, charSpace }: DocGridProperties): string =>
  element("w:docGrid", {
    "w:charSpace": charSpace ? decimalNumber(charSpace) : undefined,
    "w:linePitch": decimalNumber(linePitch),
    "w:type": type,
  });
