import { convertToTwip, decimalNumber, mapOptional } from "@office-open/core";
import type { PositiveUniversalMeasure } from "@office-open/core";
import { element } from "@office-open/xml";

/**
 * When line numbering restarts (ST_LineNumberRestart): "newPage" (default),
 * "newSection", or "continuous" (carry over from the previous section).
 *
 * @publicApi
 */
export const LineNumberRestartFormat = {
  NEW_PAGE: "newPage",
  NEW_SECTION: "newSection",
  CONTINUOUS: "continuous",
} as const;

export interface LineNumberProperties {
  /** Display a number only on lines whose count is a multiple of this (e.g., 5 → numbers on lines 5, 10, 15). */
  countBy?: number;
  /** First line number used after each restart (default 1). */
  start?: number;
  /** Restart point for the numbering (default "newPage"). */
  restart?: (typeof LineNumberRestartFormat)[keyof typeof LineNumberRestartFormat];
  /** Gap between the text margin and the line-number edge, in twips. */
  distance?: number | PositiveUniversalMeasure;
}

/** Line numbering settings (w:lnNumType) for a section (CT_LineNumber). */
export const createLineNumberType = ({
  countBy,
  start,
  restart,
  distance,
}: LineNumberProperties): string =>
  element("w:lnNumType", {
    "w:countBy": mapOptional(countBy, decimalNumber),
    "w:distance": mapOptional(distance, convertToTwip),
    "w:restart": restart,
    "w:start": mapOptional(start, decimalNumber),
  });
