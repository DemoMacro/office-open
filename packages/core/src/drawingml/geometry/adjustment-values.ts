/**
 * Adjustment values module for preset geometries.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";

/**
 * A single geometry guide that defines an adjustment value.
 */
export interface GeometryGuide {
  /** Guide name (identifier) */
  name: string;
  /** Guide formula (e.g., "val 16667") */
  formula: string;
}

/**
 * Build a:avLst XML string from geometry guides.
 */
export function stringifyAdjustmentValues(guides?: readonly GeometryGuide[]): string {
  if (!guides || guides.length === 0) return "<a:avLst/>";
  const inner = guides
    .map((g) => `<a:gd name="${escapeXml(g.name)}" fmla="${escapeXml(g.formula)}"/>`)
    .join("");
  return `<a:avLst>${inner}</a:avLst>`;
}
