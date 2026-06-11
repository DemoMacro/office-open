import { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle } from "@office-open/core";
/**
 * Run properties types and constants for PPTX text runs.
 *
 * @module
 */
import type { FillOptions } from "@shared/drawingml/fill";

// Re-export XSD converters for descriptor use
export { xsdStrikeStyle, xsdTextCaps, xsdUnderlineStyle };

export const UnderlineStyle = {
  SINGLE: "single",
  DOUBLE: "double",
  NONE: "none",
} as const;

export const StrikeStyle = {
  SINGLE: "sngStrike",
  DOUBLE: "dblStrike",
  NONE: "noStrike",
} as const;

export const TextCapitalization = {
  NONE: "none",
  ALL: "all",
  SMALL: "small",
} as const;

export interface HyperlinkOptions {
  url: string;
  tooltip?: string;
  action?: string;
  highlightClick?: boolean;
  endSound?: boolean;
  invalidUrl?: boolean;
}

export interface RunPropertiesOptions {
  /** Font size in points. Serialized as OOXML `a:sz` (hundredths of a point). */
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: (typeof UnderlineStyle)[keyof typeof UnderlineStyle];
  font?: string;
  lang?: string;
  fill?: FillOptions;
  hyperlink?: HyperlinkOptions;
  strike?: (typeof StrikeStyle)[keyof typeof StrikeStyle];
  baseline?: number;
  spacing?: number;
  capitalization?: (typeof TextCapitalization)[keyof typeof TextCapitalization];
  shadow?: boolean;
  outline?: boolean;
  rightToLeft?: boolean;
  noProof?: boolean;
  dirty?: boolean;
  kumimoji?: boolean;
  alternateLanguage?: string;
  normalizeHeight?: boolean;
  bookmarkMark?: string;
  smartTagId?: string;
}
