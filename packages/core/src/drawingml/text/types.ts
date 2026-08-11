/**
 * Shared DrawingML text option types, enums, and constants.
 *
 * Promoted from per-format implementations so DOCX/PPTX/XLSX share one text
 * model for runs, paragraphs, fields, and bullets. DOCX still uses
 * WordprocessingML (w:p) inside text boxes, but its DrawingML run/paragraph
 * options align with these types where they overlap.
 *
 * @module
 */

import type { FillOptions } from "../fill/fill-options";

// ── Run style enums ──

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

// ── Paragraph alignment ──

export type TextAlignment = "left" | "center" | "right" | "justify";

// ── Hyperlink (a:hlinkClick) ──

export interface HyperlinkOptions {
  url: string;
  /** Internal placeholder key ("{hlink:N}" → "N"); preserved across parse → stringify. */
  referenceId?: string;
  tooltip?: string;
  action?: string;
  highlightClick?: boolean;
  endSound?: boolean;
  invalidUrl?: boolean;
}

// ── Run properties (CT_TextCharacterProperties) ──

export interface RunPropertiesOptions {
  /** Font size in points. Serialized as OOXML a:sz (hundredths of a point). */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: (typeof UnderlineStyle)[keyof typeof UnderlineStyle];
  font?: string;
  lang?: string;
  fill?: FillOptions;
  hyperlink?: HyperlinkOptions;
  /** a:hlinkMouseOver — hover hyperlink (CT_Hyperlink). */
  mouseoverHyperlink?: HyperlinkOptions;
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

// ── Run (a:r) ──

export interface RunOptions extends RunPropertiesOptions {
  text?: string;
}

// ── Bullets ──

/** Shared bullet color/size/font styling (EG_TextBulletColorSizeFont). Each
 * dimension is a choice: an explicit value, a "follows text" toggle, or unset. */
export type BulletStyleOptions = {
  /** a:buClr > a:srgbClr — explicit bullet color (hex, no #). */
  color?: string;
  /** a:buClrTx — bullet color follows the text run color. */
  colorFollowsText?: boolean;
  /** a:buSzPct @val — bullet size as a percentage of the text size. */
  size?: number;
  /** a:buSzTx — bullet size follows the text run size. */
  sizeFollowsText?: boolean;
  /** a:buSzPts @val — bullet size in hundredths of a point. */
  sizePoints?: number;
  /** a:buFont @typeface — bullet font (defaults to Arial on fresh char/autoNum). */
  font?: string;
  /** a:buFontTx — bullet font follows the text run font. */
  fontFollowsText?: boolean;
};

export type BulletCharOptions = BulletStyleOptions & {
  type: "char";
  char?: string;
};

export type BulletAutoNumOptions = BulletStyleOptions & {
  type: "autoNum";
  format?: string;
  startAt?: number;
};

/** Picture bullet (a:buBlip r:embed). `embed` is the image relationship id. */
export type BulletPictureOptions = BulletStyleOptions & {
  type: "picture";
  embed: string;
};

export type BulletNoneOption = { type: "none" };

export type BulletOptions =
  | BulletCharOptions
  | BulletAutoNumOptions
  | BulletPictureOptions
  | BulletNoneOption;

// ── Tab stops (a:tabLst) ──

export type TextTabAlignment = "l" | "ctr" | "r" | "dec";

export interface TabStopOptions {
  /** a:tab @pos — tab stop position in EMU. */
  position?: number;
  /** a:tab @algn — tab alignment (l/ctr/r/dec). */
  alignment?: TextTabAlignment;
}

// ── Paragraph properties (a:pPr) ──

export interface ParagraphPropertiesOptions {
  alignment?: TextAlignment;
  indentLevel?: number;
  marginBottom?: number;
  marginTop?: number;
  bullet?: BulletOptions;
  lineSpacing?: number;
  lineSpacingPoints?: number;
  marginIndent?: number;
  marginRight?: number;
  defTabSize?: number;
  /** a:tabLst — explicit tab stops (emitted after bullets, before defRPr). */
  tabStops?: TabStopOptions[];
  fontAlignment?: "auto" | "t" | "ctr" | "b" | "base";
}

// ── Text field (a:fld) ──

export interface TextFieldOptions {
  /** a:fld @type — field type token (e.g. "datetimeFigureOut", "slidenum"). */
  type: string;
  /** a:fld @id — GUID identifier. */
  id?: string;
  /** a:t — display text (often a placeholder such as "‹#›" or "1/27/13"). */
  text?: string;
  /** a:rPr — run properties. */
  properties?: RunPropertiesOptions;
}

// ── Soft line break (a:br) ──

export interface BreakOptions {
  /** Marks a soft line break (a:br) within a paragraph's children. */
  break: true;
  properties?: RunPropertiesOptions;
}

// ── Defaults ──

/** Default outline width: 1pt = 12700 EMU. */
export const DEFAULT_OUTLINE_WIDTH = 12700;
/** Default shadow blur radius: ~4pt = 50800 EMU. */
export const DEFAULT_SHADOW_BLUR_RADIUS = 50800;
/** Default shadow distance: ~3pt = 38100 EMU. */
export const DEFAULT_SHADOW_DISTANCE = 38100;
/** Default shadow direction: 270° = 2700000 (60,000ths of a degree). */
export const DEFAULT_SHADOW_DIRECTION = 2700000;
/** Default shadow alpha: 40% = 40000 (100,000ths of a percent). */
export const DEFAULT_SHADOW_ALPHA = 40000;
