/**
 * Border module for WordprocessingML documents.
 *
 * Borders are used in multiple contexts (paragraphs, tables, table cells, sections)
 * and share the same CT_Border type definition. This module provides the BorderStyle
 * constants used throughout the document structure.
 *
 * Reference: http://officeopenxml.com/WPborders.php
 *
 * @see http://officeopenxml.com/WPtableBorders.php
 * @see http://officeopenxml.com/WPtableCellProperties-Borders.php
 * @see http://officeopenxml.com/WPsectionBorders.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Border">
 *   <xsd:attribute name="val" type="ST_Border" use="required"/>
 *   <xsd:attribute name="color" type="ST_HexColor" use="optional" default="auto"/>
 *   <xsd:attribute name="themeColor" type="ST_ThemeColor" use="optional"/>
 *   <xsd:attribute name="themeTint" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="themeShade" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="sz" type="ST_EighthPointMeasure" use="optional"/>
 *   <xsd:attribute name="space" type="ST_PointMeasure" use="optional" default="0"/>
 *   <xsd:attribute name="shadow" type="s:ST_OnOff" use="optional"/>
 *   <xsd:attribute name="frame" type="s:ST_OnOff" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import type { ThemeColor } from "@office-open/core";
import type { HexColorOrAuto, UcharHexNumber } from "@office-open/core";
import { attr, attrBool, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Options for configuring a border element.
 *
 * @property style - The border style (single, dashed, dotted, etc.)
 * @property color - Border color in hex format (e.g., "FF00AA" for purple)
 * @property size - Border thickness in eighths of a point (1/8 pt)
 * @property space - Spacing offset from the content in points
 */
export interface BorderOptions {
  /**
   * Border pattern (ST_Border). Token jargon: "wave"/"doubleWave" wavy,
   * "inset"/"outset" pseudo-3D, "thickThin*"/"thinThickThin*" compound lines,
   * "nil"/"none" no border. Accepts any ST_Border token — the BorderStyle
   * const lists the common line styles; Word also defines ~165 artistic and
   * pattern tokens (apples, weavingBraid, …) that round-trip as-is.
   */
  style: (typeof BorderStyle)[keyof typeof BorderStyle] | string;
  /** Border color, "auto" or hex (eg 'FF00AA') */
  color?: HexColorOrAuto;
  /** Theme color slot: "dark1"/"light1" text/background, "accent1"–"accent6" theme accents, "hyperlink"/"followedHyperlink". */
  themeColor?: ThemeColor;
  /** Theme color tint (2-char hex) */
  themeTint?: UcharHexNumber;
  /** Theme color shade (2-char hex) */
  themeShade?: UcharHexNumber;
  /** Border shadow */
  shadow?: boolean;
  /** Border frame */
  frame?: boolean;
  /** Size of the border in 1/8 pt */
  size?: number;
  /** Spacing offset. Values are specified in pt */
  space?: number;
}

/**
 * Border style tokens for the `style` field — the common line styles Word
 * renders as line patterns. Child elements of w:tblBorders specify the sides:
 * `top`, `bottom`, `start`/`left`, `end`/`right`, `insideH`, `insideV`.
 *
 * @publicApi
 */
export const BorderStyle = {
  /** A single line */
  SINGLE: "single",
  /** A line with a series of alternating thin and thick strokes */
  DASH_DOT_STROKED: "dashDotStroked",
  /** A dashed line */
  DASHED: "dashed",
  /** A dashed line with small gaps */
  DASH_SMALL_GAP: "dashSmallGap",
  /** A line with alternating dots and dashes */
  DOT_DASH: "dotDash",
  /** A line with a repeating dot - dot - dash sequence */
  DOT_DOT_DASH: "dotDotDash",
  /** A dotted line */
  DOTTED: "dotted",
  /** A double line */
  DOUBLE: "double",
  /** A double wavy line */
  DOUBLE_WAVE: "doubleWave",
  /** An inset set of lines */
  INSET: "inset",
  /** No border */
  NIL: "nil",
  /** No border */
  NONE: "none",
  /** An outset set of lines */
  OUTSET: "outset",
  /** A single line */
  THICK: "thick",
  /** A thick line contained within a thin line with a large-sized intermediate gap */
  THICK_THIN_LARGE_GAP: "thickThinLargeGap",
  /** A thick line contained within a thin line with a medium-sized intermediate gap */
  THICK_THIN_MEDIUM_GAP: "thickThinMediumGap",
  /** A thick line contained within a thin line with a small intermediate gap */
  THICK_THIN_SMALL_GAP: "thickThinSmallGap",
  /** A thin line contained within a thick line with a large-sized intermediate gap */
  THIN_THICK_LARGE_GAP: "thinThickLargeGap",
  /** A thick line contained within a thin line with a medium-sized intermediate gap */
  THIN_THICK_MEDIUM_GAP: "thinThickMediumGap",
  /** A thick line contained within a thin line with a small intermediate gap */
  THIN_THICK_SMALL_GAP: "thinThickSmallGap",
  /** A thin-thick-thin line with a large gap */
  THIN_THICK_THIN_LARGE_GAP: "thinThickThinLargeGap",
  /** A thin-thick-thin line with a medium gap */
  THIN_THICK_THIN_MEDIUM_GAP: "thinThickThinMediumGap",
  /** A thin-thick-thin line with a small gap */
  THIN_THICK_THIN_SMALL_GAP: "thinThickThinSmallGap",
  /** A three-staged gradient line, getting darker towards the paragraph */
  THREE_D_EMBOSS: "threeDEmboss",
  /** A three-staged gradient like, getting darker away from the paragraph */
  THREE_D_ENGRAVE: "threeDEngrave",
  /** A triple line */
  TRIPLE: "triple",
  /** A wavy line */
  WAVE: "wave",
} as const;

// ── Parse helper ──

/**
 * Parse one CT_Border side element. Returns undefined when the element is
 * malformed (missing `@w:val`) so callers skip the side. Any ST_Border token
 * passes through — Word's artistic styles (apples, weavingBraid, …) keep the
 * whole side instead of being dropped.
 */
export function parseBorderSide(sideEl: Element): BorderOptions | undefined {
  const style = attr(sideEl, "w:val");
  if (!style) return undefined;
  const sideOpts: BorderOptions = { style: style as BorderOptions["style"] };
  const color = attr(sideEl, "w:color");
  if (color) sideOpts.color = color;
  const size = attrNum(sideEl, "w:sz");
  if (size !== undefined) sideOpts.size = size;
  const space = attrNum(sideEl, "w:space");
  if (space !== undefined) sideOpts.space = space;
  const themeColor = attr(sideEl, "w:themeColor");
  if (themeColor) {
    sideOpts.themeColor = themeColor as ThemeColor;
  }
  const themeTint = attr(sideEl, "w:themeTint");
  if (themeTint) sideOpts.themeTint = themeTint;
  const themeShade = attr(sideEl, "w:themeShade");
  if (themeShade) sideOpts.themeShade = themeShade;
  const shadow = attrBool(sideEl, "w:shadow");
  if (shadow !== undefined) sideOpts.shadow = shadow;
  const frame = attrBool(sideEl, "w:frame");
  if (frame !== undefined) sideOpts.frame = frame;
  return sideOpts;
}
