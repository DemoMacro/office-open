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

import type { SolidFillOptions } from "../color/solid-fill";
import type { EffectListOptions } from "../effects/effect-list";
import type { FillOptions } from "../fill/fill-options";
import type { OutlineOptions } from "../outline/outline";

// ── Run style enums ──

export type UnderlineStyle = "single" | "double" | "none";

// Friendly spellings; xsdStrikeStyle maps them to ST_TextStrikeType tokens.
export type StrikeStyle = "singleStrike" | "doubleStrike" | "noStrike";

export type TextCapitalization = "none" | "all" | "small";

// ── Paragraph alignment ──

export type TextAlignment = "left" | "center" | "right" | "justify";

// ── Hyperlink (a:hlinkClick) ──

export interface TextHyperlinkOptions {
  /** External URL target (mutually exclusive with slide). */
  url?: string;
  /**
   * Internal slide target, 1-based slide number (mutually exclusive with url).
   * Emitted as r:id → slideN.xml + action="ppaction://hlinksldjump".
   *
   * PPTX only — slides exist in no other format. Formats without slides
   * reject the field at generate time instead of silently dropping it.
   */
  slide?: number;
  /** Internal placeholder key ("{hlink:N}" → "N"); preserved across parse → stringify. */
  referenceId?: string;
  tooltip?: string;
  action?: string;
  highlightClick?: boolean;
  endSound?: boolean;
  invalidUrl?: boolean;
}

// ── Text fonts (CT_TextFont: a:latin / a:ea / a:cs / a:sym) ──

/**
 * A script typeface (CT_TextFont). A bare string is the typeface only; the
 * full object carries panose/pitchFamily/charset for embedded font metrics.
 */
export type TextFont =
  | string
  | {
      typeface: string;
      /** ST_Panose — 10-byte hex panose classification. */
      panose?: string;
      pitchFamily?: number;
      charset?: number;
    };

/**
 * Run font scripts. A bare string sets latin + ea to the same typeface (the
 * common case); the object form sets each script independently.
 */
export type RunFont =
  | string
  | {
      latin?: TextFont;
      eastAsia?: TextFont;
      complexScript?: TextFont;
      symbol?: TextFont;
    };

// ── Run properties (CT_TextCharacterProperties) ──

export interface RunPropertiesOptions {
  /** Font size in points. Serialized as OOXML a:sz (hundredths of a point). */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  font?: RunFont;
  lang?: string;
  fill?: FillOptions;
  hyperlink?: TextHyperlinkOptions;
  /** a:hlinkMouseOver — hover hyperlink (CT_Hyperlink). */
  mouseoverHyperlink?: TextHyperlinkOptions;
  strike?: StrikeStyle;
  /** Baseline offset as integer percent (positive = superscript, negative = subscript). */
  baseline?: number;
  /** Character spacing in points (a:spc). */
  spacing?: number;
  capitalization?: TextCapitalization;
  /** a:highlight (CT_Color) — text highlight color. */
  highlight?: SolidFillOptions;
  /** EG_TextUnderlineLine — underline line. `true` = a:uLnTx (follow the text line); an OutlineOptions = a:uLn. */
  underlineLine?: true | OutlineOptions;
  /** EG_TextUnderlineFill — underline fill. `true` = a:uFillTx (follow the text fill); a FillOptions = a:uFill. */
  underlineFill?: true | FillOptions;
  /** @kern — kerning threshold in points (ST_TextNonNegativePoint). */
  kern?: number;
  /** a:ln (CT_LineProperties). `true` emits a sane default; a full OutlineOptions round-trips. */
  outline?: true | OutlineOptions;
  /** a:effectLst (EG_EffectProperties). `true` emits a default outer shadow; a full EffectListOptions round-trips. */
  shadow?: true | EffectListOptions;
  rightToLeft?: boolean;
  noProof?: boolean;
  dirty?: boolean;
  kumimoji?: boolean;
  alternateLanguage?: string;
  normalizeHeight?: boolean;
  bookmarkMark?: string;
  smartTagId?: string;
  /** @err — spelling error flag. */
  err?: boolean;
  /** @smtClean — smart tag clean flag (XSD default true). */
  smtClean?: boolean;
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
  /** a:buSzPts @val — bullet size in points. */
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
  /** Space after a paragraph in points (a:spcAft/a:spcPts). */
  spaceAfter?: number;
  /** Space before a paragraph in points (a:spcBef/a:spcPts). */
  spaceBefore?: number;
  bullet?: BulletOptions;
  /** Line spacing as a percentage (100 = single). */
  lineSpacingPercent?: number;
  /** Line spacing in exact points (a:lnSpc/a:spcPts). */
  lineSpacingPoints?: number;
  marginIndent?: number;
  marginRight?: number;
  defTabSize?: number;
  /** @indent — first-line indent (ST_TextIndent, EMU). */
  indent?: number;
  /** a:tabLst — explicit tab stops (emitted after bullets, before defRPr). */
  tabStops?: TabStopOptions[];
  /** a:defRPr — default run properties for the paragraph (CT_TextCharacterProperties). */
  defaultRunProperties?: RunPropertiesOptions;
  fontAlignment?: "auto" | "t" | "ctr" | "b" | "base";
  /** @rtl — paragraph right-to-left. */
  rightToLeft?: boolean;
  /** @eaLnBrk — East Asian line break. */
  eastAsianLineBreak?: boolean;
  /** @latinLnBrk — Latin line break. */
  latinLineBreak?: boolean;
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
/** Default shadow direction: 270° (degrees). */
export const DEFAULT_SHADOW_DIRECTION = 270;
/** Default shadow alpha: 40% (integer percent). */
export const DEFAULT_SHADOW_ALPHA = 40;
