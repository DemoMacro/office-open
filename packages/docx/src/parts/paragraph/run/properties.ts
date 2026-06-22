/**
 * Run properties module for WordprocessingML documents.
 *
 * This module provides the run properties (rPr) element which specifies
 * the formatting applied to a run of text.
 *
 * Reference: https://www.ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
 * Section 17.3.2.21 (page 297)
 *
 * @module
 */
import type { UniversalMeasure } from "@office-open/core";
import type { BorderOptions } from "@shared/border";
import type { ShadingAttributesProperties } from "@shared/shading";
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";

import type { EastAsianLayoutOptions } from "./east-asian-layout";
import type { EmphasisMarkType } from "./emphasis-mark";
import type { ColorOptions } from "./formatting";
import type { LanguageOptions } from "./language";
import type { FontAttributesProperties } from "./run-fonts";
import type { UnderlineType } from "./underline";

interface RunFontReference {
  name: string;
  hint?: string;
}

/**
 * Text animation effect types.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 *
 * @publicApi
 */
export const TextEffect = {
  /** Blinking background animation */
  BLINK_BACKGROUND: "blinkBackground",
  /** Lights animation effect */
  LIGHTS: "lights",
  /** Black marching ants animation */
  ANTS_BLACK: "antsBlack",
  /** Red marching ants animation */
  ANTS_RED: "antsRed",
  /** Shimmer animation effect */
  SHIMMER: "shimmer",
  /** Sparkle animation effect */
  SPARKLE: "sparkle",
  /** No text effect */
  NONE: "none",
} as const;

/**
 * Highlight color values for text highlighting.
 *
 * Reference: http://officeopenxml.com/WPtextShading.php
 */
export const HighlightColor = {
  /** Black highlight */
  BLACK: "black",
  /** Blue highlight */
  BLUE: "blue",
  /** Cyan highlight */
  CYAN: "cyan",
  /** Dark blue highlight */
  DARK_BLUE: "darkBlue",
  /** Dark cyan highlight */
  DARK_CYAN: "darkCyan",
  /** Dark gray highlight */
  DARK_GRAY: "darkGray",
  /** Dark green highlight */
  DARK_GREEN: "darkGreen",
  /** Dark magenta highlight */
  DARK_MAGENTA: "darkMagenta",
  /** Dark red highlight */
  DARK_RED: "darkRed",
  /** Dark yellow highlight */
  DARK_YELLOW: "darkYellow",
  /** Green highlight */
  GREEN: "green",
  /** Light gray highlight */
  LIGHT_GRAY: "lightGray",
  /** Magenta highlight */
  MAGENTA: "magenta",
  /** No highlight */
  NONE: "none",
  /** Red highlight */
  RED: "red",
  /** White highlight */
  WHITE: "white",
  /** Yellow highlight */
  YELLOW: "yellow",
} as const;

/**
 * Run style properties options.
 *
 * These properties define the formatting that can be applied to a run of text,
 * including font, size, bold, italic, underline, color, and other character formatting.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 */
export interface RunStylePropertiesOptions {
  noProof?: boolean;
  bold?: boolean;
  boldComplexScript?: boolean;
  italic?: boolean;
  italicComplexScript?: boolean;
  underline?: {
    color?: string;
    type?: (typeof UnderlineType)[keyof typeof UnderlineType];
  };
  effect?: (typeof TextEffect)[keyof typeof TextEffect];
  emphasisMark?: {
    type?: (typeof EmphasisMarkType)[keyof typeof EmphasisMarkType];
  };
  color?: string | ColorOptions;
  kern?: number;
  position?: string;
  /** Font size in points. Internally stored as half-points in XML (×2). */
  size?: number;
  /** Complex-script font size in points (w:szCs). Independent from {@link size}. */
  sizeComplexScript?: number;
  rightToLeft?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  subScript?: boolean;
  superScript?: boolean;
  font?: string | RunFontReference | FontAttributesProperties;
  highlight?: (typeof HighlightColor)[keyof typeof HighlightColor];
  /** Complex-script highlight color (w:highlightCs). Independent from {@link highlight}. */
  highlightComplexScript?: string;
  characterSpacing?: number | UniversalMeasure;
  shading?: ShadingAttributesProperties;
  emboss?: boolean;
  imprint?: boolean;
  revision?: RunPropertiesChangeOptions;
  language?: LanguageOptions;
  border?: BorderOptions;
  snapToGrid?: boolean;
  vanish?: boolean;
  specVanish?: boolean;
  scale?: number;
  math?: boolean;
  outline?: boolean;
  shadow?: boolean;
  webHidden?: boolean;
  fitText?: number;
  complexScript?: boolean;
  eastAsianLayout?: EastAsianLayoutOptions;
  /** Relationship ID for a content part (w:contentPart with r:id) */
  contentPartRId?: string;
  /**
   * Raw XML for w14:* text-effect children (glow/shadow/reflection/props3d/...)
   * located in the EG_RPrBase extension slot at the end of rPr. Low-frequency
   * complex subtrees kept verbatim for round-trip fidelity; the rPr backbone
   * stays structured and editable.
   */
  w14RawXml?: string;
}

/**
 * Options for configuring run properties.
 *
 * Extends RunStylePropertiesOptions with a style reference.
 */
export type RunPropertiesOptions = {
  /** Reference to a character style by name */
  style?: string;
} & RunStylePropertiesOptions;

/**
 * Options for run properties change tracking.
 *
 * Used for revision tracking when run properties have been modified.
 */
export type RunPropertiesChangeOptions = {} & RunPropertiesOptions & ChangedAttributesProperties;

export type ParagraphRunPropertiesOptions = {
  insertion?: ChangedAttributesProperties;
  deletion?: ChangedAttributesProperties;
} & RunPropertiesOptions;
