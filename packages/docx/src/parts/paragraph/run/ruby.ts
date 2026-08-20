import type { RunOptions } from "./run";

/** Ruby alignment values (ST_RubyAlign). */
export const RubyAlign = {
  CENTER: "center",
  DISTRIBUTE_LETTER: "distributeLetter",
  DISTRIBUTE_SPACE: "distributeSpace",
  LEFT: "left",
  RIGHT: "right",
  RIGHT_VERTICAL: "rightVertical",
} as const;

/** Properties of a ruby annotation (CT_RubyPr). */
export interface RubyPropertiesOptions {
  /** Alignment of the annotation text relative to its base text. */
  alignment: (typeof RubyAlign)[keyof typeof RubyAlign];
  /** Annotation font size in points. */
  fontSize: number;
  /** Vertical offset of the annotation in points. */
  raise: number;
  /** Base-text font size in points. */
  baseFontSize: number;
  /** Language identifier for the annotation text. */
  languageId: string;
  /** Whether the annotation needs recalculation. */
  dirty?: boolean;
}

/** Formatted runs inside ruby text or ruby base content (CT_RubyContent). */
export interface RubyContentOptions {
  children?: (RunOptions | string)[];
}

/** A ruby annotation (CT_Ruby). */
export interface RubyOptions {
  properties: RubyPropertiesOptions;
  /** Annotation text (`w:rt`). */
  text: RubyContentOptions;
  /** Base text being annotated (`w:rubyBase`). */
  base: RubyContentOptions;
}
