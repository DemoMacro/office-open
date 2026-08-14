import type { FillOptions, UniversalMeasure } from "@office-open/core";

export interface CellBorderOptions {
  width?: number | UniversalMeasure;
  /** Hex color string sugar, or a full fill (scheme colors, gradients). */
  color?: string | FillOptions;
  dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}
