import type { FillOptions, UniversalMeasure } from "@office-open/core";
import type { OutlineOptions } from "@office-open/core/drawing";

export interface CellBorderOptions {
  width?: number | UniversalMeasure;
  /** Hex color string sugar, or a full fill (scheme colors, gradients). */
  color?: string | FillOptions;
  dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
  /**
   * Full line properties (CT_LineProperties) for round-trip fidelity — joins,
   * line ends, compound/cap/alignment. Takes precedence over the sugar fields.
   */
  outline?: OutlineOptions;
}
