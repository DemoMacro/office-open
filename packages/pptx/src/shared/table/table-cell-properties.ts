import type { FillOptions, UniversalMeasure } from "@office-open/core";
import type { OutlineOptions, PresetDash } from "@office-open/core/drawing";

export interface CellBorderOptions {
  width?: number | UniversalMeasure;
  /** Hex color string sugar, or a full fill (scheme colors, gradients). */
  color?: string | FillOptions;
  /** Dash pattern sugar (ST_PresetLineDashVal) — same value set as
   *  `OutlineOptions.dash`; "solid" = a continuous line. */
  dashStyle?: PresetDash;
  /**
   * Full line properties (CT_LineProperties) for round-trip fidelity — joins,
   * line ends, compound/cap/alignment. Takes precedence over the sugar fields.
   */
  outline?: OutlineOptions;
}
