import type { UniversalMeasure } from "@office-open/core";

export interface CellBorderOptions {
  width?: number | UniversalMeasure;
  color?: string;
  dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}
