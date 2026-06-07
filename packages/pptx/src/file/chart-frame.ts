import type { ChartSpaceOptions } from "@office-open/core/chart";

export interface ChartOptions extends ChartSpaceOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}
