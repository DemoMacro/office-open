import type { ChartSpaceOptions } from "@office-open/core/chart";

export interface ChartOptions extends ChartSpaceOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}
