import type { UniversalMeasure } from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core/chart";

export interface ChartOptions extends ChartSpaceOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
}
