import type { NonVisualDrawingPropertiesOptions, UniversalMeasure } from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core/chart";

/**
 * Chart frame options for pptx slides (p:graphicFrame referencing a c:chartSpace).
 * The chart payload comes from {@link ChartSpaceOptions} (core shared model,
 * identical XML across packages); the cNvPr fields from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 *
 * `title` is the chart title (c:title) inherited from ChartSpaceOptions — the
 * cNvPr @title attribute is deliberately not exposed here so one JSON key
 * cannot mean two XML attributes (alt text goes in `description`).
 */
export interface ChartOptions
  extends ChartSpaceOptions, Omit<NonVisualDrawingPropertiesOptions, "title"> {
  /** Chart frame id (p:cNvPr @id). Auto-generated if omitted. */
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Pre-generated chart key (e.g. "chart_2048"). If omitted, auto-generated. */
  chartKey?: string;
}
