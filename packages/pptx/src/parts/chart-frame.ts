import type {
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core/chart";
import type { NvPrPlaceholderOptions } from "@parts/descriptors/graphic-frame";

/**
 * Chart frame (p:graphicFrame → c:chartSpace): chart payload from
 * ChartSpaceOptions (core shared model), cNvPr from
 * NonVisualDrawingPropertiesOptions. `title` is c:title; alt text in
 * `description`.
 */
export interface ChartOptions
  extends
    ChartSpaceOptions,
    Omit<NonVisualDrawingPropertiesOptions, "title">,
    NvPrPlaceholderOptions {
  /** Chart frame id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  /** Frame locking (a:graphicFrameLocks). undefined = fresh default
   * (noGrp="1"); null = empty cNvGraphicFramePr; object = explicit flags. */
  locking?: GraphicFrameLockingOptions | null;

  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Pre-generated chart key (e.g. "chart_2048"). If omitted, auto-generated. */
  chartKey?: string;
}
