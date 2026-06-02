/**
 * ChartRun module for embedding charts in WordprocessingML documents.
 *
 * Charts are stored as independent XML parts (word/charts/chart{n}.xml)
 * and referenced from document.xml via drawing relationships.
 *
 * @module
 */
import { ChartSpace } from "@file/chart/chart-space";
import type { ChartType, ChartSeriesData } from "@file/chart/chart-space";
import { Drawing } from "@file/drawing";
import type { Floating } from "@file/drawing";
import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { MediaTransformation } from "@file/media";
import { createTransformation } from "@file/media";
import type { Context, IXmlableObject } from "@file/xml-components";

import { Run } from "../run";

let nextChartId = 1;

/**
 * Options for creating a ChartRun.
 *
 * @publicApi
 */
export interface ChartOptions {
  /** Chart type */
  readonly type: ChartType;
  /** Chart data — for bubble charts, series must include xValues/yValues/bubbleSize instead of values */
  readonly data: {
    readonly categories?: readonly string[];
    readonly series: readonly ChartSeriesData[];
  };
  /** Enable 3D rendering for applicable chart types */
  readonly threeD?: boolean;
  /** Display dimensions */
  readonly transformation: MediaTransformation;
  /** Floating positioning */
  readonly floating?: Floating;
  /** Alternative text for accessibility */
  readonly altText?: DocPropertiesOptions;
  /** Chart title */
  readonly title?: string;
  /** Show legend (default: true) */
  readonly showLegend?: boolean;
  /** Chart style preset (2-48) */
  readonly style?: number;
}

/**
 * Represents an embedded chart in a WordprocessingML document.
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * new ChartRun({
 *   type: "column",
 *   data: {
 *     categories: ["Q1", "Q2", "Q3", "Q4"],
 *     series: [
 *       { name: "Sales", values: [100, 200, 300, 400] },
 *     ],
 *   },
 *   transformation: { width: 500, height: 300 },
 *   title: "Quarterly Sales",
 * });
 * ```
 */
export class ChartRun extends Run {
  private readonly chartOptions: ChartOptions;
  private readonly chartKey: string;

  public constructor(options: ChartOptions) {
    super({});
    this.chartOptions = options;

    this.chartKey = `chart_${nextChartId++}`;

    // Create media data for the drawing system
    const mediaData = {
      chartKey: this.chartKey,
      transformation: createTransformation(options.transformation),
      type: "chart" as const,
    };

    const drawing = new Drawing(mediaData, {
      docProperties: options.altText,
      floating: options.floating,
    });

    this.extraChildren.push(drawing);
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    // Register chart with the file's chart collection
    const chartSpace = new ChartSpace({
      categories: this.chartOptions.data.categories,
      series: this.chartOptions.data.series,
      showLegend: this.chartOptions.showLegend,
      style: this.chartOptions.style,
      title: this.chartOptions.title,
      type: this.chartOptions.type,
      threeD: this.chartOptions.threeD,
    });

    context.file.charts.addChart(this.chartKey, {
      chartSpace,
      key: this.chartKey,
    });

    return super.prepForXml(context);
  }
}
