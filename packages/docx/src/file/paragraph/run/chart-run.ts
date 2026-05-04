/**
 * ChartRun module for embedding charts in WordprocessingML documents.
 *
 * Charts are stored as independent XML parts (word/charts/chart{n}.xml)
 * and referenced from document.xml via drawing relationships.
 *
 * @module
 */
import { ChartSpace } from "@file/chart/chart-space";
import { Drawing } from "@file/drawing";
import type { IFloating } from "@file/drawing";
import type { DocPropertiesOptions } from "@file/drawing/doc-properties/doc-properties";
import type { IMediaTransformation } from "@file/media";
import { createTransformation } from "@file/media";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { Run } from "../run";

/**
 * Options for creating a ChartRun.
 *
 * @publicApi
 */
export interface IChartOptions {
    /** Chart type */
    readonly type: "column" | "bar" | "line" | "pie" | "area" | "scatter";
    /** Chart data */
    readonly data: {
        readonly categories: readonly string[];
        readonly series: readonly { name: string; values: readonly number[] }[];
    };
    /** Display dimensions */
    readonly transformation: IMediaTransformation;
    /** Floating positioning */
    readonly floating?: IFloating;
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
    private readonly chartOptions: IChartOptions;
    private readonly chartKey: string;

    public constructor(options: IChartOptions) {
        super({});
        this.chartOptions = options;

        // Generate a unique key based on chart data
        const hash = this.hashChartData(options);
        this.chartKey = `chart_${hash}`;

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

    public prepForXml(context: IContext): IXmlableObject | undefined {
        // Register chart with the file's chart collection
        const chartSpace = new ChartSpace({
            categories: this.chartOptions.data.categories,
            series: this.chartOptions.data.series,
            showLegend: this.chartOptions.showLegend,
            style: this.chartOptions.style,
            title: this.chartOptions.title,
            type: this.chartOptions.type,
        });

        context.file.Charts.addChart(this.chartKey, {
            chartSpace,
            key: this.chartKey,
        });

        return super.prepForXml(context);
    }

    private hashChartData(options: IChartOptions): number {
        const data = `${options.type}:${JSON.stringify(options.data)}`;
        let hash = 0;
        for (let i = 0; i < data.length; i++) {
            const char = data.charCodeAt(i);
            hash = ((hash << 5) - hash + char) | 0;
        }
        return Math.abs(hash);
    }
}
