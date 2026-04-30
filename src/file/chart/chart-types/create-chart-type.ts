/**
 * Chart type factory module.
 *
 * @module
 */
import type { IChartSeriesData } from "../chart-space";
import { AreaChart } from "./area-chart";
import { BarChart } from "./bar-chart";
import { LineChart } from "./line-chart";
import { PieChart } from "./pie-chart";
import { ScatterChart } from "./scatter-chart";

export type ChartType = "column" | "bar" | "line" | "pie" | "area" | "scatter";

export interface IChartTypeOptions {
    readonly type: ChartType;
    readonly series: readonly IChartSeriesData[];
    readonly categories: readonly string[];
}

/**
 * Creates the appropriate chart type XML component.
 */
export const createChartType = (options: IChartTypeOptions) => {
    switch (options.type) {
        case "column":
        case "bar":
            return new BarChart({
                barDirection: options.type === "column" ? "col" : "bar",
                categories: options.categories,
                series: options.series,
            });
        case "line":
            return new LineChart({
                categories: options.categories,
                series: options.series,
            });
        case "pie":
            return new PieChart({
                categories: options.categories,
                series: options.series,
            });
        case "area":
            return new AreaChart({
                categories: options.categories,
                series: options.series,
            });
        case "scatter":
            return new ScatterChart({
                categories: options.categories,
                series: options.series,
            });
        default:
            throw new Error(`Chart type "${(options as any).type}" is not supported`);
    }
};
