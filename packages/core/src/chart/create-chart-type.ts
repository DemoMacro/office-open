import { AreaChart } from "./chart-types/area-chart";
import { BarChart } from "./chart-types/bar-chart";
import { LineChart } from "./chart-types/line-chart";
import { PieChart } from "./chart-types/pie-chart";
import { ScatterChart } from "./chart-types/scatter-chart";

export interface IChartSeriesData {
    readonly name: string;
    readonly values: readonly number[];
}

export type ChartType = "column" | "bar" | "line" | "pie" | "area" | "scatter";

export interface IChartTypeOptions {
    readonly type: ChartType;
    readonly series: readonly IChartSeriesData[];
    readonly categories: readonly string[];
}

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
            throw new Error(`Unsupported chart type: ${(options as IChartTypeOptions).type}`);
    }
};
