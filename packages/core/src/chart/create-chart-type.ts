import type { XmlComponent } from "../xml-components";
import { AreaChart } from "./chart-types/area-chart";
import { Area3DChart } from "./chart-types/area3d-chart";
import { BarChart } from "./chart-types/bar-chart";
import { Bar3DChart } from "./chart-types/bar3d-chart";
import { BubbleChart } from "./chart-types/bubble-chart";
import type { BubbleSeriesData } from "./chart-types/bubble-chart";
import { DoughnutChart } from "./chart-types/doughnut-chart";
import { LineChart } from "./chart-types/line-chart";
import { Line3DChart } from "./chart-types/line3d-chart";
import { PieChart } from "./chart-types/pie-chart";
import { Pie3DChart } from "./chart-types/pie3d-chart";
import { RadarChart } from "./chart-types/radar-chart";
import { ScatterChart } from "./chart-types/scatter-chart";
import { StockChart } from "./chart-types/stock-chart";
import { SurfaceChart } from "./chart-types/surface-chart";

export interface ChartSeriesData {
  readonly name: string;
  readonly values: readonly number[];
}

export type ChartType =
  | "column"
  | "bar"
  | "line"
  | "pie"
  | "area"
  | "scatter"
  | "bubble"
  | "doughnut"
  | "radar"
  | "stock"
  | "surface";

export type AxisChartType = Exclude<ChartType, "bubble">;

export interface ChartTypeOptions {
  readonly type: AxisChartType;
  readonly series: readonly ChartSeriesData[];
  readonly categories: readonly string[];
  readonly threeD?: boolean;
}

export type BubbleChartOptions = {
  readonly type: "bubble";
  readonly series: readonly BubbleSeriesData[];
};

export type AllChartTypeOptions = ChartTypeOptions | BubbleChartOptions;

const is3D = (opts: ChartTypeOptions): boolean => opts.threeD === true;

export const createChartType = (options: AllChartTypeOptions): XmlComponent => {
  if (options.type === "bubble") {
    return new BubbleChart({ series: options.series });
  }

  const opts = options as ChartTypeOptions;

  switch (opts.type) {
    case "column":
    case "bar":
      return is3D(opts)
        ? new Bar3DChart({
            barDirection: opts.type === "column" ? "col" : "bar",
            categories: opts.categories,
            series: opts.series,
          })
        : new BarChart({
            barDirection: opts.type === "column" ? "col" : "bar",
            categories: opts.categories,
            series: opts.series,
          });
    case "line":
      return is3D(opts)
        ? new Line3DChart({ categories: opts.categories, series: opts.series })
        : new LineChart({ categories: opts.categories, series: opts.series });
    case "pie":
      return is3D(opts)
        ? new Pie3DChart({ categories: opts.categories, series: opts.series })
        : new PieChart({ categories: opts.categories, series: opts.series });
    case "area":
      return is3D(opts)
        ? new Area3DChart({ categories: opts.categories, series: opts.series })
        : new AreaChart({ categories: opts.categories, series: opts.series });
    case "scatter":
      return new ScatterChart({ categories: opts.categories, series: opts.series });
    case "doughnut":
      return new DoughnutChart({ categories: opts.categories, series: opts.series });
    case "radar":
      return new RadarChart({ categories: opts.categories, series: opts.series });
    case "stock":
      return new StockChart({ categories: opts.categories, series: opts.series });
    case "surface":
      return new SurfaceChart({ categories: opts.categories, series: opts.series });
    default:
      throw new Error(`Unsupported chart type: ${(options as AllChartTypeOptions).type}`);
  }
};
