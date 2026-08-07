/**
 * Chart data and collection for document generation.
 *
 * @module
 */

import type { ChartSeriesData, BubbleSeriesData } from "./types";

/** Original chart data for generating embedded Excel workbook */
export interface ChartOriginalData {
  categories: readonly string[];
  series: readonly ChartSeriesData[] | readonly BubbleSeriesData[];
}

export interface ChartData {
  key: string;
  chartSpaceXml: string;
  /** Original chart data for generating embedded Excel workbook (enables chart editing in Word) */
  chartData?: ChartOriginalData;
}

export class ChartCollection {
  private map: Map<string, ChartData>;

  public constructor() {
    this.map = new Map<string, ChartData>();
  }

  public addChart(key: string, chartData: ChartData): void {
    this.map.set(key, chartData);
  }

  public get array(): ChartData[] {
    return [...this.map.values()];
  }
}
