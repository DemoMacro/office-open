/**
 * Chart data and collection for document generation.
 *
 * @module
 */

export interface ChartData {
  key: string;
  chartSpaceXml: string;
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
