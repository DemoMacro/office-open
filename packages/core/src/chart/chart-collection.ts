import type { XmlComponent } from "../xml-components";

export interface ChartData {
  readonly key: string;
  readonly chartSpace: XmlComponent;
}

export class ChartCollection {
  private readonly map: Map<string, ChartData>;

  public constructor() {
    this.map = new Map<string, ChartData>();
  }

  public addChart(key: string, chartData: ChartData): void {
    this.map.set(key, chartData);
  }

  public get Array(): readonly ChartData[] {
    return [...this.map.values()];
  }
}
