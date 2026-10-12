import type { ChartSourceRelationship } from "./chart-collection";

export interface ChartExPartData {
  key: string;
  chartExXml: string;
  /** Source package path for sparse part names; round-trip only. */
  sourcePath?: string;
  /** Source ChartEx-part relationships preserved with their original rIds. */
  sourceRelationships?: ChartSourceRelationship[];
}

export class ChartExCollection {
  private map = new Map<string, ChartExPartData>();
  private cachedArray: ChartExPartData[] | undefined;

  public addChartEx(key: string, chartExData: ChartExPartData): void {
    this.map.set(key, chartExData);
    this.cachedArray = undefined;
  }

  public get array(): ChartExPartData[] {
    if (this.cachedArray === undefined) this.cachedArray = [...this.map.values()];
    return this.cachedArray;
  }
}
