/**
 * Chart data and collection for document generation.
 *
 * @module
 */
import type { DataType } from "../util/data-type";

export interface ChartData {
  key: string;
  chartSpaceXml: string;
  /**
   * Embedded workbook for c:externalData (round-trip). The compiler emits the
   * chart part's own rels plus the word/embeddings part so the rId resolves.
   */
  embedding?: {
    relationshipId: string;
    fileName: string;
    data: DataType;
  };
  /**
   * User-shapes part behind c:userShapes (both directions). The compiler
   * emits the chart part's own rels entry plus the shapes part body.
   */
  userShapes?: {
    relationshipId: string;
    xml: string;
  };
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
