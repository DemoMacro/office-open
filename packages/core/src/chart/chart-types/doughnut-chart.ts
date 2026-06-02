import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";
import type { ChartSeriesData } from "../create-chart-type";
import { createStrRef, createNumRef } from "../series/series-data";

interface DoughnutChartOptions {
  readonly holeSize?: number;
  readonly series: readonly ChartSeriesData[];
  readonly categories: readonly string[];
}

export class DoughnutChart extends XmlComponent {
  public constructor(options: DoughnutChartOptions) {
    super("c:doughnutChart");
    this.root.push(wrapEl("c:varyColors", chartAttr({ val: true })));

    if (options.holeSize !== undefined) {
      this.root.push(wrapEl("c:firstSliceAng", chartAttr({ val: 0 })));
    }

    for (let i = 0; i < options.series.length; i++) {
      this.root.push(new DoughnutSeries(i, options.series[i], options.categories));
    }
  }
}

class DoughnutSeries extends XmlComponent {
  public constructor(index: number, series: ChartSeriesData, categories: readonly string[]) {
    super("c:ser");
    this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
    this.root.push(wrapEl("c:order", chartAttr({ val: index })));
    this.root.push(new SeriesTx(series.name));
    this.root.push(new EmptyElement("c:spPr"));
    this.root.push(new SeriesCat(categories));
    this.root.push(new SeriesVal(series.values));
  }
}

class SeriesTx extends XmlComponent {
  public constructor(name: string) {
    super("c:tx");
    this.root.push(createStrRef(name));
  }
}

class SeriesCat extends XmlComponent {
  public constructor(categories: readonly string[]) {
    super("c:cat");
    this.root.push(createStrRef(categories));
  }
}

class SeriesVal extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:val");
    this.root.push(createNumRef(values));
  }
}
