import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";
import type { ChartSeriesData } from "../create-chart-type";
import { createStrRef, createNumRef } from "../series/series-data";

interface StockChartOptions {
  readonly categories: readonly string[];
  readonly series: readonly ChartSeriesData[];
}

export class StockChart extends XmlComponent {
  public constructor(options: StockChartOptions) {
    super("c:stockChart");

    for (let i = 0; i < options.series.length; i++) {
      this.root.push(new StockSeries(i, options.series[i]));
    }

    this.root.push(new DropLines());
    this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
    this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
  }
}

class StockSeries extends XmlComponent {
  public constructor(index: number, series: ChartSeriesData) {
    super("c:ser");
    this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
    this.root.push(wrapEl("c:order", chartAttr({ val: index })));
    this.root.push(new SeriesTx(series.name));
    this.root.push(new EmptyElement("c:spPr"));
    this.root.push(new SeriesVal(series.values));
  }
}

class SeriesTx extends XmlComponent {
  public constructor(name: string) {
    super("c:tx");
    this.root.push(createStrRef(name));
  }
}

class SeriesVal extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:val");
    this.root.push(createNumRef(values));
  }
}

class DropLines extends XmlComponent {
  public constructor() {
    super("c:dropLines");
    this.root.push(new EmptyElement("c:spPr"));
  }
}
