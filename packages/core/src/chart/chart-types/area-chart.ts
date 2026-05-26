import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";
import type { ChartSeriesData } from "../create-chart-type";
import { createStrRef, createNumRef } from "../series/series-data";

interface AreaChartOptions {
  readonly categories: readonly string[];
  readonly series: readonly ChartSeriesData[];
}

export class AreaChart extends XmlComponent {
  public constructor(options: AreaChartOptions) {
    super("c:areaChart");
    this.root.push(wrapEl("c:grouping", chartAttr({ val: "standard" })));

    for (let i = 0; i < options.series.length; i++) {
      this.root.push(new AreaSeries(i, options.series[i], options.categories));
    }

    this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
    this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
  }
}

class AreaSeries extends XmlComponent {
  public constructor(index: number, series: ChartSeriesData, categories: readonly string[]) {
    super("c:ser");
    this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
    this.root.push(wrapEl("c:order", chartAttr({ val: index })));
    this.root.push(new SeriesTx(series.name));
    this.root.push(new SeriesCat(categories));
    this.root.push(new SeriesVal(series.values));
    this.root.push(new EmptyElement("c:spPr"));
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
