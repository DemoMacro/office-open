import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";
import type { ChartSeriesData } from "../create-chart-type";
import { DataLabels, ErrorBars, Trendline } from "../extensions";
import { createStrRef, createNumRef } from "../series/series-data";

interface BarChartOptions {
  readonly barDirection: "col" | "bar";
  readonly categories: readonly string[];
  readonly series: readonly ChartSeriesData[];
}

export class BarChart extends XmlComponent {
  public constructor(options: BarChartOptions) {
    super("c:barChart");
    this.root.push(wrapEl("c:barDir", chartAttr({ val: options.barDirection })));
    this.root.push(wrapEl("c:grouping", chartAttr({ val: "clustered" })));

    for (let i = 0; i < options.series.length; i++) {
      this.root.push(new BarSeries(i, options.series[i], options.categories));
    }

    this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
    this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
  }
}

class BarSeries extends XmlComponent {
  public constructor(index: number, series: ChartSeriesData, categories: readonly string[]) {
    super("c:ser");
    this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
    this.root.push(wrapEl("c:order", chartAttr({ val: index })));
    this.root.push(new SeriesTx(series.name));
    this.root.push(new EmptyElement("c:spPr"));
    // XSD order: dLbls, trendline, errBars BEFORE cat/val
    if (series.dataLabels) {
      this.root.push(new DataLabels(series.dataLabels));
    }
    if (series.trendlines) {
      for (const tl of series.trendlines) {
        this.root.push(new Trendline(tl));
      }
    }
    if (series.errorBars) {
      this.root.push(new ErrorBars(series.errorBars));
    }
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
