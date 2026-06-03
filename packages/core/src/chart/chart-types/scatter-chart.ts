import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";
import type { ChartSeriesData } from "../create-chart-type";
import { DataLabels, ErrorBars, Trendline } from "../extensions";
import { createStrRef, createNumRef } from "../series/series-data";

interface ScatterChartOptions {
  readonly categories: readonly string[];
  readonly series: readonly ChartSeriesData[];
}

export class ScatterChart extends XmlComponent {
  public constructor(options: ScatterChartOptions) {
    super("c:scatterChart");
    this.root.push(wrapEl("c:scatterStyle", chartAttr({ val: "line" })));

    for (let i = 0; i < options.series.length; i++) {
      this.root.push(new ScatterSeries(i, options.series[i], options.categories));
    }

    this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
    this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
  }
}

class ScatterSeries extends XmlComponent {
  public constructor(index: number, series: ChartSeriesData, categories: readonly string[]) {
    super("c:ser");
    this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
    this.root.push(wrapEl("c:order", chartAttr({ val: index })));
    this.root.push(new SeriesTx(series.name));
    this.root.push(new EmptyElement("c:spPr"));
    // XSD order: dLbls, trendline, errBars BEFORE xVal/yVal
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
    this.root.push(new SeriesXVal(categories));
    this.root.push(new SeriesYVal(series.values));
  }
}

class SeriesTx extends XmlComponent {
  public constructor(name: string) {
    super("c:tx");
    this.root.push(createStrRef(name));
  }
}

class SeriesXVal extends XmlComponent {
  public constructor(categories: readonly string[]) {
    super("c:xVal");
    this.root.push(createStrRef(categories));
  }
}

class SeriesYVal extends XmlComponent {
  public constructor(values: readonly number[]) {
    super("c:yVal");
    this.root.push(createNumRef(values));
  }
}
