import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";

import type { IChartSeriesData } from "../create-chart-type";
import { createStrRef, createNumRef } from "../series/series-data";

interface IPieChartOptions {
    readonly categories: readonly string[];
    readonly series: readonly IChartSeriesData[];
}

export class PieChart extends XmlComponent {
    public constructor(options: IPieChartOptions) {
        super("c:pieChart");
        this.root.push(wrapEl("c:varyColors", chartAttr({ val: true })));

        for (let i = 0; i < options.series.length; i++) {
            this.root.push(new PieSeries(i, options.series[i], options.categories));
        }
    }
}

class PieSeries extends XmlComponent {
    public constructor(index: number, series: IChartSeriesData, categories: readonly string[]) {
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
