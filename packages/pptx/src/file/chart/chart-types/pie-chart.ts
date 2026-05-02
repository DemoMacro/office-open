import { BuilderElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

import type { IChartSeriesData } from "../chart-space";
import { createStrRef, createNumRef } from "../series/series-data";

interface IPieChartOptions {
    readonly categories: readonly string[];
    readonly series: readonly IChartSeriesData[];
}

export class PieChart extends XmlComponent {
    public constructor(options: IPieChartOptions) {
        super("c:pieChart");
        this.root.push(wrapEl("c:varyColors", chartAttr({ val: 1 })));
        const series = options.series[0];
        if (series) {
            this.root.push(new PieSeries(series, options.categories));
        }
    }
}

class PieSeries extends XmlComponent {
    public constructor(series: IChartSeriesData, categories: readonly string[]) {
        super("c:ser");
        this.root.push(wrapEl("c:idx", chartAttr({ val: 0 })));
        this.root.push(wrapEl("c:order", chartAttr({ val: 0 })));
        this.root.push(new SeriesTx(series.name));
        this.root.push(new SeriesCat(categories));
        this.root.push(new SeriesVal(series.values));
        for (let i = 0; i < categories.length; i++) {
            this.root.push(wrapEl("c:dPt", chartAttr({ idx: i })));
        }
        this.root.push(new BuilderElement({ name: "c:spPr" }));
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
