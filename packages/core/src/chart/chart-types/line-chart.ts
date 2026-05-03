import { EmptyElement, XmlComponent, chartAttr, wrapEl } from "../../xml-components";

import type { IChartSeriesData } from "../create-chart-type";
import { createStrRef, createNumRef } from "../series/series-data";

interface ILineChartOptions {
    readonly categories: readonly string[];
    readonly series: readonly IChartSeriesData[];
}

export class LineChart extends XmlComponent {
    public constructor(options: ILineChartOptions) {
        super("c:lineChart");
        this.root.push(wrapEl("c:grouping", chartAttr({ val: "standard" })));

        for (let i = 0; i < options.series.length; i++) {
            this.root.push(new LineSeries(i, options.series[i], options.categories));
        }

        this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
        this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
    }
}

class LineSeries extends XmlComponent {
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
