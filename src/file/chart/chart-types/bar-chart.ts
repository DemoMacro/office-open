/**
 * Bar/Column chart XML component (c:barChart).
 *
 * @module
 */
import { EmptyElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

import type { IChartSeriesData } from "../chart-space";
import { createStrRef, createNumRef } from "../series/series-data";

interface IBarChartOptions {
    readonly barDirection: "col" | "bar";
    readonly categories: readonly string[];
    readonly series: readonly IChartSeriesData[];
}

/**
 * CT_BarChart — bar or column chart type.
 */
export class BarChart extends XmlComponent {
    public constructor(options: IBarChartOptions) {
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
