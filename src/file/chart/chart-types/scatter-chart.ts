/**
 * Scatter chart XML component (c:scatterChart).
 *
 * @module
 */
import { EmptyElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

import type { IChartSeriesData } from "../chart-space";
import { createNumRef } from "../series/series-data";

interface IScatterChartOptions {
    readonly categories: readonly string[];
    readonly series: readonly IChartSeriesData[];
}

export class ScatterChart extends XmlComponent {
    public constructor(options: IScatterChartOptions) {
        super("c:scatterChart");
        for (let i = 0; i < options.series.length; i++) {
            this.root.push(new ScatterSeries(i, options.series[i], options.categories));
        }
        this.root.push(wrapEl("c:axId", chartAttr({ val: 10 })));
        this.root.push(wrapEl("c:axId", chartAttr({ val: 20 })));
        this.root.push(wrapEl("c:axId", chartAttr({ val: 30 })));
    }
}

class ScatterSeries extends XmlComponent {
    public constructor(index: number, series: IChartSeriesData, categories: readonly string[]) {
        super("c:ser");
        this.root.push(wrapEl("c:idx", chartAttr({ val: index })));
        this.root.push(wrapEl("c:order", chartAttr({ val: index })));
        this.root.push(new SeriesTx(series.name));
        this.root.push(new XValues(categories));
        this.root.push(new YValues(series.values));
        this.root.push(new EmptyElement("c:spPr"));
        this.root.push(new Marker());
    }
}

class SeriesTx extends XmlComponent {
    public constructor(_name: string) {
        super("c:tx");
        this.root.push(createNumRef([1]));
    }
}

class XValues extends XmlComponent {
    public constructor(categories: readonly string[]) {
        super("c:xVal");
        const xValues = categories.map((_, i) => i + 1);
        this.root.push(createNumRef(xValues));
    }
}

class YValues extends XmlComponent {
    public constructor(values: readonly number[]) {
        super("c:yVal");
        this.root.push(createNumRef(values));
    }
}

class Marker extends XmlComponent {
    public constructor() {
        super("c:marker");
        this.root.push(wrapEl("c:symbol", chartAttr({ val: "circle" })));
        this.root.push(new EmptyElement("c:size"));
    }
}
