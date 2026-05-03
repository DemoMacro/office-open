/**
 * ChartSpace — root element for chart XML parts (c:chartSpace).
 *
 * @module
 */
import { EmptyElement, XmlComponent } from "@file/xml-components";
import { chartAttr, wrapEl } from "@file/xml-components";

import { CatAx, ValAx } from "./axes";
import { createChartType } from "./chart-types/create-chart-type";
import type { ChartType } from "./chart-types/create-chart-type";
import { ChartTitle } from "./title";

export interface IChartSeriesData {
    readonly name: string;
    readonly values: readonly number[];
}

export interface IChartSpaceOptions {
    readonly title?: string;
    readonly type: ChartType;
    readonly categories: readonly string[];
    readonly series: readonly IChartSeriesData[];
    readonly showLegend?: boolean;
    readonly style?: number;
}

/**
 * c:chartSpace — root element for chart XML parts.
 */
export class ChartSpace extends XmlComponent {
    public constructor(options: IChartSpaceOptions) {
        super("c:chartSpace");

        // Namespaces on root element
        this.root.push(
            chartAttr({
                "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }),
        );

        const chart = new ChartContainer();

        if (options.title) {
            chart["root"].push(new ChartTitle(options.title));
        }

        chart["root"].push(new EmptyElement("c:autoTitleDeleted"));

        const plotArea = new PlotArea();
        plotArea["root"].push(
            createChartType({
                categories: options.categories,
                series: options.series,
                type: options.type,
            }),
        );

        if (options.type !== "pie") {
            if (options.type === "scatter") {
                plotArea["root"].push(new ValAx(10, 20));
                plotArea["root"].push(new ValAx(20, 10));
            } else {
                plotArea["root"].push(new CatAx(10, 20));
                plotArea["root"].push(new ValAx(20, 10));
            }
        }

        chart["root"].push(plotArea);

        if (options.showLegend !== false) {
            chart["root"].push(createLegend());
        }

        this.root.push(chart);

        if (options.style !== undefined) {
            this.root.push(new ChartStyle(options.style));
        }
    }
}

class ChartContainer extends XmlComponent {
    public constructor() {
        super("c:chart");
    }
}

class PlotArea extends XmlComponent {
    public constructor() {
        super("c:plotArea");
    }
}

function createLegend(): XmlComponent {
    const legend = new (class extends XmlComponent {
        public constructor() {
            super("c:legend");
        }
    })();
    legend["root"].push(wrapEl("c:legendPos", chartAttr({ val: "b" })));
    legend["root"].push(new EmptyElement("c:layout"));
    legend["root"].push(wrapEl("c:overlay", chartAttr({ val: 0 })));
    return legend;
}

class ChartStyle extends XmlComponent {
    public constructor(val: number) {
        super("c:style");
        this.root.push(chartAttr({ val: String(val) }));
    }
}
