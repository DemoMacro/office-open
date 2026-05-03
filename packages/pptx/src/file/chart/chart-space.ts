import { BuilderElement, XmlComponent } from "@file/xml-components";
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

export class ChartSpace extends XmlComponent {
    public constructor(options: IChartSpaceOptions) {
        super("c:chartSpace");

        this.root.push(
            chartAttr({
                "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }),
        );

        this.root.push(wrapEl("c:date1904", chartAttr({ val: 0 })));
        this.root.push(wrapEl("c:lang", chartAttr({ val: "en-US" })));
        this.root.push(wrapEl("c:roundedCorners", chartAttr({ val: 0 })));

        const chart = new ChartContainer();

        if (options.title) {
            chart["root"].push(new ChartTitle(options.title));
        }

        chart["root"].push(wrapEl("c:autoTitleDeleted", chartAttr({ val: 0 })));

        const plotArea = new PlotArea();
        plotArea["root"].push(new BuilderElement({ name: "c:layout" }));

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
        this.root.push(createChartSpPr());
        this.root.push(createChartTxPr());

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
    legend["root"].push(new BuilderElement({ name: "c:layout" }));
    legend["root"].push(wrapEl("c:overlay", chartAttr({ val: 0 })));
    legend["root"].push(createNoFillSpPr());
    legend["root"].push(createTxPr());
    return legend;
}

function createNoFillSpPr(): XmlComponent {
    const spPr = new (class extends XmlComponent {
        public constructor() {
            super("c:spPr");
        }
    })();
    spPr["root"].push(
        new (class extends XmlComponent {
            public constructor() {
                super("a:noFill");
            }
        })(),
    );
    spPr["root"].push(
        new (class extends XmlComponent {
            public constructor() {
                super("a:ln");
                this.root.push(
                    new (class extends XmlComponent {
                        public constructor() {
                            super("a:noFill");
                        }
                    })(),
                );
            }
        })(),
    );
    spPr["root"].push(new BuilderElement({ name: "a:effectLst" }));
    return spPr;
}

function createChartSpPr(): XmlComponent {
    const spPr = new (class extends XmlComponent {
        public constructor() {
            super("c:spPr");
        }
    })();
    spPr["root"].push(
        new (class extends XmlComponent {
            public constructor() {
                super("a:noFill");
            }
        })(),
    );
    spPr["root"].push(
        new (class extends XmlComponent {
            public constructor() {
                super("a:ln");
                this.root.push(
                    new (class extends XmlComponent {
                        public constructor() {
                            super("a:noFill");
                        }
                    })(),
                );
            }
        })(),
    );
    spPr["root"].push(new BuilderElement({ name: "a:effectLst" }));
    return spPr;
}

function createChartTxPr(): XmlComponent {
    const txPr = new (class extends XmlComponent {
        public constructor() {
            super("c:txPr");
        }
    })();
    txPr["root"].push(new BuilderElement({ name: "a:bodyPr" }));
    txPr["root"].push(new BuilderElement({ name: "a:lstStyle" }));
    txPr["root"].push(createTextParagraph());
    return txPr;
}

function createTxPr(): XmlComponent {
    const txPr = new (class extends XmlComponent {
        public constructor() {
            super("c:txPr");
        }
    })();
    txPr["root"].push(createBodyPr());
    txPr["root"].push(new BuilderElement({ name: "a:lstStyle" }));
    txPr["root"].push(createTextParagraph());
    return txPr;
}

function createBodyPr(): XmlComponent {
    return new BuilderElement({
        name: "a:bodyPr",
        attributes: {
            rot: { key: "rot", value: "0" },
            spcFirstLastPara: { key: "spcFirstLastPara", value: "1" },
            vertOverflow: { key: "vertOverflow", value: "ellipsis" },
            vert: { key: "vert", value: "horz" },
            wrap: { key: "wrap", value: "square" },
            anchor: { key: "anchor", value: "ctr" },
            anchorCtr: { key: "anchorCtr", value: "1" },
        },
    });
}

function createTextParagraph(): XmlComponent {
    const p = new (class extends XmlComponent {
        public constructor() {
            super("a:p");
        }
    })();
    const pPr = new (class extends XmlComponent {
        public constructor() {
            super("a:pPr");
        }
    })();
    pPr["root"].push(new BuilderElement({ name: "a:defRPr" }));
    p["root"].push(pPr);
    p["root"].push(
        new BuilderElement({
            name: "a:endParaRPr",
            attributes: { lang: { key: "lang", value: "en-US" } },
        }),
    );
    return p;
}

class ChartStyle extends XmlComponent {
    public constructor(val: number) {
        super("c:style");
        this.root.push(chartAttr({ val: String(val) }));
    }
}
