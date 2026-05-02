import { Transform2D } from "@file/drawingml/transform-2d";
import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

import { ChartCollection } from "./chart-collection";
import type { IChartSpaceOptions } from "./chart-space";
import { ChartSpace } from "./chart-space";

let nextChartFrameId = 2048;

export interface IChartFrameOptions extends IChartSpaceOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
}

/**
 * p:graphicFrame — Slide-level graphic frame wrapping a chart.
 *
 * The chart is stored as a separate part (ppt/charts/chart{n}.xml)
 * and referenced via a relationship ID placeholder {chart:key}.
 */
export class ChartFrame extends XmlComponent {
    private readonly chartOptions: IChartSpaceOptions;
    private readonly chartKey: string;

    public constructor(options: IChartFrameOptions) {
        super("p:graphicFrame");

        this.chartOptions = options;
        this.chartKey = `chart_${nextChartFrameId++}`;

        const id = nextChartFrameId++;
        this.root.push(new GraphicFrameNonVisual(id));
        this.root.push(
            new Transform2D(
                {
                    x: pixelsToEmus(options.x ?? 0),
                    y: pixelsToEmus(options.y ?? 0),
                    width: pixelsToEmus(options.width ?? 0),
                    height: pixelsToEmus(options.height ?? 0),
                },
                "p",
            ),
        );

        this.root.push(new ChartGraphic(this.chartKey));
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const file = context.fileData as { Charts: ChartCollection };
        if (file?.Charts) {
            file.Charts.addChart(this.chartKey, {
                chartSpace: new ChartSpace(this.chartOptions),
                key: this.chartKey,
            });
        }

        return super.prepForXml(context);
    }

    public get ChartKey(): string {
        return this.chartKey;
    }
}

class GraphicFrameNonVisual extends XmlComponent {
    public constructor(id: number) {
        super("p:nvGraphicFramePr");
        this.root.push(
            new BuilderElement({
                name: "p:cNvPr",
                attributes: {
                    id: { key: "id", value: id },
                    name: { key: "name", value: `Chart ${id}` },
                },
            }),
        );
        this.root.push(
            new BuilderElement({
                name: "p:cNvGraphicFramePr",
                children: [
                    new BuilderElement({
                        name: "a:graphicFrameLocks",
                        attributes: { noGrp: { key: "noGrp", value: 1 } },
                    }),
                ],
            }),
        );
        this.root.push(new BuilderElement({ name: "p:nvPr" }));
    }
}

class ChartGraphic extends XmlComponent {
    public constructor(chartKey: string) {
        super("a:graphic");
        this.root.push(new ChartGraphicData(chartKey));
    }
}

class ChartGraphicData extends XmlComponent {
    public constructor(chartKey: string) {
        super("a:graphicData");
        this.root.push(
            new NextAttributeComponent({
                uri: {
                    key: "uri",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            }),
        );
        this.root.push(new ChartRef(chartKey));
    }
}

class ChartRef extends XmlComponent {
    public constructor(chartKey: string) {
        super("c:chart");
        this.root.push(
            new NextAttributeComponent({
                xmlnsC: {
                    key: "xmlns:c",
                    value: "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                xmlnsR: {
                    key: "xmlns:r",
                    value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                },
                rId: { key: "r:id", value: `{chart:${chartKey}}` },
            }),
        );
    }
}
