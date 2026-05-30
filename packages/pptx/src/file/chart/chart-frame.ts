import { Transform2D } from "@file/drawingml/transform-2d";
import { buildAttrObject, BuilderElement, XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { emuPosition } from "@util/position";

import { ChartCollection } from "./chart-collection";
import type { ChartSpaceOptions } from "./chart-space";
import { ChartSpace } from "./chart-space";

let nextChartId = 2048;

export interface ChartOptions extends ChartSpaceOptions {
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
export class Chart extends XmlComponent {
  private readonly chartOptions: ChartSpaceOptions;
  public readonly chartKey: string;

  public constructor(options: ChartOptions) {
    super("p:graphicFrame");

    this.chartOptions = options;
    this.chartKey = `chart_${nextChartId++}`;

    const id = nextChartId++;
    this.root.push(new GraphicFrameNonVisual(id));
    this.root.push(
      new Transform2D(
        {
          ...emuPosition(options),
        },
        "p",
      ),
    );

    this.root.push(new ChartGraphic(this.chartKey));
  }

  /** Register chart data with the File's Chart collection. */
  private registerChart(context: Context): void {
    const file = context.fileData as { charts: ChartCollection };
    if (file?.charts) {
      file.charts.addChart(this.chartKey, {
        chartSpace: new ChartSpace(this.chartOptions),
        key: this.chartKey,
      });
    }
  }

  public override toXml(context: Context): string {
    this.registerChart(context);
    return super.toXml(context);
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
      buildAttrObject({
        uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
      }),
    );
    this.root.push(new ChartRef(chartKey));
  }
}

class ChartRef extends XmlComponent {
  public constructor(chartKey: string) {
    super("c:chart");
    this.root.push(
      buildAttrObject({
        "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "r:id": `{chart:${chartKey}}`,
      }),
    );
  }
}
