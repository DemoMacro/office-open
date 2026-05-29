import { Transform2D } from "@file/drawingml/transform-2d";
import { XmlComponent } from "@file/xml-components";
import { emuPosition } from "@util/position";

import { Graphic } from "./graphic";
import { GraphicFrameNonVisual } from "./graphic-frame-non-visual";
import type { DrawingTableOptions } from "./table";
import { DrawingTable } from "./table";

export interface TableOptions extends DrawingTableOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}

/**
 * p:graphicFrame — Slide-level graphic frame wrapping a table.
 *
 * x/y/width/height accept pixel values, converted to EMUs internally.
 */
export class Table extends XmlComponent {
  public constructor(options: TableOptions) {
    super("p:graphicFrame");

    this.root.push(new GraphicFrameNonVisual());

    this.root.push(
      new Transform2D(
        {
          ...emuPosition(options),
        },
        "p",
      ),
    );

    const table = new DrawingTable(options);
    this.root.push(new Graphic(table));
  }
}
