import { Transform2D } from "@file/drawingml/transform-2d";
import { XmlComponent } from "@file/xml-components";
import { convertPixelsToEmu } from "@office-open/core";

import { Graphic } from "./graphic";
import { GraphicFrameNonVisual } from "./graphic-frame-non-visual";
import type { ITableOptions } from "./table";
import { Table } from "./table";

export interface ITableFrameOptions extends ITableOptions {
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
export class TableFrame extends XmlComponent {
    public constructor(options: ITableFrameOptions) {
        super("p:graphicFrame");

        this.root.push(new GraphicFrameNonVisual());

        this.root.push(
            new Transform2D(
                {
                    x: convertPixelsToEmu(options.x ?? 0),
                    y: convertPixelsToEmu(options.y ?? 0),
                    width: convertPixelsToEmu(options.width ?? 0),
                    height: convertPixelsToEmu(options.height ?? 0),
                },
                "p",
            ),
        );

        const table = new Table(options);
        this.root.push(new Graphic(table));
    }
}
