import { GroupShapeProperties } from "@file/drawingml/group-shape-properties";
import type { IGroupTransform2DOptions } from "@file/drawingml/group-transform-2d";
import { GroupShapeNonVisualProperties } from "@file/shape-tree/group-shape-non-visual";
import { XmlComponent } from "@file/xml-components";
import { pixelsToEmus } from "@util/types";

export interface IGroupShapeOptions {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly rotation?: number;
    readonly flipH?: boolean;
    readonly children: readonly XmlComponent[];
}

/**
 * p:grpSp — Group shape containing child shapes.
 *
 * Children can be Shape, Picture, ChartFrame, or nested GroupShape.
 * x/y/width/height define the group's bounding box in EMUs (pixel values auto-converted).
 */
export class GroupShape extends XmlComponent {
    private static nextId = 100;

    public constructor(options: IGroupShapeOptions) {
        super("p:grpSp");

        const id = GroupShape.nextId++;
        const name = `Group ${id}`;

        this.root.push(new GroupShapeNonVisualProperties(id, name));

        const transformOptions: IGroupTransform2DOptions = {
            x: options.x !== undefined ? pixelsToEmus(options.x) : undefined,
            y: options.y !== undefined ? pixelsToEmus(options.y) : undefined,
            width: options.width !== undefined ? pixelsToEmus(options.width) : undefined,
            height: options.height !== undefined ? pixelsToEmus(options.height) : undefined,
            flipH: options.flipH,
            rotation: options.rotation,
        };
        this.root.push(new GroupShapeProperties(transformOptions));

        this.root.push(...options.children);
    }
}
