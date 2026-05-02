import { GroupShapeProperties } from "@file/drawingml/group-shape-properties";
import { XmlComponent } from "@file/xml-components";

import { GroupShapeNonVisualProperties } from "./group-shape-non-visual";

/**
 * p:spTree — Shape tree containing all shapes on a slide.
 */
export class ShapeTree extends XmlComponent {
    public constructor(children: readonly XmlComponent[]) {
        super("p:spTree");
        this.root.push(new GroupShapeNonVisualProperties());
        this.root.push(new GroupShapeProperties());
        this.root.push(...children);
    }
}
