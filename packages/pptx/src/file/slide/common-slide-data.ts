import { ShapeTree } from "@file/shape-tree/shape-tree";
import { XmlComponent } from "@file/xml-components";

/**
 * p:cSld — Common slide data (background + shape tree).
 */
export class CommonSlideData extends XmlComponent {
    public constructor(children: readonly XmlComponent[]) {
        super("p:cSld");
        this.root.push(new ShapeTree(children));
    }
}
