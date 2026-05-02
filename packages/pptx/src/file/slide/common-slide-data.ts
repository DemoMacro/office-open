import type { Background } from "@file/background/background";
import { ShapeTree } from "@file/shape-tree/shape-tree";
import { XmlComponent } from "@file/xml-components";

/**
 * p:cSld — Common slide data (background + shape tree).
 */
export class CommonSlideData extends XmlComponent {
    public constructor(children: readonly XmlComponent[], background?: Background) {
        super("p:cSld");
        if (background) {
            this.root.push(background);
        }
        this.root.push(new ShapeTree(children));
    }
}
