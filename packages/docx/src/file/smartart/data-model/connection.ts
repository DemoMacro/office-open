/**
 * dgm:cxn — SmartArt data model connection (edge).
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { chartAttr } from "@file/xml-components";

/**
 * CT_Cxn — a single connection in the data model.
 */
export class Connection extends XmlComponent {
    public constructor(
        modelId: number,
        srcId: number,
        destId: number,
        type: string = "parOf",
        srcOrd: number = 0,
        destOrd: number = 0,
    ) {
        super("dgm:cxn");
        this.root.push(chartAttr({ modelId, srcId, destId, type, srcOrd, destOrd }));
    }
}
