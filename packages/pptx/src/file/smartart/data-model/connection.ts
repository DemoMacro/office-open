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
        modelId: string,
        srcId: string,
        destId: string,
        type?: string,
        srcOrd: number = 0,
        destOrd: number = 0,
        parTransId?: string,
        sibTransId?: string,
    ) {
        super("dgm:cxn");
        const attrs: Record<string, string | number> = { modelId, srcId, destId, srcOrd, destOrd };
        if (type) attrs.type = type;
        if (parTransId) attrs.parTransId = parTransId;
        if (sibTransId) attrs.sibTransId = sibTransId;
        this.root.push(chartAttr(attrs));
    }
}
