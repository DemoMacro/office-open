import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

/**
 * a:graphic > a:graphicData — DrawingML graphic wrapper for table.
 * Lazy: stores table reference, builds IXmlableObject in prepForXml.
 */
export class Graphic extends BaseXmlComponent {
    private readonly table: BaseXmlComponent;

    public constructor(table: BaseXmlComponent) {
        super("a:graphic");
        this.table = table;
    }

    public override prepForXml(context: IContext): IXmlableObject {
        const tableObj = this.table.prepForXml(context);
        const graphicDataChildren: IXmlableObject[] = [
            { _attr: { uri: "http://schemas.openxmlformats.org/drawingml/2006/table" } },
        ];
        if (tableObj) graphicDataChildren.push(tableObj);
        return {
            "a:graphic": [{ "a:graphicData": graphicDataChildren }],
        };
    }
}
