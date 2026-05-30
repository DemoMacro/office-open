import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

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

  public override prepForXml(context: Context): IXmlableObject {
    const tableObj = this.table.prepForXml(context);
    const graphicDataChildren: IXmlableObject[] = [
      { _attr: { uri: "http://schemas.openxmlformats.org/drawingml/2006/table" } },
    ];
    if (tableObj) graphicDataChildren.push(tableObj);
    return {
      "a:graphic": [{ "a:graphicData": graphicDataChildren }],
    };
  }

  public override toXml(context: Context): string {
    const tableXml = this.table.toXml(context);
    return `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tableXml}</a:graphicData></a:graphic>`;
  }
}
