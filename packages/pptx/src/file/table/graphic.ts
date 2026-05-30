import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";

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

  public override toXml(context: Context): string {
    const tableXml = this.table.toXml(context);
    return `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tableXml}</a:graphicData></a:graphic>`;
  }
}
