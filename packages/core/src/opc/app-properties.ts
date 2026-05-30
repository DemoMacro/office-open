import type { Context } from "../xml-components/base";
import { ImportedXmlComponent } from "../xml-components/imported";

const APP_PROPS_XML =
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>';

export class AppProperties extends ImportedXmlComponent {
  private static instance = ImportedXmlComponent.fromXmlString(APP_PROPS_XML);

  public constructor() {
    super("Properties");
  }

  public override prepForXml() {
    return AppProperties.instance.prepForXml({ stack: [] } as any);
  }

  public override toXml(_context: Context): string {
    return AppProperties.instance.toXml(_context);
  }
}
