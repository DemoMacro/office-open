import type { Context } from "../xml-components/base";
import { ImportedXmlComponent } from "../xml-components/imported";

/** Static app properties XML constant. */
export const APP_PROPS_XML =
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office Word</Application></Properties>';

export class AppProperties extends ImportedXmlComponent {
  private static instance = ImportedXmlComponent.fromXmlString(APP_PROPS_XML);

  public constructor() {
    super("Properties");
  }

  public override toXml(_context: Context): string {
    return AppProperties.instance.toXml(_context);
  }

  /** Context-free serialization — returns constant XML string. */
  public override serialize(): string {
    return APP_PROPS_XML;
  }
}
