import type { BaseXmlComponent, Context } from "./xml-components/base";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Converts an XmlComponent tree into a serializable XML object or string.
 */
export class Formatter {
  public format<T extends Context = Context>(
    input: BaseXmlComponent,
    context: T = { stack: [] } as unknown as T,
  ): Record<string, any> {
    const output = input.prepForXml(context);
    if (output) {
      return output;
    }
    throw new Error("XMLComponent did not format correctly");
  }

  /**
   * Serialize a component directly to XML string via the toXml() fast path.
   */
  public formatToXml(
    input: BaseXmlComponent,
    context: Context,
    declaration?: boolean | { encoding?: string; standalone?: string },
  ): string {
    const body = input.toXml(context);

    if (!body) return "";

    if (declaration) {
      const decl =
        declaration === true
          ? XML_DECL
          : `<?xml version="1.0" encoding="${declaration.encoding || "UTF-8"}"${declaration.standalone ? ` standalone="${declaration.standalone}"` : ""}?>`;
      return decl + body;
    }

    return body;
  }
}
