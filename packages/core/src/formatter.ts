import type { BaseXmlComponent, IContext, IXmlableObject } from "./xml-components";

/**
 * Converts an XmlComponent tree into a serializable XML object.
 */
export class Formatter {
    public format<T extends IContext = IContext>(
        input: BaseXmlComponent,
        context: T = { stack: [] } as unknown as T,
    ): IXmlableObject {
        const output = input.prepForXml(context);
        if (output) {
            return output;
        }
        throw new Error("XMLComponent did not format correctly");
    }
}
