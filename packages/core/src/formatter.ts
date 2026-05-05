import { xml } from "@office-open/xml";
import type { BaseXmlComponent, IContext, IXmlableObject } from "./xml-components";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

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

    public formatToXml(
        input: BaseXmlComponent,
        context: IContext,
        declaration?: boolean,
    ): string {
        const str = typeof input.toXml === "function"
            ? input.toXml(context)
            : xml(input.prepForXml(context)!);
        return declaration ? XML_DECL + str : str;
    }
}
