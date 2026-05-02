/**
 * App Properties module for PresentationML documents.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesExtended.xsd
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { AppPropertiesAttributes } from "./app-properties-attributes";

export class AppProperties extends XmlComponent {
    public constructor() {
        super("Properties");
        this.root.push(
            new AppPropertiesAttributes({
                vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            }),
        );
    }
}
