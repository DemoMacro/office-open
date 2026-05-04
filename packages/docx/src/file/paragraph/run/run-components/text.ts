/**
 * Text module for WordprocessingML run content.
 *
 * This module provides the buildText function for text content within runs.
 *
 * @module
 */
import { SpaceType } from "@file/shared";
import { BaseXmlComponent } from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";

interface ITextOptions {
    readonly space?: (typeof SpaceType)[keyof typeof SpaceType];
    readonly text?: string;
}

// <xsd:complexType name="CT_Text">
//     <xsd:simpleContent>
//         <xsd:extension base="s:ST_String">
//             <xsd:attribute ref="xml:space" use="optional" />
//         </xsd:extension>
//     </xsd:simpleContent>
// </xsd:complexType>

/**
 * Builds a text element XML object.
 *
 * @param options - Text string or options with space/text
 * @returns IXmlableObject for the w:t element
 *
 * @internal
 */
export function buildText(options: string | ITextOptions): IXmlableObject {
    if (typeof options === "string") {
        return {
            "w:t": [{ _attr: { "xml:space": SpaceType.PRESERVE } }, options],
        };
    }
    return {
        "w:t": [{ _attr: { "xml:space": options.space ?? SpaceType.DEFAULT } }, options.text ?? ""],
    };
}

/**
 * @deprecated Use buildText() instead.
 */
export class Text extends BaseXmlComponent {
    private readonly _opts: string | ITextOptions;

    public constructor(options: string | ITextOptions) {
        super("w:t");
        this._opts = options;
    }

    public prepForXml(): IXmlableObject {
        return buildText(this._opts);
    }
}
