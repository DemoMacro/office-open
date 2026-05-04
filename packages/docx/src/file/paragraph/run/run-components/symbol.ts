/**
 * Symbol module for WordprocessingML run content.
 *
 * This module provides support for inserting symbol characters
 * from symbol fonts like Wingdings or Symbol.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";

/**
 * Builds a symbol element XML object.
 *
 * @param char - Character code (hex value)
 * @param symbolfont - Symbol font name (default: "Wingdings")
 * @returns IXmlableObject for the w:sym element
 *
 * @internal
 */
export function buildSymbol(char: string = "", symbolfont: string = "Wingdings"): IXmlableObject {
    return {
        "w:sym": { _attr: { "w:char": char, "w:font": symbolfont } },
    };
}

/**
 * @deprecated Use buildSymbol() instead.
 */
export class Symbol extends BaseXmlComponent {
    private readonly _char: string;
    private readonly _font: string;

    public constructor(char: string = "", symbolfont: string = "Wingdings") {
        super("w:sym");
        this._char = char;
        this._font = symbolfont;
    }

    public prepForXml(): IXmlableObject {
        return buildSymbol(this._char, this._font);
    }
}
