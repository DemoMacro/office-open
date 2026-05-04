/**
 * Page number field instruction module for WordprocessingML documents.
 *
 * This module provides field instruction elements for displaying
 * page numbers, page counts, and section information.
 *
 * Reference: http://officeopenxml.com/WPfields.php
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";

const INSTR_TEXT_ATTRS = { _attr: { "xml:space": "preserve" } };

/**
 * Builds a w:instrText element for field instructions.
 *
 * @internal
 */
function buildInstrText(instruction: string): IXmlableObject {
    return { "w:instrText": [INSTR_TEXT_ATTRS, instruction] };
}

/**
 * Builds a PAGE field instruction.
 *
 * @internal
 */
export const buildPage = () => buildInstrText("PAGE");

/**
 * Builds a NUMPAGES field instruction.
 *
 * @internal
 */
export const buildNumberOfPages = () => buildInstrText("NUMPAGES");

/**
 * Builds a SECTIONPAGES field instruction.
 *
 * @internal
 */
export const buildNumberOfPagesSection = () => buildInstrText("SECTIONPAGES");

/**
 * Builds a SECTION field instruction.
 *
 * @internal
 */
export const buildCurrentSection = () => buildInstrText("SECTION");

/**
 * @deprecated Use buildPage() instead.
 */
export class Page extends BaseXmlComponent {
    public constructor() {
        super("w:instrText");
    }
    public prepForXml(): IXmlableObject {
        return buildPage();
    }
}

/**
 * @deprecated Use buildNumberOfPages() instead.
 */
export class NumberOfPages extends BaseXmlComponent {
    public constructor() {
        super("w:instrText");
    }
    public prepForXml(): IXmlableObject {
        return buildNumberOfPages();
    }
}

/**
 * @deprecated Use buildNumberOfPagesSection() instead.
 */
export class NumberOfPagesSection extends BaseXmlComponent {
    public constructor() {
        super("w:instrText");
    }
    public prepForXml(): IXmlableObject {
        return buildNumberOfPagesSection();
    }
}

/**
 * @deprecated Use buildCurrentSection() instead.
 */
export class CurrentSection extends BaseXmlComponent {
    public constructor() {
        super("w:instrText");
    }
    public prepForXml(): IXmlableObject {
        return buildCurrentSection();
    }
}
