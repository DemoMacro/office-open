/**
 * Run module for WordprocessingML documents.
 *
 * A run is a region of text with a common set of properties. It is the primary
 * unit of inline content in a paragraph.
 *
 * Reference: http://officeopenxml.com/WPtext.php
 *
 * @module
 */
import type { FootnoteReferenceRun } from "@file/footnotes/footnote/run/reference-run";
import type { FieldInstruction } from "@file/table-of-contents/field-instruction";
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { createBreak } from "./break";
import type {
    AnnotationReference,
    CarriageReturn,
    ContinuationSeparator,
    DayLong,
    DayShort,
    EndnoteReference,
    FootnoteReferenceElement,
    LastRenderedPageBreak,
    MonthLong,
    MonthShort,
    NoBreakHyphen,
    PageNumberElement,
    Separator,
    SoftHyphen,
    Tab,
    YearLong,
    YearShort,
} from "./empty-children";
import { createBegin, createEnd, createSeparate } from "./field";
import {
    buildCurrentSection,
    buildNumberOfPages,
    buildNumberOfPagesSection,
    buildPage,
} from "./page-number";
import { buildRunProperties } from "./properties";
import type { IParagraphRunPropertiesOptions, IRunPropertiesOptions } from "./properties";
import { buildText } from "./run-components/text";

interface IRunOptionsBase {
    readonly children?: readonly (
        | FieldInstruction
        | (typeof PageNumber)[keyof typeof PageNumber]
        | FootnoteReferenceRun
        | string
        | AnnotationReference
        | CarriageReturn
        | ContinuationSeparator
        | DayLong
        | DayShort
        | EndnoteReference
        | FootnoteReferenceElement
        | LastRenderedPageBreak
        | MonthLong
        | MonthShort
        | NoBreakHyphen
        | PageNumberElement
        | Separator
        | SoftHyphen
        | Tab
        | YearLong
        | YearShort
        | BaseXmlComponent
    )[];
    readonly break?: number;
    readonly text?: string;
}

/**
 * Options for creating a Run element.
 *
 * The run element specifies a region of text with a common set of properties.
 * The children property can contain various inline content elements.
 *
 * @see {@link Run}
 */
export type IRunOptions = IRunOptionsBase & IRunPropertiesOptions;

export type IParagraphRunOptions = IRunOptionsBase & IParagraphRunPropertiesOptions;

/**
 * Constants for page number field types.
 *
 * These values are used to insert dynamic page number fields into a document.
 *
 * Reference: http://officeopenxml.com/WPfields.php
 *
 * @publicApi
 */
export const PageNumber = {
    /** Inserts the current page number */
    CURRENT: "CURRENT",
    /** Inserts the total number of pages in the document */
    TOTAL_PAGES: "TOTAL_PAGES",
    /** Inserts the total number of pages in the current section */
    TOTAL_PAGES_IN_SECTION: "TOTAL_PAGES_IN_SECTION",
    /** Inserts the current section number */
    CURRENT_SECTION: "SECTION",
} as const;

/**
 * Represents a run of text with uniform formatting in a WordprocessingML document.
 *
 * A run is the lowest level unit of text in a paragraph. All content within a run
 * shares the same formatting properties (bold, italic, font, size, etc.).
 *
 * Lazy: stores options, builds IXmlableObject in prepForXml.
 */
export class Run extends BaseXmlComponent {
    private readonly options: IRunOptions;
    protected extraChildren: (BaseXmlComponent | IXmlableObject)[] = [];

    public constructor(options: IRunOptions) {
        super("w:r");
        this.options = options;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        const rPr = buildRunProperties(this.options);
        if (rPr) children.push(rPr);

        if (this.options.break) {
            for (let i = 0; i < this.options.break; i++) {
                children.push(createBreak().prepForXml(context) as IXmlableObject);
            }
        }

        if (this.options.children) {
            for (const child of this.options.children) {
                if (typeof child === "string") {
                    switch (child) {
                        case PageNumber.CURRENT: {
                            children.push(createBegin().prepForXml(context) as IXmlableObject);
                            children.push(buildPage());
                            children.push(createSeparate().prepForXml(context) as IXmlableObject);
                            children.push(createEnd().prepForXml(context) as IXmlableObject);
                            break;
                        }
                        case PageNumber.TOTAL_PAGES: {
                            children.push(createBegin().prepForXml(context) as IXmlableObject);
                            children.push(buildNumberOfPages());
                            children.push(createSeparate().prepForXml(context) as IXmlableObject);
                            children.push(createEnd().prepForXml(context) as IXmlableObject);
                            break;
                        }
                        case PageNumber.TOTAL_PAGES_IN_SECTION: {
                            children.push(createBegin().prepForXml(context) as IXmlableObject);
                            children.push(buildNumberOfPagesSection());
                            children.push(createSeparate().prepForXml(context) as IXmlableObject);
                            children.push(createEnd().prepForXml(context) as IXmlableObject);
                            break;
                        }
                        case PageNumber.CURRENT_SECTION: {
                            children.push(createBegin().prepForXml(context) as IXmlableObject);
                            children.push(buildCurrentSection());
                            children.push(createSeparate().prepForXml(context) as IXmlableObject);
                            children.push(createEnd().prepForXml(context) as IXmlableObject);
                            break;
                        }
                        default: {
                            children.push(buildText(child));
                            break;
                        }
                    }
                    continue;
                }

                if (child instanceof BaseXmlComponent) {
                    const obj = child.prepForXml(context);
                    if (obj) children.push(obj);
                }
            }
        } else if (this.options.text !== undefined) {
            children.push(buildText(this.options.text));
        }

        for (const child of this.extraChildren) {
            if (child instanceof BaseXmlComponent) {
                const obj = child.prepForXml(context);
                if (obj) children.push(obj);
            } else {
                children.push(child);
            }
        }

        return { "w:r": children.length > 0 ? children : {} };
    }
}
