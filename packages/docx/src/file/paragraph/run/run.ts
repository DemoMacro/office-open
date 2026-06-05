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
import type { Context, IXmlableObject } from "@file/xml-components";
import { xml } from "@office-open/xml";

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
import type { IParagraphRunPropertiesOptions, RunPropertiesOptions } from "./properties";
import { buildText } from "./run-components/text";

interface RunOptionsBase {
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
    | IXmlableObject
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
export type RunOptions = RunOptionsBase &
  RunPropertiesOptions & {
    /** Revision save ID for run properties (hex string, e.g. "00123456"). */
    readonly rsidRPr?: string;
    /** Revision save ID when run was deleted (hex string). */
    readonly rsidDel?: string;
  };

export type IParagraphRunOptions = RunOptionsBase & IParagraphRunPropertiesOptions;

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
 * Pre-computes run properties and text at construction time to minimise
 * allocation during serialization.
 */
export class Run extends BaseXmlComponent {
  private readonly options: RunOptions;
  protected extraChildren: (BaseXmlComponent | IXmlableObject)[] = [];

  // Cached at construction time — options never change.
  private readonly _prebuilt: IXmlableObject | undefined;

  public constructor(options: RunOptions) {
    super("w:r");
    this.options = options;

    // Pre-build the entire IXmlableObject for the simple case:
    // text-only (no children array, no break, no extra children, no rsid attrs).
    if (!options.children && !options.break && !options.rsidRPr && !options.rsidDel) {
      const rPr = buildRunProperties(options);
      const text = options.text !== undefined ? buildText(options.text) : undefined;
      if (rPr || text) {
        const children: IXmlableObject[] = [];
        if (rPr) children.push(rPr);
        if (text) children.push(text);
        this._prebuilt = { "w:r": children };
      }
    }
  }

  /**
   * Hook for subclasses to register resources before serialization.
   * Called once per run during toXml().
   */
  protected registerResources(_context: Context): void {
    // Override in subclasses (ImageRun, ChartRun, etc.)
  }

  public override toXml(context: Context): string {
    // Fast path: pre-built object, no extra children
    if (this._prebuilt && this.extraChildren.length === 0) {
      return xml(this._prebuilt);
    }

    // Let subclasses register their resources (media, charts, smartArts)
    this.registerResources(context);

    // Build child XML parts
    const parts: string[] = [];

    const rPr = buildRunProperties(this.options);
    if (rPr) parts.push(xml(rPr));

    if (this.options.break) {
      for (let i = 0; i < this.options.break; i++) {
        parts.push(createBreak().toXml(context));
      }
    }

    if (this.options.children) {
      for (const child of this.options.children) {
        if (typeof child === "string") {
          switch (child) {
            case PageNumber.CURRENT:
              parts.push(createBegin().toXml(context));
              parts.push(xml(buildPage()));
              parts.push(createSeparate().toXml(context));
              parts.push(createEnd().toXml(context));
              break;
            case PageNumber.TOTAL_PAGES:
              parts.push(createBegin().toXml(context));
              parts.push(xml(buildNumberOfPages()));
              parts.push(createSeparate().toXml(context));
              parts.push(createEnd().toXml(context));
              break;
            case PageNumber.TOTAL_PAGES_IN_SECTION:
              parts.push(createBegin().toXml(context));
              parts.push(xml(buildNumberOfPagesSection()));
              parts.push(createSeparate().toXml(context));
              parts.push(createEnd().toXml(context));
              break;
            case PageNumber.CURRENT_SECTION:
              parts.push(createBegin().toXml(context));
              parts.push(xml(buildCurrentSection()));
              parts.push(createSeparate().toXml(context));
              parts.push(createEnd().toXml(context));
              break;
            default:
              parts.push(xml(buildText(child)));
              break;
          }
          continue;
        }

        if (child instanceof BaseXmlComponent) {
          const s = child.toXml(context);
          if (s) parts.push(s);
        } else {
          parts.push(xml(child));
        }
      }
    } else if (this.options.text !== undefined) {
      parts.push(xml(buildText(this.options.text)));
    }

    for (const child of this.extraChildren) {
      if (child instanceof BaseXmlComponent) {
        const s = child.toXml(context);
        if (s) parts.push(s);
      } else {
        parts.push(xml(child));
      }
    }

    const body = parts.join("");

    // Build rsid attributes on <w:r>
    const rsidAttrs: string[] = [];
    if (this.options.rsidRPr) rsidAttrs.push(` w:rsidRPr="${this.options.rsidRPr}"`);
    if (this.options.rsidDel) rsidAttrs.push(` w:rsidDel="${this.options.rsidDel}"`);
    const attr = rsidAttrs.join("");

    return body.length === 0 ? (attr ? `<w:r${attr}/>` : "<w:r/>") : `<w:r${attr}>${body}</w:r>`;
  }
}
