import type { FootnoteReferenceRun } from "@file/footnotes";
/**
 * Paragraph module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { xml } from "@office-open/xml";
import { uniqueId } from "@util/convenience-functions";

import type { AltChunk } from "../alt-chunk";
import type { CheckBox } from "../checkbox";
import type { SectionProperties } from "../document/body/section-properties/section-properties";
import type { FileChild } from "../file-child";
import { FontWrapper } from "../fonts/font-wrapper";
import type { PermEnd, PermStart } from "../permissions";
import { TargetModeType } from "../relationships/relationship/relationship";
import type { StructuredDocumentTagRun } from "../sdt";
import type { SubDoc } from "../sub-doc";
import type { MovedFromTextRun, MovedToTextRun } from "../track-revision";
import type { ChangedAttributesProperties } from "../track-revision/track-revision";
import { DeletedTextRun } from "../track-revision/track-revision-components/deleted-text-run";
// Import from specific submodule to avoid circular dependency:
// paragraph.ts → track-revision barrel → inserted-text-run → ../../index (paragraph barrel)
import { InsertedTextRun } from "../track-revision/track-revision-components/inserted-text-run";
import type { ColumnBreak, PageBreak } from "./formatting/break";
import { PageBreak as PageBreakCls, ColumnBreak as ColumnBreakCls } from "./formatting/break";
import {
  Bookmark,
  BookmarkEnd,
  BookmarkStart,
  ConcreteHyperlink,
  ExternalHyperlink,
  InternalHyperlink,
} from "./links";
import type { Dir, Bdo } from "./links/bidi";
import type {
  MoveFromRangeEnd,
  MoveFromRangeStart,
  MoveToRangeEnd,
  MoveToRangeStart,
} from "./links/move-bookmark";
import { Math as MathCls } from "./math";
import type { MathOptions, MathComponent } from "./math";
import { coerceMathJson, type MathJson } from "./math/math-coerce";
type Math = InstanceType<typeof MathCls>;
// Same pattern for endnotes — specific submodule avoids barrel circular dependency.
import { EndnoteReferenceRun as EndnoteRefCls } from "@file/endnotes/endnote/run/reference-run";
// Import from specific submodule to avoid circular dependency:
// paragraph.ts → @file/footnotes barrel → footnotes.ts → imports Paragraph → paragraph.ts
// But @file/footnotes/footnote/run/reference-run only imports Run from @file/paragraph/run — no cycle.
import { FootnoteReferenceRun as FootnoteRefCls } from "@file/footnotes/footnote/run/reference-run";

import { buildParagraphProperties } from "./properties";
import type { ParagraphPropertiesOptions } from "./properties";
import { TextRun } from "./run";
import type { Run, SequentialIdentifier, SimpleField, SimpleMailMergeField } from "./run";
import type { RunOptions } from "./run";
import type {
  ChartRun as ChartRunType,
  ImageRun as ImageRunType,
  SmartArtRun as SmartArtRunType,
} from "./run";
import { ChartRun } from "./run/chart-run";
import type { ChartOptions } from "./run/chart-run";
import { Comment, CommentRangeEnd, CommentRangeStart, CommentReference } from "./run/comment-run";
import type { Comments } from "./run/comment-run";
import { ImageRun } from "./run/image-run";
import type { IImageOptions } from "./run/image-run";
import { SmartArtRun } from "./run/smartart-run";
import type { SmartArtOptions } from "./run/smartart-run";
import { SymbolRun } from "./run/symbol-run";
import type { ISymbolRunOptions } from "./run/symbol-run";

/**
 * The types of children that can be contained within a Paragraph element.
 * This union type represents all valid inline content elements that can appear
 * within a paragraph in WordprocessingML.
 */
export type ParagraphChild =
  | TextRun
  | ImageRunType
  | SymbolRun
  | Bookmark
  | PageBreak
  | ColumnBreak
  | SequentialIdentifier
  | FootnoteReferenceRun
  | InternalHyperlink
  | ExternalHyperlink
  | InsertedTextRun
  | DeletedTextRun
  | Math
  | SimpleField
  | SimpleMailMergeField
  | ChartRunType
  | SmartArtRunType
  | Comments
  | Comment
  | CommentRangeStart
  | CommentRangeEnd
  | CommentReference
  | CheckBox
  | StructuredDocumentTagRun
  | MoveFromRangeStart
  | MoveFromRangeEnd
  | MoveToRangeStart
  | MoveToRangeEnd
  | MovedFromTextRun
  | MovedToTextRun
  | PermStart
  | PermEnd
  | Dir
  | Bdo
  | AltChunk
  | SubDoc;

/** JSON-friendly wrapper for ChartRun options in paragraph children. */
export interface ChartChild {
  readonly chart: ChartOptions;
}

/** JSON-friendly wrapper for SmartArtRun options in paragraph children. */
export interface SmartArtChild {
  readonly smartArt: SmartArtOptions;
}

/** JSON-friendly wrapper for ImageRun options in paragraph children. */
export interface ImageChild {
  readonly image: IImageOptions;
}

/** JSON-friendly wrapper for Math options in paragraph children.
 *  Unlike MathOptions, children accept recursive JSON objects via {@link MathJson}. */
export interface MathChild {
  readonly math: Omit<MathOptions, "children"> & {
    readonly children?: readonly MathJson[];
  };
}

/** JSON-friendly wrappers for simple paragraph child types (JSON API). */
export type IParagraphJsonChild =
  | ChartChild
  | SmartArtChild
  | ImageChild
  | MathChild
  | { readonly symbolRun: ISymbolRunOptions }
  | { readonly footnoteReference: number }
  | { readonly endnoteReference: number }
  | { readonly pageBreak: true }
  | { readonly columnBreak: true }
  | { readonly commentRangeStart: number }
  | { readonly commentRangeEnd: number }
  | { readonly commentReference: number }
  | { readonly insertion: RunOptions & ChangedAttributesProperties }
  | { readonly deletion: RunOptions & ChangedAttributesProperties }
  | {
      readonly hyperlink: {
        readonly link?: string;
        readonly anchor?: string;
        readonly tooltip?: string;
        readonly children?: readonly (RunOptions | string)[];
      };
    }
  | { readonly bookmarkStart: { readonly id: number; readonly name: string } }
  | { readonly bookmarkEnd: number };

/**
 * Options for creating a Paragraph element.
 *
 * @property text - Simple text content for the paragraph (creates a single TextRun)
 * @property children - Array of child elements (runs, hyperlinks, bookmarks, etc.)
 */
export type ParagraphOptions = {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  readonly text?: string;
  /** Array of child elements such as TextRun, ImageRun, Hyperlink, Bookmark, etc.
   *  Accepts class instances, plain RunOptions objects (coerced to TextRun),
   *  strings (coerced to TextRun), or JSON-friendly wrappers
   *  ({ chart }, { smartArt }, { image }, { math }, { symbolRun }, etc.). */
  readonly children?: readonly (ParagraphChild | RunOptions | IParagraphJsonChild | string)[];
} & ParagraphPropertiesOptions;

/**
 * Represents a paragraph in a WordprocessingML document.
 *
 * A paragraph is the primary unit of block-level content in a document and can contain
 * various inline elements such as text runs, images, hyperlinks, and bookmarks.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @publicApi
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_P">
 *   <xsd:sequence>
 *     <xsd:element name="pPr" type="CT_PPr" minOccurs="0"/>
 *     <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rsidRPr" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidR" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidDel" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidP" type="ST_LongHexNumber"/>
 *   <xsd:attribute name="rsidRDefault" type="ST_LongHexNumber"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Simple paragraph with text
 * new Paragraph("Hello World");
 *
 * // Paragraph with options
 * new Paragraph({
 *   children: [new TextRun("Hello"), new TextRun({ text: "World", bold: true })],
 *   alignment: AlignmentType.CENTER,
 * });
 * ```
 */
export class Paragraph extends BaseXmlComponent implements FileChild {
  public readonly fileChild = Symbol();
  private readonly options: ParagraphOptions;
  private frontRuns: Run[] = [];
  private sectionProperties?: SectionProperties;

  // Cached at construction time — options never change.
  private readonly _props: ReturnType<typeof buildParagraphProperties>;
  // Pre-created TextRun for options.text shorthand (avoids allocation in prepForXml).
  private readonly _textRun: TextRun | undefined;

  public constructor(options: string | ParagraphOptions) {
    super("w:p");

    if (typeof options === "string") {
      this.options = { text: options };
    } else {
      this.options = options;
    }

    this._props = buildParagraphProperties(this.options);
    this._textRun = this.options.text ? new TextRun(this.options.text) : undefined;
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    const children: IXmlableObject[] = [];

    // Use cached paragraph properties
    const { xml: pPrObj, numberingReferences } = this._props;

    // Register numbering references (same logic as ParagraphProperties.prepForXml)
    if (!(context.viewWrapper instanceof FontWrapper)) {
      for (const reference of numberingReferences) {
        context.file.numbering.createConcreteNumberingInstance(
          reference.reference,
          reference.instance,
        );
      }
    }

    // Append section properties to pPr children if present
    let finalPPrObj = pPrObj;
    if (this.sectionProperties) {
      const sectPrObj = this.sectionProperties.prepForXml(context);
      if (sectPrObj) {
        if (finalPPrObj) {
          (finalPPrObj["w:pPr"] as IXmlableObject[]).push(sectPrObj);
        } else {
          finalPPrObj = { "w:pPr": [sectPrObj] };
        }
      }
    }

    if (finalPPrObj) children.push(finalPPrObj);

    // Front runs (added via addRunToFront)
    for (const run of this.frontRuns) {
      const obj = run.prepForXml(context);
      if (obj) children.push(obj);
    }

    // Simple text shorthand — use pre-created TextRun
    if (this._textRun) {
      const obj = this._textRun.prepForXml(context);
      if (obj) children.push(obj);
    }

    // Children
    if (this.options.children) {
      for (const rawChild of this.options.children) {
        // Bookmark and ExternalHyperlink are NOT BaseXmlComponent subclasses
        // — they need special handling before coerce.
        if (rawChild instanceof ExternalHyperlink) {
          const concreteHyperlink = new ConcreteHyperlink(rawChild.options.children, uniqueId(), {
            tooltip: rawChild.options.tooltip as string | undefined,
          });
          context.viewWrapper.relationships.addRelationship(
            concreteHyperlink.linkId,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            rawChild.options.link,
            TargetModeType.EXTERNAL,
          );
          const obj = concreteHyperlink.prepForXml(context);
          if (obj) children.push(obj);
          continue;
        }
        if (rawChild instanceof Bookmark) {
          const startObj = rawChild.start.prepForXml(context);
          if (startObj) children.push(startObj);
          for (const textRun of rawChild.children) {
            if (textRun instanceof BaseXmlComponent) {
              const obj = textRun.prepForXml(context);
              if (obj) children.push(obj);
            }
          }
          const endObj = rawChild.end.prepForXml(context);
          if (endObj) children.push(endObj);
          continue;
        }

        // Coerce strings, JSON wrappers, and plain RunOptions into Run instances
        let child: BaseXmlComponent;
        if (typeof rawChild === "string") {
          child = new TextRun(rawChild);
        } else if (rawChild instanceof BaseXmlComponent) {
          child = rawChild;
        } else if ("chart" in rawChild) {
          child = new ChartRun(rawChild.chart);
        } else if ("smartArt" in rawChild) {
          child = new SmartArtRun(rawChild.smartArt);
        } else if ("image" in rawChild) {
          child = new ImageRun(rawChild.image);
        } else if (typeof rawChild === "object" && rawChild !== null && "hyperlink" in rawChild) {
          // JSON API: { text: "...", hyperlink: { link?, anchor?, tooltip?, children? }, ...formatting }
          const { hyperlink, ...runOpts } = rawChild as Record<string, unknown> & {
            hyperlink: Record<string, unknown>;
          };
          const hlChildren = hyperlink.children as readonly (RunOptions | string)[] | undefined;
          const textRuns: TextRun[] = [];
          if (hlChildren && hlChildren.length > 0) {
            for (const rc of hlChildren) {
              textRuns.push(new TextRun(rc as RunOptions));
            }
          } else {
            textRuns.push(new TextRun(runOpts as RunOptions));
          }

          if ("link" in hyperlink) {
            const ext = new ExternalHyperlink({
              link: hyperlink.link as string,
              tooltip: hyperlink.tooltip as string | undefined,
              children: textRuns,
            });
            const concrete = new ConcreteHyperlink(ext.options.children, uniqueId(), {
              tooltip: ext.options.tooltip as string | undefined,
            });
            context.viewWrapper.relationships.addRelationship(
              concrete.linkId,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
              ext.options.link,
              TargetModeType.EXTERNAL,
            );
            const obj = concrete.prepForXml(context);
            if (obj) children.push(obj);
          } else if ("anchor" in hyperlink) {
            const internal = new InternalHyperlink({
              anchor: hyperlink.anchor as string,
              tooltip: hyperlink.tooltip as string | undefined,
              children: textRuns,
            });
            const obj = internal.prepForXml(context);
            if (obj) children.push(obj);
          }
          continue;
        } else if (
          "math" in rawChild &&
          typeof rawChild === "object" &&
          rawChild !== null &&
          typeof rawChild.math === "object" &&
          rawChild.math !== null
        ) {
          // Coerce MathJson values (strings, plain objects, class instances)
          // to MathComponent instances via recursive coerceMathJson.
          const mathOpts = rawChild.math as Omit<MathOptions, "children"> & {
            readonly children?: readonly MathJson[];
          };
          const coercedChildren = mathOpts.children?.map(coerceMathJson) as
            | readonly MathComponent[]
            | undefined;
          child = new MathCls(coercedChildren ? { children: coercedChildren } : { children: [] });
        } else if ("symbolRun" in rawChild) {
          child = new SymbolRun(rawChild.symbolRun);
        } else if ("footnoteReference" in rawChild) {
          child = new FootnoteRefCls(rawChild.footnoteReference);
        } else if ("endnoteReference" in rawChild) {
          child = new EndnoteRefCls(rawChild.endnoteReference);
        } else if ("pageBreak" in rawChild) {
          child = new PageBreakCls();
        } else if ("columnBreak" in rawChild) {
          child = new ColumnBreakCls();
        } else if ("commentRangeStart" in rawChild) {
          child = new CommentRangeStart(rawChild.commentRangeStart);
        } else if ("commentRangeEnd" in rawChild) {
          child = new CommentRangeEnd(rawChild.commentRangeEnd);
        } else if ("commentReference" in rawChild) {
          child = new TextRun({ children: [new CommentReference(rawChild.commentReference)] });
        } else if ("bookmarkStart" in rawChild) {
          const startEl = new BookmarkStart(rawChild.bookmarkStart.name, rawChild.bookmarkStart.id);
          const startObj = startEl.prepForXml(context);
          if (startObj) children.push(startObj);
          continue;
        } else if ("bookmarkEnd" in rawChild) {
          const endEl = new BookmarkEnd(rawChild.bookmarkEnd);
          const endObj = endEl.prepForXml(context);
          if (endObj) children.push(endObj);
          continue;
        } else if ("insertion" in rawChild) {
          child = new InsertedTextRun(rawChild.insertion);
        } else if ("deletion" in rawChild) {
          child = new DeletedTextRun(rawChild.deletion);
        } else {
          child = new TextRun(rawChild as RunOptions);
        }

        const obj = child.prepForXml(context);
        if (obj) children.push(obj);
      }
    }

    return { "w:p": children.length > 0 ? children : {} };
  }

  public addRunToFront(run: Run): Paragraph {
    this.frontRuns.push(run);
    return this;
  }

  /** @internal Used by Body to attach section properties for non-last sections. */
  public setSectionProperties(section: SectionProperties): void {
    this.sectionProperties = section;
  }

  /**
   * Direct XML serialization — bypasses IXmlableObject tree construction.
   * Coerces children and calls their toXml() methods for zero-allocation output.
   */
  public override toXml(context: Context): string {
    const parts: string[] = [];

    // 1. Numbering registration (side effect — same as prepForXml)
    if (!(context.viewWrapper instanceof FontWrapper)) {
      for (const reference of this._props.numberingReferences) {
        context.file.numbering.createConcreteNumberingInstance(
          reference.reference,
          reference.instance,
        );
      }
    }

    // 2. Paragraph properties (pre-built IXmlableObject → xml())
    let finalPPrObj = this._props.xml;
    if (this.sectionProperties) {
      const sectPrObj = this.sectionProperties.prepForXml(context);
      if (sectPrObj) {
        if (finalPPrObj) {
          (finalPPrObj["w:pPr"] as IXmlableObject[]).push(sectPrObj);
        } else {
          finalPPrObj = { "w:pPr": [sectPrObj] };
        }
      }
    }
    if (finalPPrObj) {
      parts.push(xml(finalPPrObj));
    }

    // 3. Front runs
    for (const run of this.frontRuns) {
      const s = run.toXml(context);
      if (s) parts.push(s);
    }

    // 4. Pre-created textRun
    if (this._textRun) {
      const s = this._textRun.toXml(context);
      if (s) parts.push(s);
    }

    // 5. Children
    if (this.options.children) {
      for (const rawChild of this.options.children) {
        for (const s of this.serializeChild(rawChild, context)) {
          if (s) parts.push(s);
        }
      }
    }

    const body = parts.join("");
    return body ? `<w:p>${body}</w:p>` : "<w:p/>";
  }

  /**
   * Serialize a single child element, handling all type coercions and side effects.
   * Returns an array of XML strings (Bookmark yields 3 parts: start + children + end).
   */
  private serializeChild(
    rawChild: ParagraphChild | RunOptions | IParagraphJsonChild | string,
    context: Context,
  ): string[] {
    // ExternalHyperlink — side effect: register relationship
    if (rawChild instanceof ExternalHyperlink) {
      const concreteHyperlink = new ConcreteHyperlink(rawChild.options.children, uniqueId(), {
        tooltip: rawChild.options.tooltip as string | undefined,
      });
      context.viewWrapper.relationships.addRelationship(
        concreteHyperlink.linkId,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        rawChild.options.link,
        TargetModeType.EXTERNAL,
      );
      return [concreteHyperlink.toXml(context)];
    }

    // Bookmark — split into start + children + end
    if (rawChild instanceof Bookmark) {
      const parts: string[] = [];
      parts.push(rawChild.start.toXml(context));
      for (const textRun of rawChild.children) {
        if (textRun instanceof BaseXmlComponent) {
          parts.push(textRun.toXml(context));
        }
      }
      parts.push(rawChild.end.toXml(context));
      return parts;
    }

    // JSON hyperlink wrapper — side effect: register relationship
    if (typeof rawChild === "object" && rawChild !== null && "hyperlink" in rawChild) {
      const { hyperlink, ...runOpts } = rawChild as Record<string, unknown> & {
        hyperlink: Record<string, unknown>;
      };
      const hlChildren = hyperlink.children as readonly (RunOptions | string)[] | undefined;
      const textRuns: TextRun[] = [];
      if (hlChildren && hlChildren.length > 0) {
        for (const rc of hlChildren) {
          textRuns.push(new TextRun(rc as RunOptions));
        }
      } else {
        textRuns.push(new TextRun(runOpts as RunOptions));
      }

      if ("link" in hyperlink) {
        const ext = new ExternalHyperlink({
          link: hyperlink.link as string,
          tooltip: hyperlink.tooltip as string | undefined,
          children: textRuns,
        });
        const concrete = new ConcreteHyperlink(ext.options.children, uniqueId(), {
          tooltip: ext.options.tooltip as string | undefined,
        });
        context.viewWrapper.relationships.addRelationship(
          concrete.linkId,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
          ext.options.link,
          TargetModeType.EXTERNAL,
        );
        return [concrete.toXml(context)];
      }
      if ("anchor" in hyperlink) {
        const internal = new InternalHyperlink({
          anchor: hyperlink.anchor as string,
          tooltip: hyperlink.tooltip as string | undefined,
          children: textRuns,
        });
        return [internal.toXml(context)];
      }
      return [];
    }

    // bookmarkStart/bookmarkEnd JSON wrappers — direct serialization
    if (typeof rawChild === "object" && rawChild !== null && "bookmarkStart" in rawChild) {
      const startEl = new BookmarkStart(rawChild.bookmarkStart.name, rawChild.bookmarkStart.id);
      return [startEl.toXml(context)];
    }
    if (typeof rawChild === "object" && rawChild !== null && "bookmarkEnd" in rawChild) {
      const endEl = new BookmarkEnd(rawChild.bookmarkEnd);
      return [endEl.toXml(context)];
    }

    // Coerce to BaseXmlComponent and serialize
    const child = this.coerceChild(rawChild);
    return child ? [child.toXml(context)] : [];
  }

  /** Coerce a raw child option into a BaseXmlComponent instance. */
  private coerceChild(
    rawChild: ParagraphChild | RunOptions | IParagraphJsonChild | string,
  ): BaseXmlComponent | undefined {
    if (typeof rawChild === "string") return new TextRun(rawChild);
    if (rawChild instanceof BaseXmlComponent) return rawChild;
    if ("chart" in rawChild) return new ChartRun(rawChild.chart);
    if ("smartArt" in rawChild) return new SmartArtRun(rawChild.smartArt);
    if ("image" in rawChild) return new ImageRun(rawChild.image);
    if (
      "math" in rawChild &&
      typeof rawChild === "object" &&
      rawChild !== null &&
      typeof rawChild.math === "object" &&
      rawChild.math !== null
    ) {
      const mathOpts = rawChild.math as Omit<MathOptions, "children"> & {
        readonly children?: readonly MathJson[];
      };
      const coercedChildren = mathOpts.children?.map(coerceMathJson) as
        | readonly MathComponent[]
        | undefined;
      return new MathCls(coercedChildren ? { children: coercedChildren } : { children: [] });
    }
    if ("symbolRun" in rawChild) return new SymbolRun(rawChild.symbolRun);
    if ("footnoteReference" in rawChild) return new FootnoteRefCls(rawChild.footnoteReference);
    if ("endnoteReference" in rawChild) return new EndnoteRefCls(rawChild.endnoteReference);
    if ("pageBreak" in rawChild) return new PageBreakCls();
    if ("columnBreak" in rawChild) return new ColumnBreakCls();
    if ("commentRangeStart" in rawChild) return new CommentRangeStart(rawChild.commentRangeStart);
    if ("commentRangeEnd" in rawChild) return new CommentRangeEnd(rawChild.commentRangeEnd);
    if ("commentReference" in rawChild)
      return new TextRun({ children: [new CommentReference(rawChild.commentReference)] });
    if ("insertion" in rawChild) return new InsertedTextRun(rawChild.insertion);
    if ("deletion" in rawChild) return new DeletedTextRun(rawChild.deletion);
    // hyperlink/bookmarkStart/bookmarkEnd are handled in serializeChild
    if ("hyperlink" in rawChild || "bookmarkStart" in rawChild || "bookmarkEnd" in rawChild) {
      return undefined;
    }
    return new TextRun(rawChild as RunOptions);
  }
}
