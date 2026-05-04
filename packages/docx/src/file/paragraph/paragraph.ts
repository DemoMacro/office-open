import type { FootnoteReferenceRun } from "@file/footnotes";
/**
 * Paragraph module for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";

import type { AltChunk } from "../alt-chunk";
import type { CheckBox } from "../checkbox";
import type { SectionProperties } from "../document/body/section-properties/section-properties";
import type { FileChild } from "../file-child";
import type { PermEnd, PermStart } from "../permissions";
import { TargetModeType } from "../relationships/relationship/relationship";
import type { StructuredDocumentTagRun } from "../sdt";
import type { SubDoc } from "../sub-doc";
import type {
    DeletedTextRun,
    InsertedTextRun,
    MovedFromTextRun,
    MovedToTextRun,
} from "../track-revision";
import type { ColumnBreak, PageBreak } from "./formatting/break";
import { Bookmark, ConcreteHyperlink, ExternalHyperlink } from "./links";
import type { InternalHyperlink } from "./links";
import type { Dir, Bdo } from "./links/bidi";
import type {
    MoveFromRangeEnd,
    MoveFromRangeStart,
    MoveToRangeEnd,
    MoveToRangeStart,
} from "./links/move-bookmark";
import type { Math } from "./math";
import { ParagraphProperties } from "./properties";
import type { IParagraphPropertiesOptions } from "./properties";
import { TextRun } from "./run";
import type {
    ChartRun,
    ImageRun,
    Run,
    SequentialIdentifier,
    SimpleField,
    SimpleMailMergeField,
    SmartArtRun,
    SymbolRun,
} from "./run";
import type {
    Comment,
    CommentRangeEnd,
    CommentRangeStart,
    CommentReference,
    Comments,
} from "./run/comment-run";

/**
 * The types of children that can be contained within a Paragraph element.
 * This union type represents all valid inline content elements that can appear
 * within a paragraph in WordprocessingML.
 */
export type ParagraphChild =
    | TextRun
    | ImageRun
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
    | ChartRun
    | SmartArtRun
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

/**
 * Options for creating a Paragraph element.
 *
 * @property text - Simple text content for the paragraph (creates a single TextRun)
 * @property children - Array of child elements (runs, hyperlinks, bookmarks, etc.)
 */
export type IParagraphOptions = {
    /** Simple text content for the paragraph. Creates a single TextRun. */
    readonly text?: string;
    /** Array of child elements such as TextRun, ImageRun, Hyperlink, Bookmark, etc. */
    readonly children?: readonly ParagraphChild[];
} & IParagraphPropertiesOptions;

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
    private readonly options: IParagraphOptions;
    private frontRuns: Run[] = [];
    private sectionProperties?: SectionProperties;

    public constructor(options: string | IParagraphOptions) {
        super("w:p");

        if (typeof options === "string") {
            this.options = { text: options };
        } else {
            this.options = options;
        }
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        // Build paragraph properties (including optional section properties)
        const pPr = new ParagraphProperties(this.options);
        if (this.sectionProperties) {
            pPr.push(this.sectionProperties);
        }
        const pPrObj = pPr.prepForXml(context);
        if (pPrObj) children.push(pPrObj);

        // Front runs (added via addRunToFront)
        for (const run of this.frontRuns) {
            const obj = run.prepForXml(context);
            if (obj) children.push(obj);
        }

        // Simple text shorthand
        if (this.options.text) {
            const obj = new TextRun(this.options.text).prepForXml(context);
            if (obj) children.push(obj);
        }

        // Children
        if (this.options.children) {
            for (const child of this.options.children) {
                if (child instanceof ExternalHyperlink) {
                    const concreteHyperlink = new ConcreteHyperlink(
                        child.options.children,
                        uniqueId(),
                    );
                    context.viewWrapper.Relationships.addRelationship(
                        concreteHyperlink.linkId,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                        child.options.link,
                        TargetModeType.EXTERNAL,
                    );
                    const obj = concreteHyperlink.prepForXml(context);
                    if (obj) children.push(obj);
                    continue;
                }
                if (child instanceof Bookmark) {
                    const startObj = child.start.prepForXml(context);
                    if (startObj) children.push(startObj);
                    for (const textRun of child.children) {
                        if (textRun instanceof BaseXmlComponent) {
                            const obj = textRun.prepForXml(context);
                            if (obj) children.push(obj);
                        }
                    }
                    const endObj = child.end.prepForXml(context);
                    if (endObj) children.push(endObj);
                    continue;
                }
                if (child instanceof BaseXmlComponent) {
                    const obj = child.prepForXml(context);
                    if (obj) children.push(obj);
                }
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
}
