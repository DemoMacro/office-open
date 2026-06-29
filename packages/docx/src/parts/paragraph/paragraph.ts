/**
 * Paragraph types for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */

import type { CustomXmlRunOptions } from "@parts/custom-xml";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";
import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";

import type {
  BookmarkOptions,
  MarkupRangeOptions,
  BookmarkStartOptions,
  MoveRangeStartOptions,
  MoveRangeOptions,
} from "./links/bookmark";
import type { MathInput } from "./math";
import type { ParagraphPropertiesOptions } from "./properties";
import type { RunOptions, RunPropertiesOptions } from "./run";
import type { ChartOptions } from "./run/chart-run";
import type { CommentChildOptions } from "./run/comment-run";
import type { FormFieldOptions } from "./run/form-field";
import type { ImageOptions } from "./run/image-run";
import type { RubyOptions } from "./run/ruby";
import type { SimpleFieldOptions } from "./run/simple-field";
import type { SmartArtOptions } from "./run/smartart-run";
import type { SymbolRunOptions } from "./run/symbol-run";
import type { WpgGroupRunOptions } from "./run/wpg-group-run";
import type { WpsShapeRunOptions } from "./run/wps-shape-run";

/** Options for an inline (run-level) structured document tag (CT_SdtRun). */
export interface SdtRunOptions {
  properties: SdtPropertiesOptions;
  children?: (ParagraphChild | string)[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
}

/** Options for a footnote/endnote reference (CT_FtnEdnRef). */
export interface FootnoteEndnoteReferenceOptions {
  /** Footnote/endnote id (w:footnoteReference/@w:id or w:endnoteReference/@w:id, required). */
  id: number;
  /** Whether a custom reference mark follows the reference (w:customMarkFollows). */
  customMarkFollows?: boolean;
}

/** Discriminated union of all paragraph child types (inline elements, runs, etc.). */
export type ParagraphChild =
  | { chart: ChartOptions }
  | { smartArt: SmartArtOptions }
  | { image: ImageOptions }
  | { math: { children?: MathInput[] } }
  | { symbolRun: SymbolRunOptions }
  | { footnoteReference: number | FootnoteEndnoteReferenceOptions }
  | { endnoteReference: number | FootnoteEndnoteReferenceOptions }
  | { pageBreak: true }
  | { columnBreak: true }
  | { commentRangeStart: MarkupRangeOptions }
  | { commentRangeEnd: MarkupRangeOptions }
  | { commentReference: number }
  | { comment: CommentChildOptions }
  | { insertion: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  | { deletion: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  | {
      hyperlink: {
        link?: string;
        anchor?: string;
        tooltip?: string;
        /** Target frame for the hyperlink (CT_Hyperlink @tgtFrame) */
        tgtFrame?: string;
        /** Location within the target document (CT_Hyperlink @docLocation) */
        docLocation?: string;
        /** Add the target to the navigation history (CT_Hyperlink @history) */
        history?: boolean;
        children?: (RunOptions | string)[];
      };
      /**
       * Display-text shorthand for the hyperlink (emitted as a single text run).
       * Alternative to `hyperlink.children`; without it `{ text, hyperlink }`
       * would serialize an empty `<w:hyperlink>`.
       */
      text?: string;
    }
  | { bookmarkStart: BookmarkStartOptions }
  | { bookmarkEnd: MarkupRangeOptions }
  | { bookmark: BookmarkOptions }
  | { wpsShape: WpsShapeRunOptions }
  | { wpgGroup: WpgGroupRunOptions }
  // Proof error markers
  | { proofErr: "spellStart" | "spellEnd" | "gramStart" | "gramEnd" }
  // Positional tab
  | { positionalTab: { alignment: string; leader: string; relativeTo: string } }
  // Permission range markers
  | {
      permStart: {
        id: number | string;
        ed?: string;
        editGroup?: string;
        colFirst?: number;
        colLast?: number;
      };
    }
  | { permEnd: number | string }
  // Move revision range markers
  | { moveFromRangeStart: MoveRangeStartOptions }
  | { moveFromRangeEnd: MarkupRangeOptions }
  | { moveToRangeStart: MoveRangeStartOptions }
  | { moveToRangeEnd: MarkupRangeOptions }
  // Move revision text runs
  | { movedFrom: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  | { movedTo: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  // Move revision sugar — library allocates range + run ids and pairs markers
  | { moveFrom: MoveRangeOptions }
  | { moveTo: MoveRangeOptions }
  // Custom XML range markers (track changes)
  | { customXmlInsRangeStart: { id: number; author?: string; date?: string } }
  | { customXmlInsRangeEnd: number }
  | { customXmlDelRangeStart: { id: number; author?: string; date?: string } }
  | { customXmlDelRangeEnd: number }
  | { customXmlMoveFromRangeStart: { id: number; author?: string; date?: string } }
  | { customXmlMoveFromRangeEnd: number }
  | { customXmlMoveToRangeStart: { id: number; author?: string; date?: string } }
  | { customXmlMoveToRangeEnd: number }
  // Ruby annotation (East Asian pronunciation guides)
  | { ruby: RubyOptions }
  // Simple field
  | { simpleField: SimpleFieldOptions }
  // Form field (checkbox, dropdown list, text input)
  | { formField: FormFieldOptions }
  // Complex field (PAGE/DATE/TOC/HYPERLINK... — any fldChar field without
  // w:ffData). `instruction` is the raw field code (incl. surrounding spaces);
  // `result` is the cached result-run text, if any. `rPrXml` is the verbatim
  // run-properties of the control runs (begin/instrText/separate/end);
  // `resultRPrXml` is that of the result run(s) — carried so field formatting
  // survives round-trip (Word writes the same rPr across a field's runs).
  | {
      complexField: {
        instruction: string;
        result?: string;
        rPrXml?: string;
        resultRPrXml?: string;
      };
    }
  // Sequential identifier (SEQ field)
  | { seqIdentifier: string }
  // Page reference (PAGEREF field)
  | { pageReference: { bookmarkId: string; hyperlink?: boolean; useRelativePosition?: boolean } }
  // Bidirectional text containers
  | { dir: { val: "ltr" | "rtl"; children?: (ParagraphChild | string)[] } }
  | { bdo: { val: "ltr" | "rtl"; children?: (ParagraphChild | string)[] } }
  // Smart tag
  | {
      smartTag: {
        uri?: string;
        element: string;
        properties?: Array<{ uri?: string; name: string; val: string }>;
        children?: (ParagraphChild | string)[];
      };
    }
  // Custom XML run (CT_CustomXmlRun)
  | {
      customXml: CustomXmlRunOptions & {
        children?: (ParagraphChild | string)[];
      };
    }
  // Inline structured document tag (CT_SdtRun)
  | { sdt: SdtRunOptions }
  // Text run
  | RunOptions;

// ── ParagraphOptions ──

/**
 * Options for creating a Paragraph element.
 */
export type ParagraphOptions = {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  text?: string;
  /** Array of child elements. */
  children?: (ParagraphChild | string)[];
  /** Revision save ID for the paragraph mark (w:rsidR, CT_LongHexNumber hex string). */
  rsid?: string;
  /** Default revision save ID for runs in this paragraph (w:rsidRDefault). */
  defaultRunRsid?: string;
  /** Revision save ID for the paragraph properties (w:rsidP). */
  propertiesRsid?: string;
  /** Revision save ID for the paragraph mark run properties (w:rsidRPr). */
  runPropertiesRsid?: string;
  /** Revision save ID when the paragraph was deleted (w:rsidDel). */
  deletionRsid?: string;
  /** Unique paragraph identifier (w14:paraId, 8-digit hex string). */
  paraId?: string;
  /** Paragraph text identifier (w14:textId, 8-digit hex string). */
  textId?: string;
} & ParagraphPropertiesOptions;
