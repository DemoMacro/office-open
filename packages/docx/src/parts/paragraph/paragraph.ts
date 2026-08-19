/**
 * Paragraph types for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */

import type { CustomXmlRunOptions } from "@parts/custom-xml";
import type { ObjectElementOptions } from "@parts/object";
import type { PermStartOptions } from "@parts/perm-start";
import type { PictOptions } from "@parts/pict";
import type { SdtPropertiesOptions } from "@parts/table-of-contents";
import type { ContentPartOptions } from "@shared/media/data";
import type { ChangedProperties } from "@shared/track-revision/track-revision";

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
import type { PictureOptions } from "./run/picture-run";
import type { PositionalTabOptions } from "./run/positional-tab";
import type { RubyOptions } from "./run/ruby";
import type { SimpleFieldOptions } from "./run/simple-field";
import type { SmartArtOptions } from "./run/smartart-run";
import type { SymbolRunOptions } from "./run/symbol-run";
import type { GroupOptions } from "./run/wpg-group-run";
import type { ShapeOptions } from "./run/wps-shape-run";

/** Options for an inline (run-level) structured document tag (CT_SdtRun). */
export interface SdtRunOptions {
  properties: SdtPropertiesOptions;
  children?: (ParagraphChild | string)[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
}

/** Options for a footnote/endnote reference (CT_FtnEdnRef). */
export interface FootnoteEndnoteReferenceOptions {
  /** Footnote/endnote id (w:footnoteReference/`@w:id` or w:endnoteReference/`@w:id`, required). */
  id: number;
  /** Whether a custom reference mark follows the reference (w:customMarkFollows). */
  customMarkFollows?: boolean;
}

/**
 * A complex field (PAGE/DATE/TOC/HYPERLINK... — any fldChar field without
 * w:ffData). `instruction` is the raw field code (incl. surrounding spaces);
 * `result` is the cached result-run text, if any. `rPrXml` is the verbatim
 * run-properties of the control runs (begin/instrText/separate/end);
 * `resultRPrXml` is that of the result run(s) — carried so field formatting
 * survives round-trip (Word writes the same rPr across a field's runs).
 */
export interface ComplexFieldOptions {
  instruction: string;
  result?: string;
  rPrXml?: string;
  resultRPrXml?: string;
  /** Verbatim run-properties of the end fldChar run (may differ from the
   *  control rPr — Word styles the end run like the result). */
  endRPrXml?: string;
  /** Verbatim XML of the instruction runs (begin → separate/end) when the
   *  source split them across runs with per-run properties or line breaks —
   *  shapes the plain instruction template cannot reproduce. */
  instrRunsXml?: string;
  /** Verbatim XML of the result runs (separate → end) when the source split
   *  the cached value across runs with per-run properties — e.g. Word's
   *  locale-mixed date results where every segment carries its own rFonts. */
  resultRunsXml?: string;
  /** A pagination hint Word parked on the begin run itself (w:r >
   * w:lastRenderedPageBreak + w:fldChar begin in one run). */
  lastRenderedPageBreak?: boolean;
}

/**
 * Children allowed inside a track-change wrapper (w:ins/w:del/w:moveFrom/
 * w:moveTo, CT_RunTrackChange) — runs, the comment range markers Word anchors
 * directly inside the wrapper, complete field chains, and nested same-type
 * wrappers.
 */
export type TrackChangeChild =
  | RunOptions
  | string
  | { commentRangeStart: MarkupRangeOptions }
  | { commentRangeEnd: MarkupRangeOptions }
  | { pageBreak: true }
  | { columnBreak: true }
  | { complexField: ComplexFieldOptions }
  | { formField: FormFieldOptions }
  | { proofErr: "spellStart" | "spellEnd" | "gramStart" | "gramEnd" }
  // Drawings inserted as revisions (w:ins around the drawing's w:r)
  | { picture: PictureOptions }
  | { chart: ChartOptions }
  | { wpsShape: ShapeOptions }
  | { wpgGroup: GroupOptions }
  | { insertion: ChangedProperties & { children: TrackChangeChild[] } }
  | { deletion: ChangedProperties & { children: TrackChangeChild[] } };

/** Discriminated union of all paragraph child types (inline elements, runs, etc.). */
export type ParagraphChild =
  | { chart: ChartOptions }
  | { smartArt: SmartArtOptions }
  | { picture: PictureOptions }
  | {
      math: {
        children?: MathInput[];
        /**
         * Wrap the equation in a display `m:oMathPara` container instead of an
         * inline `m:oMath` (preserved from a parsed source even without a
         * justification).
         */
        display?: boolean;
        /**
         * Display-math paragraph justification (`m:oMathPara/m:oMathParaPr/m:jc`).
         * Present → the equation is wrapped in a display `m:oMathPara`; absent →
         * inline `m:oMath`.
         */
        justification?: "left" | "right" | "center" | "centerGroup";
      };
    }
  | { symbolRun: SymbolRunOptions }
  | {
      footnoteReference: number | FootnoteEndnoteReferenceOptions;
      properties?: RunPropertiesOptions;
    }
  | {
      endnoteReference: number | FootnoteEndnoteReferenceOptions;
      properties?: RunPropertiesOptions;
    }
  | { pageBreak: true }
  | { columnBreak: true }
  | { commentRangeStart: MarkupRangeOptions }
  | { commentRangeEnd: MarkupRangeOptions }
  | { commentReference: number; properties?: RunPropertiesOptions }
  | { comment: CommentChildOptions }
  | { insertion: ChangedProperties & { children: TrackChangeChild[] } }
  | { deletion: ChangedProperties & { children: TrackChangeChild[] } }
  | {
      hyperlink: {
        url?: string;
        anchor?: string;
        tooltip?: string;
        /** Target frame for the hyperlink (CT_Hyperlink `@tgtFrame`) */
        targetFrame?: string;
        /** Location within the target document (CT_Hyperlink `@docLocation`) */
        docLocation?: string;
        /** Add the target to the navigation history (CT_Hyperlink `@history`) */
        history?: boolean;
        /** Link content: text runs plus drawings (image links) and other
         *  run-level children the paragraph dispatch serializes. */
        children?: (RunOptions | string | ParagraphChild)[];
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
  | { wpsShape: ShapeOptions }
  | { wpgGroup: GroupOptions }
  // OLE object (w:object) — occupies its own paragraph-child slot
  | { object: ObjectElementOptions }
  // VML picture (w:pict) — run-level pre-DrawingML drawing, own paragraph child
  | { pict: PictOptions }
  | { contentPart: ContentPartOptions }
  // Proof error markers
  | { proofErr: "spellStart" | "spellEnd" | "gramStart" | "gramEnd" }
  // Positional tab
  | { positionalTab: PositionalTabOptions }
  // Permission range markers
  | { permStart: PermStartOptions }
  | { permEnd: number | string }
  // Move revision range markers
  | { moveFromRangeStart: MoveRangeStartOptions }
  | { moveFromRangeEnd: MarkupRangeOptions }
  | { moveToRangeStart: MoveRangeStartOptions }
  | { moveToRangeEnd: MarkupRangeOptions }
  // Move revision text runs
  | { movedFrom: ChangedProperties & { children: TrackChangeChild[] } }
  | { movedTo: ChangedProperties & { children: TrackChangeChild[] } }
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
  // Complex field (PAGE/DATE/TOC/HYPERLINK... — see ComplexFieldOptions)
  | { complexField: ComplexFieldOptions }
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
  // Verbatim run-level XML captured on parse for shapes without a structured
  // form (unrecognized drawings, future graphicData payloads)
  | { rawXml: string }
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
  additionRsid?: string;
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
