/**
 * Paragraph types for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */

import type { LongHexNumber } from "@office-open/core";
import type { CustomXmlRunOptions } from "@parts/custom-xml";
import type { ObjectElementOptions } from "@parts/object";
import type { PermStartOptions } from "@parts/perm-start";
import type { PictOptions } from "@parts/pict";
import type { SubDocOptions } from "@parts/sub-doc/sub-doc";
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
import type { SimpleFieldOptions } from "./run/simple-field";
import type { SmartArtOptions } from "./run/smartart-run";
import type { SymbolRunOptions } from "./run/symbol-run";
import type { UnsupportedDrawingOptions } from "./run/unsupported-drawing-run";
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
 * A complex field (PAGE/DATE/TOC/HYPERLINK — any fldChar field without
 * w:ffData). `instruction` is the raw field code incl. surrounding spaces;
 * `result` is the cached result text; `rPrXml`/`resultRPrXml` carry the
 * control/result run properties for round-trip fidelity.
 */
export interface ComplexFieldOptions {
  instruction: string;
  /** Field lock (w:fldChar/`@w:fldLock` on the begin marker). */
  fieldLock?: boolean;
  /** Ordered members when nested simple fields occur inside a field code. */
  instructionMembers?: ComplexFieldInstructionMember[];
  /** Source `xml:space="preserve"` marker on the plain instruction text. */
  instructionPreserveSpace?: boolean;
  result?: string;
  /** Source `xml:space="preserve"` marker on the plain result text. */
  resultPreserveSpace?: boolean;
  rPrXml?: string;
  /** Revision save ID of the begin run (w:rsidR, round-trip). */
  additionRsid?: string;
  /** Revision save ID of the begin run properties (w:rsidRPr, round-trip). */
  runPropertiesRsid?: LongHexNumber;
  /** Revision save ID of the instruction run (w:rsidR, round-trip). */
  instructionAdditionRsid?: LongHexNumber;
  /** Revision save ID of the instruction run properties (w:rsidRPr, round-trip). */
  instructionRunPropertiesRsid?: LongHexNumber;
  /** Verbatim instruction-run properties; empty string means explicitly absent. */
  instructionRPrXml?: string;
  /** Revision save ID of the separator run (w:rsidR, round-trip). */
  separatorAdditionRsid?: LongHexNumber;
  /** Revision save ID of the separator run properties (w:rsidRPr, round-trip). */
  separatorRunPropertiesRsid?: LongHexNumber;
  /** Verbatim separator-run properties; empty string means explicitly absent. */
  separatorRPrXml?: string;
  resultRPrXml?: string;
  /** Revision save ID of the first plain result run (w:rsidR, round-trip). */
  resultAdditionRsid?: string;
  /** Revision save ID of the first plain result run's properties (w:rsidRPr). */
  resultRunPropertiesRsid?: string;
  /** Verbatim run-properties of the end fldChar run (may differ from the
   *  control rPr — Word styles the end run like the result). */
  endRPrXml?: string;
  /** Revision save ID of the end run (w:rsidR, round-trip). */
  endAdditionRsid?: LongHexNumber;
  /** Revision save ID of the end run properties (w:rsidRPr, round-trip). */
  endRunPropertiesRsid?: LongHexNumber;
  /** Verbatim XML of the instruction runs (begin → separate/end) when the
   *  source split them across runs with per-run properties or line breaks —
   *  shapes the plain instruction template cannot reproduce. */
  instrRunsXml?: string;
  /** Verbatim XML of the result runs (separate → end) when the source split
   *  the cached value across runs with per-run properties — e.g. Word's
   *  locale-mixed date results where every segment carries its own rFonts. */
  resultRunsXml?: string;
  /** Ordered typed members when non-run markers interrupt the cached result. */
  resultMembers?: ComplexFieldInstructionMember[];
  /** A pagination hint Word parked on the begin run itself (w:r >
   * w:lastRenderedPageBreak + w:fldChar begin in one run). */
  lastRenderedPageBreak?: boolean;
}

/**
 * One ordered member of a mixed complex-field instruction. Runs retain the
 * existing verbatim fidelity channel; nested `w:fldSimple` fields remain
 * canonical `SimpleFieldOptions`.
 */
export type ComplexFieldInstructionMember =
  | { run: RunOptions }
  | { simpleField: SimpleFieldOptions }
  | { bookmarkStart: BookmarkStartOptions }
  | { bookmarkEnd: MarkupRangeOptions };

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
  | {
      formField: FormFieldOptions;
      /** Revision save ID of the begin run (w:rsidR, round-trip). */
      additionRsid?: LongHexNumber;
      /** Revision save ID of the begin run properties (w:rsidRPr, round-trip). */
      runPropertiesRsid?: LongHexNumber;
      /** Control-run properties shared by begin/separate (round-trip). */
      rPrXml?: string;
      /** Whether the source carried a separate marker and result (round-trip). */
      hasResult?: boolean;
      /** Revision save ID of the separator run (w:rsidR, round-trip). */
      separatorAdditionRsid?: LongHexNumber;
      /** Revision save ID of the separator run properties (w:rsidRPr, round-trip). */
      separatorRunPropertiesRsid?: LongHexNumber;
      /** Verbatim separator-run properties; empty string means explicitly absent. */
      separatorRPrXml?: string;
      /** Revision save ID of the instruction run (w:rsidR, round-trip). */
      instructionAdditionRsid?: LongHexNumber;
      /** Revision save ID of the instruction run properties (w:rsidRPr, round-trip). */
      instructionRunPropertiesRsid?: LongHexNumber;
      /** Verbatim instruction-run properties (round-trip; defaults to control rPr). */
      instructionRPrXml?: string;
      /** Complete instruction-stage runs (round-trip only — do not hand-author). */
      instrRunsXml?: string;
      /** Revision save ID of the result run (w:rsidR, round-trip). */
      resultAdditionRsid?: LongHexNumber;
      /** Revision save ID of the result run properties (w:rsidRPr, round-trip). */
      resultRunPropertiesRsid?: LongHexNumber;
      /** Result-run properties (round-trip; differs from the control rPr). */
      resultRPrXml?: string;
      /** Complete result-stage runs after the separate marker (round-trip only). */
      resultRunsXml?: string;
      /** Revision save ID of the end run (w:rsidR, round-trip). */
      endAdditionRsid?: LongHexNumber;
      /** Revision save ID of the end run properties (w:rsidRPr, round-trip). */
      endRunPropertiesRsid?: LongHexNumber;
      /** Verbatim run-properties of the end fldChar run (round-trip). */
      endRPrXml?: string;
    }
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
      additionRsid?: LongHexNumber;
      runPropertiesRsid?: LongHexNumber;
      deletionRsid?: LongHexNumber;
    }
  | {
      endnoteReference: number | FootnoteEndnoteReferenceOptions;
      properties?: RunPropertiesOptions;
      additionRsid?: LongHexNumber;
      runPropertiesRsid?: LongHexNumber;
      deletionRsid?: LongHexNumber;
    }
  | { pageBreak: true }
  | { columnBreak: true }
  | { commentRangeStart: MarkupRangeOptions }
  | { commentRangeEnd: MarkupRangeOptions }
  | {
      commentReference: number;
      properties?: RunPropertiesOptions;
      additionRsid?: LongHexNumber;
      runPropertiesRsid?: LongHexNumber;
      deletionRsid?: LongHexNumber;
    }
  | { comment: CommentChildOptions }
  | { insertion: ChangedProperties & { children: TrackChangeChild[] } }
  | { deletion: ChangedProperties & { children: TrackChangeChild[] } }
  | {
      hyperlink: {
        url?: string;
        /** Source document relationship number; round-trip only. */
        sourceRelationshipId?: number;
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
  // Merge-conflict range markers (w14, Word 2010+)
  | { customXmlConflictInsRangeStart: { id: number; author?: string; date?: string } }
  | { customXmlConflictInsRangeEnd: number }
  | { customXmlConflictDelRangeStart: { id: number; author?: string; date?: string } }
  | { customXmlConflictDelRangeEnd: number }
  // Merge-conflict text runs (w14, Word 2010+ — CT_RunTrackChange wrappers)
  | { conflictIns: ChangedProperties & { children: TrackChangeChild[] } }
  | { conflictDel: ChangedProperties & { children: TrackChangeChild[] } }
  // Simple field
  | { simpleField: SimpleFieldOptions }
  // Form field (checkbox, dropdown list, text input) — control-run rsids and
  // run properties round-trip like ComplexFieldOptions.
  | {
      formField: FormFieldOptions;
      /** Revision save ID of the begin run (w:rsidR, round-trip). */
      additionRsid?: LongHexNumber;
      /** Revision save ID of the begin run properties (w:rsidRPr, round-trip). */
      runPropertiesRsid?: LongHexNumber;
      /** Control-run properties shared by begin/separate (round-trip). */
      rPrXml?: string;
      /** Whether the source carried a separate marker and result (round-trip). */
      hasResult?: boolean;
      /** Revision save ID of the separator run (w:rsidR, round-trip). */
      separatorAdditionRsid?: LongHexNumber;
      /** Revision save ID of the separator run properties (w:rsidRPr, round-trip). */
      separatorRunPropertiesRsid?: LongHexNumber;
      /** Verbatim separator-run properties; empty string means explicitly absent. */
      separatorRPrXml?: string;
      /** Revision save ID of the instruction run (w:rsidR, round-trip). */
      instructionAdditionRsid?: LongHexNumber;
      /** Revision save ID of the instruction run properties (w:rsidRPr, round-trip). */
      instructionRunPropertiesRsid?: LongHexNumber;
      /** Verbatim instruction-run properties (round-trip; defaults to control rPr). */
      instructionRPrXml?: string;
      /** Complete instruction-stage runs (round-trip only — do not hand-author). */
      instrRunsXml?: string;
      /** Revision save ID of the result run (w:rsidR, round-trip). */
      resultAdditionRsid?: LongHexNumber;
      /** Revision save ID of the result run properties (w:rsidRPr, round-trip). */
      resultRunPropertiesRsid?: LongHexNumber;
      /** Result-run properties (round-trip; differs from the control rPr). */
      resultRPrXml?: string;
      /** Complete result-stage runs after the separate marker (round-trip only). */
      resultRunsXml?: string;
      /** Revision save ID of the end run (w:rsidR, round-trip). */
      endAdditionRsid?: LongHexNumber;
      /** Revision save ID of the end run properties (w:rsidRPr, round-trip). */
      endRunPropertiesRsid?: LongHexNumber;
      /** Verbatim run-properties of the end fldChar run (round-trip). */
      endRPrXml?: string;
    }
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
  // Sub-document insertion point (w:subDoc, CT_Rel — EG_PContent member like
  // hyperlink, never a body-level child)
  | { subDoc: SubDocOptions }
  | { unsupportedDrawing: UnsupportedDrawingOptions }
  // Verbatim run-level XML for non-DrawingML constructs without a typed model.
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
  /** Preserve leading/trailing spaces in text (xml:space="preserve"). */
  preserveSpace?: boolean;
  /** Array of child elements. */
  children?: (ParagraphChild | string)[];
  /** Revision save ID for the paragraph mark (w:rsidR, CT_LongHexNumber hex string). */
  additionRsid?: LongHexNumber;
  /** Default revision save ID for runs in this paragraph (w:rsidRDefault). */
  defaultRunRsid?: LongHexNumber;
  /** Revision save ID for the paragraph properties (w:rsidP). */
  propertiesRsid?: LongHexNumber;
  /** Revision save ID for the paragraph mark run properties (w:rsidRPr). */
  runPropertiesRsid?: LongHexNumber;
  /** Revision save ID when the paragraph was deleted (w:rsidDel). */
  deletionRsid?: LongHexNumber;
  /** Unique paragraph identifier (w14:paraId, 8-digit hex string). */
  paraId?: LongHexNumber;
  /** Paragraph text identifier (w14:textId, 8-digit hex string). */
  textId?: LongHexNumber;
  /** Paragraph edit identifier (w14:editId, 8-digit hex string). */
  editId?: LongHexNumber;
  /** Whether this paragraph is excluded from spell check (w14:noSpellErr). */
  noSpellErr?: boolean;
} & ParagraphPropertiesOptions;
