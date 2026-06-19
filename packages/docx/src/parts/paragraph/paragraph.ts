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

import type { MathInput } from "./math";
import type { ParagraphPropertiesOptions } from "./properties";
import type { RunOptions, RunPropertiesOptions } from "./run";
import type { ChartOptions } from "./run/chart-run";
import type { FormFieldOptions } from "./run/form-field";
import type { ImageOptions } from "./run/image-run";
import type { RubyOptions } from "./run/ruby";
import type { SmartArtOptions } from "./run/smartart-run";
import type { SymbolRunOptions } from "./run/symbol-run";
import type { WpgGroupRunOptions } from "./run/wpg-group-run";
import type { WpsShapeRunOptions } from "./run/wps-shape-run";

// ── JSON child wrappers ──

/** JSON-friendly wrapper for ChartRun options in paragraph children. */
export interface ChartChild {
  chart: ChartOptions;
}

/** JSON-friendly wrapper for SmartArtRun options in paragraph children. */
export interface SmartArtChild {
  smartArt: SmartArtOptions;
}

/** JSON-friendly wrapper for ImageRun options in paragraph children. */
export interface ImageChild {
  image: ImageOptions;
}

/** JSON-friendly wrapper for Math options in paragraph children. */
export interface MathChild {
  math: {
    children?: MathInput[];
  };
}

/** Options for an inline (run-level) structured document tag (CT_SdtRun). */
export interface SdtRunOptions {
  properties: SdtPropertiesOptions;
  children?: (ParagraphChild | string)[];
  /** Run properties for the SDT end mark (w:sdtEndPr). */
  endProperties?: RunPropertiesOptions;
}

/** Discriminated union of all paragraph child types (inline elements, runs, etc.). */
export type ParagraphChild =
  | ChartChild
  | SmartArtChild
  | ImageChild
  | MathChild
  | { symbolRun: SymbolRunOptions }
  | { footnoteReference: number }
  | { endnoteReference: number }
  | { pageBreak: true }
  | { columnBreak: true }
  | { commentRangeStart: number }
  | { commentRangeEnd: number }
  | { commentReference: number }
  | { insertion: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  | { deletion: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  | {
      hyperlink: {
        link?: string;
        anchor?: string;
        tooltip?: string;
        children?: (RunOptions | string)[];
      };
    }
  | { bookmarkStart: { id: number; name: string } }
  | { bookmarkEnd: number }
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
  | { moveFromRangeStart: { id: number; name?: string; author?: string; date?: string } }
  | { moveFromRangeEnd: number }
  | { moveToRangeStart: { id: number; name?: string; author?: string; date?: string } }
  | { moveToRangeEnd: number }
  // Move revision text runs
  | { movedFrom: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  | { movedTo: ChangedAttributesProperties & { children: (RunOptions | string)[] } }
  // Custom XML range markers (track changes)
  | { customXmlInsRangeStart: { id: number; author: string; date?: string } }
  | { customXmlInsRangeEnd: number }
  | { customXmlDelRangeStart: { id: number; author: string; date?: string } }
  | { customXmlDelRangeEnd: number }
  | { customXmlMoveFromRangeStart: { id: number; author: string; date?: string } }
  | { customXmlMoveFromRangeEnd: number }
  | { customXmlMoveToRangeStart: { id: number; author: string; date?: string } }
  | { customXmlMoveToRangeEnd: number }
  // Ruby annotation (East Asian pronunciation guides)
  | { ruby: RubyOptions }
  // Simple field
  | { simpleField: { instruction: string; cachedValue?: string } }
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
  | { dir: { val: "ltr" | "rtl"; children?: (RunOptions | string)[] } }
  | { bdo: { val: "ltr" | "rtl"; children?: (RunOptions | string)[] } }
  // Smart tag
  | {
      smartTag: {
        uri?: string;
        element: string;
        properties?: Array<{ uri?: string; name: string; val: string }>;
        children?: (RunOptions | string)[];
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
  /** Revision save ID for the paragraph mark. */
  rsidR?: string;
  /** Revision save ID for the paragraph properties. */
  rsidRPr?: string;
  /** Revision save ID for the default run properties. */
  rsidRDefault?: string;
  /** Revision save ID when paragraph was deleted. */
  rsidDel?: string;
  /** Revision save ID for the paragraph. */
  rsidP?: string;
} & ParagraphPropertiesOptions;
