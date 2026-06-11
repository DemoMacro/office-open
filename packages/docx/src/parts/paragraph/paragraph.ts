/**
 * Paragraph types for WordprocessingML documents.
 *
 * Reference: http://officeopenxml.com/WPparagraph.php
 *
 * @module
 */

import type { ChangedAttributesProperties } from "@shared/track-revision/track-revision";

import type { MathInput } from "./math";
import type { ParagraphPropertiesOptions } from "./properties";
import type { RunOptions } from "./run";
import type { ChartOptions } from "./run/chart-run";
import type { IImageOptions } from "./run/image-run";
import type { RubyOptions } from "./run/ruby";
import type { ISymbolRunOptions } from "./run/symbol-run";
import type { IWpsShapeOptions } from "./run/wps-shape-run";

// ── JSON child wrappers ──

/** JSON-friendly wrapper for ChartRun options in paragraph children. */
export interface ChartChild {
  chart: ChartOptions;
}

/** JSON-friendly wrapper for SmartArtRun options in paragraph children. */
export interface SmartArtChild {
  smartArt: import("./run/smartart-run").SmartArtOptions;
}

/** JSON-friendly wrapper for ImageRun options in paragraph children. */
export interface ImageChild {
  image: IImageOptions;
}

/** JSON-friendly wrapper for Math options in paragraph children. */
export interface MathChild {
  math: {
    children?: MathInput[];
  };
}

/** JSON-friendly wrappers for simple paragraph child types (JSON API). */
export type IParagraphJsonChild =
  | ChartChild
  | SmartArtChild
  | ImageChild
  | MathChild
  | { symbolRun: ISymbolRunOptions }
  | { footnoteReference: number }
  | { endnoteReference: number }
  | { pageBreak: true }
  | { columnBreak: true }
  | { commentRangeStart: number }
  | { commentRangeEnd: number }
  | { commentReference: number }
  | { insertion: RunOptions & ChangedAttributesProperties }
  | { deletion: RunOptions & ChangedAttributesProperties }
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
  | { wpsShape: Omit<IWpsShapeOptions, "type"> }
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
  | { movedFrom: RunOptions & ChangedAttributesProperties }
  | { movedTo: RunOptions & ChangedAttributesProperties }
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
  // Custom XML run
  | {
      customXml: {
        element: string;
        uri?: string;
        customXmlPr?: {
          placeholder?: string;
          attrs?: Array<{ name: string; val: string; uri?: string }>;
        };
        children?: (RunOptions | IParagraphJsonChild | string)[];
      };
    };

// ── ParagraphOptions ──

/**
 * Options for creating a Paragraph element.
 */
export type ParagraphOptions = {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  text?: string;
  /** Array of child elements. */
  children?: (RunOptions | IParagraphJsonChild | string)[];
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
