/**
 * WordprocessingCanvas run types for WordprocessingML documents.
 *
 * @module
 */

import type { LongHexNumber } from "@office-open/core";
import type {
  BackgroundFormattingOptions,
  OfficeArtExtensionOptions,
  WholeE2oFormattingOptions,
} from "@office-open/core/drawing";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { DocPropertiesOptions } from "@parts/drawing/doc-properties/doc-properties";
import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { CanvasMediaData, GroupChildMediaData, MediaTransformation } from "@shared/media";

import type { Floating } from "../../drawing";
import type { DrawingExtensionIds, GraphicFrameLocksOptions } from "../../drawing/descriptor";

/**
 * Wordprocessing canvas (`wpc:wpc`). Children use the same typed media union
 * as `wpg:wgp`; canvas-level formatting uses shared DrawingML containers.
 *
 * @publicApi
 */
export interface WordprocessingCanvasOptions extends Omit<CanvasMediaData, "type"> {
  children: GroupChildMediaData[];
  transformation: MediaTransformation;
  background?: BackgroundFormattingOptions;
  whole?: WholeE2oFormattingOptions;
  extensions?: OfficeArtExtensionOptions[];
  floating?: Floating;
  altText?: DocPropertiesOptions;
  /**
   * Raw XML of the mc:Fallback (VML equivalent) — carried verbatim so the full
   * mc:AlternateContent round-trips. Round-trip channel: captured from a
   * parsed source document — do not hand-author.
   */
  vmlFallback?: string;
  /** Media referenced by `vmlFallback` `{fileName}` placeholders, registered on generate. */
  vmlFallbackMedia?: BackgroundRawMediaOptions[];
  /** mc:Choice Requires attribute used to regenerate the AlternateContent wrapper. */
  mcChoiceRequires?: string;
  /** Structured run properties of the wrapping w:r (round-trip). */
  runProperties?: RunPropertiesOptions;
  /** Revision save ID of the wrapping w:r (w:rsidR, round-trip). */
  additionRsid?: LongHexNumber;
  /** Revision save ID when the wrapping w:r was deleted (w:rsidDel, round-trip). */
  deletionRsid?: LongHexNumber;
  /** Revision save ID of the wrapping w:r properties (w:rsidRPr, round-trip). */
  runPropertiesRsid?: LongHexNumber;
  /** A w:lastRenderedPageBreak shared the drawing's run (round-trip). */
  lastRenderedPageBreak?: boolean;
  /** Graphic frame locks (wp:cNvGraphicFramePr) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
  /** Word 2010 wrapper ids; round-trip only. */
  extensionIds?: DrawingExtensionIds;
}
