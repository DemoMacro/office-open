/**
 * Presentation options types for PPTX.
 *
 * The Presentation XmlComponent class has been replaced by the
 * descriptor pipeline (compile/descriptors/presentation.ts).
 *
 * @module
 */

import type { Relationships } from "@office-open/core";

export interface PhotoAlbumOptions {
  blackWhite?: boolean;
  showCaptions?: boolean;
  layout?: "fitToSlide" | "1pic" | "2pic" | "4pic" | "1picTitle" | "2picTitle" | "4picTitle";
  frame?:
    | "frameStyle1"
    | "frameStyle2"
    | "frameStyle3"
    | "frameStyle4"
    | "frameStyle5"
    | "frameStyle6"
    | "frameStyle7"
    | "none";
}

export interface ModifyVerifierOptions {
  /** Plaintext password — automatically hashed to hashValue/saltValue when provided */
  password?: string;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinValue?: number;
  cryptoProviderType?: string;
  cryptoAlgorithmClass?: string;
  cryptoAlgorithmType?: string;
  cryptoAlgorithmSid?: number;
  spinCount?: number;
  saltData?: string;
  hashData?: string;
  cryptoProvider?: string;
  algorithmExtensionId?: number;
  algorithmExtensionSource?: string;
  cryptoProviderTypeExtension?: number;
  cryptoProviderTypeExtensionSource?: string;
}

export interface EmbeddedFontOptions {
  font: {
    typeface: string;
    panose?: string;
    pitchFamily?: number;
    charset?: number;
  };
  regular?: string;
  bold?: string;
  italic?: string;
  boldItalic?: string;
}

export interface CustomShowOptions {
  name: string;
  id: number;
  slides: { rId: string }[];
}

export interface KinsokuOptions {
  lang?: string;
  invalStChars: string;
  invalEndChars: string;
}

/** Tag entry for inline tagLst generation. */
export interface StringTagOptions {
  name: string;
  val: string;
}

export interface CustomerDataOptions {
  data?: { rId: string }[];
  tags?: { rId: string };
  /** Inline tags — generates <p:tagLst><p:tag name="..." val="..."/>...</p:tagLst> */
  tagList?: StringTagOptions[];
}

export interface PresentationPartOptions {
  slideWidth?: number;
  slideHeight?: number;
  slideIds: number[];
  masterCount: number;
  notesMasterRId?: number;
  handoutMasterRId?: number;
  serverZoom?: string;
  firstSlideNum?: number;
  showSpecialPlsOnTitleSld?: boolean;
  rtl?: boolean;
  removePersonalInfoOnSave?: boolean;
  compatMode?: boolean;
  strictFirstAndLastChars?: boolean;
  embedTrueTypeFonts?: boolean;
  saveSubsetFonts?: boolean;
  autoCompressPictures?: boolean;
  bookmarkIdSeed?: number;
  conformance?: "strict" | "transitional";
  photoAlbum?: PhotoAlbumOptions;
  modifyVerifier?: ModifyVerifierOptions;
  embeddedFonts?: EmbeddedFontOptions[];
  customShows?: CustomShowOptions[];
  kinsoku?: KinsokuOptions[];
  customerData?: CustomerDataOptions;
  /** Slide sections (p14:sectionLst); slides are grouped by name. */
  sections?: PresentationSectionGroup[];
}

/** A named group of slides forming one p14:section in presentation.xml. */
export interface PresentationSectionGroup {
  name: string;
  /** Indices into PresentationPartOptions.slideIds belonging to this section. */
  slideIndices: number[];
}

// ── Legacy types ──

export interface ViewWrapper {
  view: { currentOptions: unknown };
  relationships: Relationships;
}
