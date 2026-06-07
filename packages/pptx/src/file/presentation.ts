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

export interface CustomerDataOptions {
  data?: { rId: string }[];
  tags?: { rId: string };
}

export interface PresentationOptions {
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
}

// ── Legacy types ──

export interface ViewWrapper {
  readonly view: { readonly currentOptions: unknown };
  readonly relationships: Relationships;
}
