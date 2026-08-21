/**
 * Presentation options types for PPTX.
 *
 * The Presentation XmlComponent class has been replaced by the
 * descriptor pipeline (compile/descriptors/presentation.ts).
 *
 * @module
 */

export interface PhotoAlbumOptions {
  blackWhite?: boolean;
  showCaptions?: boolean;
  /** Pictures per slide; "*Title" variants repeat the title placeholder on each slide. */
  layout?: "fitToSlide" | "1pic" | "2pic" | "4pic" | "1picTitle" | "2picTitle" | "4picTitle";
  /** Frame shape (ST_PhotoAlbumFrameShape); omit for no frame. */
  frame?:
    | "frameStyle1"
    | "frameStyle2"
    | "frameStyle3"
    | "frameStyle4"
    | "frameStyle5"
    | "frameStyle6"
    | "frameStyle7";
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
  /**
   * Slide size class (`p:sldSz/@type`, ST_SlideSizeType). Only non-default
   * values are emitted; cx/cy decide the actual size, the class is advisory.
   */
  slideSizeType?:
    | "screen4x3"
    | "letter"
    | "A4"
    | "35mm"
    | "overhead"
    | "banner"
    | "custom"
    | "ledger"
    | "A3"
    | "B4ISO"
    | "B5ISO"
    | "B4JIS"
    | "B5JIS"
    | "hagakiCard"
    | "screen16x9"
    | "screen16x10";
  /** Notes page width in EMU (`p:notesSz/@cx`, required in XML; default 6858000). */
  notesWidth?: number;
  /** Notes page height in EMU (`p:notesSz/@cy`, required in XML; default 9144000). */
  notesHeight?: number;
  slideIds: number[];
  masterCount: number;
  notesMasterRId?: number;
  handoutMasterRId?: number;
  /** Server zoom in percent (p:presentation/@serverZoom, ST_Percentage). */
  serverZoom?: number;
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
  /**
   * Default text style (p:defaultTextStyle) as raw inner XML. Fresh emits
   * PowerPoint's default 9-level style; a parsed source preserves its value;
   * false omits the element.
   */
  defaultTextStyle?: string | false;
  customerData?: CustomerDataOptions;
  /** Slide sections (p14:sectionLst); slides are grouped by name. */
  sections?: PresentationSectionGroup[];
  /**
   * Raw `<p:ext>` entries of the presentation's trailing extLst outside the
   * modeled sectionLst extension (e.g. p15:sldGuideLst slide guides) —
   * verbatim round-trip, emitted alongside the sections extension.
   */
  ext?: string;
  /** Smart tags (p:smartTags) — r:id to the smart-tags part. */
  smartTags?: { rId: string };
}

/** A named group of slides forming one p14:section in presentation.xml. */
export interface PresentationSectionGroup {
  name: string;
  /** Indices into PresentationPartOptions.slideIds belonging to this section. */
  slideIndices: number[];
}
