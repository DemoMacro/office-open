import type {
  DataType,
  EncryptedContainerOptions,
  FormContainerOptions,
  TableStyleListOptions,
} from "@office-open/core";
import type { ContentTypesInput } from "@office-open/core";
import type {
  AppPropertiesOptions,
  CorePropertiesOptions,
  CustomPropertyOptions,
  ThemeOverrideOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { DateTime } from "@office-open/core";
import type { SolidFillOptions } from "@office-open/core/drawing";
import type { BackgroundOptions } from "@parts/background";
import type { ColorMappingOverrideOptions } from "@parts/descriptors/color-map-override";
import type { NotesSlideOptions } from "@parts/descriptors/notes-slide";
import type { HandoutMasterOptions } from "@parts/handout-master";
import type { NotesMasterOptions } from "@parts/notes-master";
import type {
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  SlideSizeType,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
  TagPartOptions,
  CustomXmlItemOptions,
} from "@parts/presentation";
import type {
  WebPropertiesOptions,
  PrintPropertiesOptions,
  HtmlPublishPropertiesOptions,
} from "@parts/presentation-properties";
import type { SlideLayoutType } from "@parts/slide-layout";
import type { SlideMasterOptions } from "@parts/slide-master";
import type { ControlOptions } from "@parts/slide/slide";
import type { SlideChild } from "@parts/slide/slide-child";
import type { SlideSyncOptions } from "@parts/slide/slide-sync-properties";
import type { ViewPropertiesOptions } from "@parts/view-properties";
import type { AnimationEntry, AnimationsOptions } from "@shared/animation/timing";
import type { SlideCustomerDataReferenceOptions } from "@shared/customer-data";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { PlaceholderMapOptions } from "@shared/placeholder";
import type { ThemeOptions } from "@shared/theme";
import type { TransitionOptions } from "@shared/transition";

// ── Public interfaces ──

/** Slide size — named classes or explicit dimensions in EMU or UniversalMeasure. */
export type SlideSize =
  | "16:9"
  | "4:3"
  | { width: number | UniversalMeasure; height: number | UniversalMeasure };

/** Placeholder slot map — `false` hides the slot, a definition overrides its
 * position and facets, omitted shows the default. Same value shape as
 * `MasterPlaceholderOptions`. */
export type LayoutPlaceholderOptions = PlaceholderMapOptions;

export interface LayoutDefinition {
  // Layout identity (p:sldLayout attributes)
  type?: SlideLayoutType | string;
  name?: string;
  matchingName?: string;
  preserve?: boolean;
  userDrawn?: boolean;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  /** Source p:sldLayoutId @id — kept so round-trip reuses it instead of renumbering (PowerPoint rejects renumbered ids on real-open). */
  layoutId?: number;
  /** Root spTree cNvPr id — round-trip only. */
  shapeTreeId?: number;
  /** Root spTree cNvPr name — round-trip only. */
  shapeTreeName?: string;
  /** Whether the source package carried this layout's .rels part (round-trip only). */
  sourceOwnRels?: boolean;
  // Structured cSld content (round-trip, mirrors SlideDescriptorOptions)
  children?: SlideChild[];
  background?: BackgroundOptions;
  headerFooter?: SlideHeaderFooterOptions;
  controls?: ControlOptions[];
  customerData?: SlideCustomerDataReferenceOptions[];
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip (captured from a parsed source; do not hand-author).
   */
  cSldExt?: string;
  // Child slide elements
  colorMappingOverride?: ColorMappingOverrideOptions;
  transition?: TransitionOptions;
  animations?: AnimationsOptions;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions (captured from a parsed source; do not hand-author). */
  ext?: string;
  // Fresh API (placeholder-template generation)
  placeholders?: LayoutPlaceholderOptions;
  /**
   * Verbatim layout XML; when set, the compiler parses this instead of
   * synthesizing layout XML from the structured fields. Round-trip channel:
   * captured from a parsed source document — do not hand-author.
   */
  layout?: string;
  /** Theme override (themeOverride{n}.xml part) — per-layout deviations from the owning master's theme. */
  themeOverride?: ThemeOverrideOptions;
}

/** A slide master plus its theme and layout family (p:sldMaster + p:sldLayout children). */
export interface MasterDefinition extends SlideMasterOptions {
  /** Source p:sldMasterIdLst identity (`p:sldMasterId/@id`). */
  masterId?: number;
  name?: string;
  theme?: ThemeOptions;
  layouts?: LayoutDefinition[];
}

export interface SlideCommentOptions {
  author: string;
  text: string;
  /** Anchor X in EMU or UniversalMeasure (e.g. "200px", "5cm"). */
  x: number | UniversalMeasure;
  /** Anchor Y in EMU or UniversalMeasure (e.g. "50px", "2cm"). */
  y: number | UniversalMeasure;
  initials?: string;
  date?: DateTime;
  modified?: boolean;
  /** Source author identity (`p:cm/@authorId`); round-trip only. */
  authorId?: number;
  /** Source ordinal (`p:cm/@idx`); round-trip only. */
  idx?: number;
  /** Verbatim inner XML of `p:extLst`; round-trip only. */
  ext?: string;
}

export interface CommentAuthorOptions {
  /** Stable author identifier (`p:cmAuthor/@id`). */
  id: number;
  /** Display name (`p:cmAuthor/@name`). */
  name: string;
  /** Initials (`p:cmAuthor/@initials`). */
  initials: string;
  /** Author color index (`p:cmAuthor/@clrIdx`). */
  clrIdx: number;
  /** Next comment ordinal (`p:cmAuthor/@lastIdx`). */
  lastIdx: number;
  /** Verbatim inner XML of `p:extLst`; round-trip only. */
  ext?: string;
}

// Alias of AnimationEntry — the slide-level timing entry is structurally identical.
export type SlideAnimation = AnimationEntry;
export type { AnimationsOptions } from "@shared/animation/timing";

/** One slide (p:sld) — children, notes, transition, and slide-level settings. */
export interface SlideOptions {
  /** Source p:sldIdLst identity (`p:sldId/@id`); defaults to 256 + slide index. */
  slideId?: number;
  /** `p:cSld/`@name`` (CT_CommonSlideData optional display name). */
  name?: string;
  children?: SlideChild[];
  background?: BackgroundOptions;
  /** Speaker notes — plain text shorthand, or a structured notes-slide object. */
  notes?: string | NotesSlideOptions;
  /** Form containers anchored to this slide. */
  forms?: FormContainerOptions[];
  /**
   * Slide transition. Structured form covers plain p:transition; a string is
   * the verbatim mc block a source emits for reader-version extensions
   * (mc:Choice p14:dur + mc:Fallback), re-emitted as written.
   */
  transition?: TransitionOptions | string;
  headerFooter?: SlideHeaderFooterOptions;
  /** p:clrMapOvr — override the master color mapping for this slide. */
  colorMappingOverride?: ColorMappingOverrideOptions;
  comments?: SlideCommentOptions[];
  layout?: SlideLayoutType | string;
  /** Stable layout lookup key — round-trip uses the source p:sldLayoutId @id when present. */
  layoutKey?: string;
  master?: string;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  /** Hidden slide — excluded from slideshow (emits p:sld/`@show`="0"). */
  hidden?: boolean;
  controls?: ControlOptions[];
  customerData?: SlideCustomerDataReferenceOptions[];
  slideSync?: SlideSyncOptions;
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip (captured from a parsed source; do not hand-author).
   */
  cSldExt?: string;
  /** Structured entries, or verbatim p:timing inner XML when the source tree exceeds the model. */
  animations?: SlideAnimation[] | string;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions (captured from a parsed source; do not hand-author). */
  ext?: string;
  /** Section name — slides sharing a name form one p14:section in presentation.xml. */
  section?: string;
}

/** Slide-show setup (p:presentation > p:showPr). */
export interface ShowOptions {
  loop?: boolean;
  /** Slide-show mode: "present" full screen, "browse" in a window, "kiosk" full screen, no exit. */
  type?: "present" | "browse" | "kiosk";
  showScrollbar?: boolean;
  restart?: number;
  showNarration?: boolean;
  showAnimation?: boolean;
  useTimings?: boolean;
  slideRange?: { start: number; end: number };
  /** Laser-pointer color: RGB shorthand (`"FF0000"`) or a DrawingML color choice. */
  penColor?: string | SolidFillOptions;
  /** Verbatim children of the showPr p:extLst (p14 laser pointer extensions) — round-trip channel: captured from a parsed source; do not hand-author. */
  ext?: string;
}

/** Root options for a .pptx package — maps to presentation.xml plus its satellite parts. */
export interface PresentationOptions extends CorePropertiesOptions {
  size?: SlideSize;
  /** Slide size class (`p:sldSz/@type`); cx/cy decide actual size. */
  slideSizeType?: SlideSizeType;
  /** Notes page width in EMU (`p:notesSz/@cx`, default 6858000). */
  notesWidth?: number;
  /** Notes page height in EMU (`p:notesSz/@cy`, default 9144000). */
  notesHeight?: number;
  /**
   * Source file is an encrypted OOXML package (OLE2/CFB). Round-trip only:
   * original bytes carried verbatim and re-emitted unchanged; every other
   * field stays empty — mixing slides is rejected.
   */
  encrypted?: EncryptedContainerOptions;
  masters?: MasterDefinition[];
  slides?: SlideOptions[];
  show?: ShowOptions;
  view?: ViewPropertiesOptions;
  includeHandoutMaster?: boolean;
  includeNotesMaster?: boolean;
  handoutMasterOptions?: HandoutMasterOptions;
  notesMasterOptions?: NotesMasterOptions;
  /** Canonical author list backing ppt/commentAuthors.xml; parsed sources preserve ids and extensions. */
  commentAuthors?: CommentAuthorOptions[];
  tableStyles?: TableStyleListOptions;
  web?: WebPropertiesOptions;
  print?: PrintPropertiesOptions;
  htmlPublish?: HtmlPublishPropertiesOptions;
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
  /** Package dialect to emit: ISO strict or transitional namespace set (round-trips a strict source). */
  conformance?: "strict" | "transitional";
  photoAlbum?: PhotoAlbumOptions;
  modifyVerifier?: ModifyVerifierOptions;
  embeddedFonts?: EmbeddedFontOptions[];
  customShows?: CustomShowOptions[];
  /** OPC custom-XML items (customXml/itemN.xml plus optional itemPropsN.xml). */
  customXml?: CustomXmlItemOptions[];
  /**
   * Default text style (p:defaultTextStyle) as raw inner XML. Fresh emits
   * PowerPoint's default 9-level style; a parsed source preserves its value;
   * false omits the element.
   */
  defaultTextStyle?: string | false;
  kinsoku?: KinsokuOptions[];
  customerData?: CustomerDataOptions;
  /** User-defined tags parts (ppt/tags/tagsN.xml); compiler wires the presentation-owned reference. */
  tags?: TagPartOptions[];
  /** Smart tags (p:smartTags) — r:id to the smart-tags part. */
  smartTags?: { rId: string };
  colorMru?: string[];
  /** Verbatim inner XML of p:presentationPr's p:extLst (presProps extensions) — round-trip channel: captured from a parsed source; do not hand-author. */
  presentationPropertiesExt?: string;
  /**
   * Raw `<p:ext>` entries of ppt/presentation.xml's trailing extLst outside
   * the modeled sectionLst extension (e.g. p15:sldGuideLst) — verbatim
   * round-trip, do not hand-author. Same slot as SlideOptions.ext.
   */
  ext?: string;
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /**
   * Root relationship Type for docProps/app.xml; preserved for non-canonical
   * producers. Round-trip only — do not hand-author.
   */
  appPropertiesRelationshipType?: string;
  /**
   * Content types from the source [Content_Types].xml (round-trip only).
   * Present, generate() keeps the source Default/Override entries as the
   * base declaration table and derives only what they leave uncovered.
   */
  contentTypes?: ContentTypesInput;
  /**
   * Custom properties (docProps/custom.xml). Round-trip is presence-based:
   * a source part round-trips even when it carries no properties; fresh
   * documents omit the field (and the part) entirely.
   */
  customProperties?: CustomPropertyOptions[];
  /**
   * Parts generate() does not rebuild (handout masters, customXml, unknown
   * extension parts), carried verbatim from a parsed source with bytes and
   * content-type intact — do not hand-author.
   */
  rawParts?: { path: string; data: DataType; contentType?: string }[];
  /**
   * Source .rels relationships from rebuilt parts that point at rawParts
   * (e.g. presentation.xml → handoutMaster). Re-emitted verbatim, fresh rId.
   * Round-trip only — do not hand-author.
   */
  passthroughRelationships?: {
    source: string;
    relationshipType: string;
    target: string;
    rId: string;
    targetMode?: "External";
  }[];
}
