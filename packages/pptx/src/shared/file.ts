import type { DataType, EncryptedContainerOptions, TableStyleListOptions } from "@office-open/core";
import type { ContentTypesInput } from "@office-open/core";
import type {
  AppPropertiesOptions,
  CorePropertiesOptions,
  CustomPropertyOptions,
  ThemeOverrideOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { BackgroundOptions } from "@parts/background";
import type { ColorMappingOverrideOptions } from "@parts/descriptors/color-map-override";
import type { NotesSlideOptions } from "@parts/descriptors/notes-slide";
import type { HandoutMasterOptions } from "@parts/handout-master";
import type { NotesMasterOptions } from "@parts/notes-master";
import type {
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
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
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { PlaceholderDefinition } from "@shared/placeholder";
import type { ThemeOptions } from "@shared/theme";
import type { TransitionOptions } from "@shared/transition";

// ── Public interfaces ──

export type SlideSize = "16:9" | "4:3" | { width: number; height: number };

/** Placeholder slot map — `false` hides the slot, a definition overrides its
 * position and facets, omitted shows the default. Same value shape as
 * {@link MasterPlaceholderOptions}. */
export interface LayoutPlaceholderOptions {
  title?: PlaceholderDefinition | false;
  body?: PlaceholderDefinition | false;
  subtitle?: PlaceholderDefinition | false;
  date?: PlaceholderDefinition | false;
  footer?: PlaceholderDefinition | false;
  slideNumber?: PlaceholderDefinition | false;
}

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
  // Structured cSld content (round-trip, mirrors SlideDescriptorOptions)
  children?: SlideChild[];
  background?: BackgroundOptions;
  headerFooter?: SlideHeaderFooterOptions;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip.
   */
  cSldExt?: string;
  // Child slide elements
  colorMappingOverride?: ColorMappingOverrideOptions;
  transition?: TransitionOptions;
  animations?: AnimationsOptions;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
  // Fresh API (placeholder-template generation)
  placeholders?: LayoutPlaceholderOptions;
  /** Verbatim layout XML; when set, the compiler parses this instead of synthesizing layout XML from the structured fields. */
  layout?: string;
  /** Theme override (themeOverride{n}.xml part) — per-layout deviations from the owning master's theme. */
  themeOverride?: ThemeOverrideOptions;
}

export interface MasterDefinition extends SlideMasterOptions {
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
  date?: string;
  modified?: boolean;
}

// Alias of AnimationEntry — the slide-level timing entry is structurally identical.
export type SlideAnimation = AnimationEntry;
export type { AnimationsOptions } from "@shared/animation/timing";

export interface SlideOptions {
  children?: SlideChild[];
  background?: BackgroundOptions;
  /** Speaker notes — plain text shorthand, or a structured notes-slide object. */
  notes?: string | NotesSlideOptions;
  /**
   * Slide transition. The structured form covers the plain p:transition
   * element; a string is the verbatim markup-compatibility block a source
   * emits when the transition carries reader-version extensions (mc:Choice
   * p14:dur with an mc:Fallback twin) — re-emitted as written.
   */
  transition?: TransitionOptions | string;
  headerFooter?: SlideHeaderFooterOptions;
  /** p:clrMapOvr — override the master color mapping for this slide. */
  colorMappingOverride?: ColorMappingOverrideOptions;
  comments?: SlideCommentOptions[];
  layout?: SlideLayoutType | string;
  master?: string;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  /** Hidden slide — excluded from slideshow (emits p:sld/`@show`="0"). */
  hidden?: boolean;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  slideSync?: SlideSyncOptions;
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip.
   */
  cSldExt?: string;
  /** Structured entries, or verbatim p:timing inner XML when the source tree exceeds the model. */
  animations?: SlideAnimation[] | string;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
  /** Section name — slides sharing a name form one p14:section in presentation.xml. */
  section?: string;
}

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
  penColor?: string;
  /** Verbatim children of the showPr p:extLst (p14 laser pointer extensions). */
  ext?: string;
}

export interface PresentationOptions extends CorePropertiesOptions {
  size?: SlideSize;
  /**
   * The source file is an encrypted OOXML package (OLE2/CFB container).
   * Round-trip only: the plaintext needs the password, so the original bytes
   * are carried verbatim and generate() re-emits them unchanged — every
   * other field stays empty. Mixing slides is rejected — it would be
   * silently dropped.
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
  /**
   * Default text style (p:defaultTextStyle) as raw inner XML. Fresh emits
   * PowerPoint's default 9-level style; a parsed source preserves its value;
   * false omits the element.
   */
  defaultTextStyle?: string | false;
  kinsoku?: KinsokuOptions[];
  customerData?: CustomerDataOptions;
  /** Smart tags (p:smartTags) — r:id to the smart-tags part. */
  smartTags?: { rId: string };
  colorMru?: string[];
  /** Verbatim inner XML of p:presentationPr's p:extLst (presProps extensions). */
  presentationPropertiesExt?: string;
  /**
   * Raw `<p:ext>` entries of ppt/presentation.xml's trailing extLst outside
   * the modeled sectionLst extension (e.g. p15:sldGuideLst slide guides) —
   * verbatim round-trip. Same slot naming as SlideOptions.ext.
   */
  ext?: string;
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
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
   * Parts carried verbatim from the source that generate() does not rebuild
   * (handout masters, customXml, any unknown extension part). Collected
   * wholesale by the core passthrough pipeline — everything the model did not
   * absorb survives with bytes and content-type declaration intact. Parts the
   * compiler happens to rebuild under the same path win over the passthrough
   * copy (assembly order), so media/charts/notes need no exclusion here.
   */
  rawParts?: { path: string; data: DataType; contentType?: string }[];
  /**
   * Relationships from rebuilt parts' source .rels that point at rawParts
   * (e.g. presentation.xml → handoutMaster). Re-emitted verbatim — target
   * unchanged (passthrough paths never move), fresh rId. Round-trip only.
   */
  passthroughRelationships?: {
    source: string;
    relationshipType: string;
    target: string;
    rId: string;
    targetMode?: "External";
  }[];
}
