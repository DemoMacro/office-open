import type { TableStyleListOptions } from "@office-open/core";
import type {
  AppPropertiesOptions,
  CorePropertiesOptions,
  CustomPropertyOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { BackgroundOptions } from "@parts/background";
import type { TimingDescriptorOptions } from "@parts/descriptors/animation";
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
import type { AnimationEntry } from "@shared/animation/timing";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { PlaceholderDefinition } from "@shared/placeholder";
import type { ShapeOptions } from "@shared/shape/shape";
import type { ThemeOptions } from "@shared/theme";
import type { TransitionOptions } from "@shared/transition";

// ── Public interfaces ──

export type MasterChild = { shape: ShapeOptions };

export type SlideSize = "16:9" | "4:3" | { width: number; height: number };

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
  // Structured cSld content (round-trip, mirrors SlideDescriptorOptions)
  children?: SlideChild[];
  background?: BackgroundOptions;
  headerFooter?: SlideHeaderFooterOptions;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  // Child slide elements
  colorMappingOverride?: ColorMappingOverrideOptions;
  transition?: TransitionOptions;
  timing?: TimingDescriptorOptions;
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
  // Fresh API (placeholder-template generation)
  placeholders?: LayoutPlaceholderOptions;
  /** Verbatim layout XML; when set, the compiler parses this instead of synthesizing layout XML from the structured fields. */
  layout?: string;
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

export interface SlideOptions {
  children?: SlideChild[];
  background?: BackgroundOptions;
  /** Speaker notes — plain text shorthand, or a structured notes-slide object. */
  notes?: string | NotesSlideOptions;
  transition?: TransitionOptions;
  headerFooter?: SlideHeaderFooterOptions;
  comments?: SlideCommentOptions[];
  layout?: SlideLayoutType | string;
  master?: string;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  /** Hidden slide — excluded from slideshow (emits p:sld/@show="0"). */
  hidden?: boolean;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  slideSync?: SlideSyncOptions;
  animations?: SlideAnimation[];
  /** Raw extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
  /** Section name — slides sharing a name form one p14:section in presentation.xml. */
  section?: string;
}

export interface ShowOptions {
  loop?: boolean;
  type?: "present" | "browse" | "kiosk";
  showScrollbar?: boolean;
  restart?: number;
  showNarration?: boolean;
  showAnimation?: boolean;
  useTimings?: boolean;
  slideRange?: { start: number; end: number };
  penColor?: string;
}

export interface PresentationOptions extends CorePropertiesOptions {
  size?: SlideSize;
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
  /** Smart tags (p:smartTags) — r:id to the smart-tags part. */
  smartTags?: { rId: string };
  colorMru?: string[];
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /** Custom properties (docProps/custom.xml); omitted from the package when empty */
  customProperties?: CustomPropertyOptions[];
}
