import type { TableStyleListOptions } from "@office-open/core";
import type {
  AppPropertiesOptions,
  CorePropertiesOptions,
  CustomPropertyOptions,
} from "@office-open/core";
import type { BackgroundOptions } from "@parts/background";
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
import type { MasterPlaceholderPosition, SlideMasterOptions } from "@parts/slide-master";
import type { ControlOptions } from "@parts/slide/slide";
import type { SlideChild } from "@parts/slide/slide-child";
import type { SlideSyncOptions } from "@parts/slide/slide-sync-properties";
import type { ViewPropertiesOptions } from "@parts/view-properties";
import type { AnimationOptions } from "@shared/animation/types";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { ShapeOptions } from "@shared/shape/shape";
import type { ThemeOptions } from "@shared/theme";
import type { TransitionOptions } from "@shared/transition";

// ── Public interfaces ──

export type MasterChild = { shape: ShapeOptions };

export type SlideSize = "16:9" | "4:3" | { width: number; height: number };

export interface LayoutPlaceholderOptions {
  title?: MasterPlaceholderPosition | false;
  body?: MasterPlaceholderPosition | false;
  subtitle?: MasterPlaceholderPosition | false;
  date?: MasterPlaceholderPosition | false;
  footer?: MasterPlaceholderPosition | false;
  slideNumber?: MasterPlaceholderPosition | false;
}

export interface LayoutDefinition {
  type?: SlideLayoutType;
  name?: string;
  matchingName?: string;
  placeholders?: LayoutPlaceholderOptions;
  children?: MasterChild[];
}

export interface MasterDefinition extends SlideMasterOptions {
  name?: string;
  theme?: ThemeOptions;
  layouts?: LayoutDefinition[];
}

export interface SlideCommentOptions {
  author: string;
  text: string;
  x: number;
  y: number;
  initials?: string;
  date?: string;
}

export interface SlideAnimation {
  shapeId: number;
  options: AnimationOptions;
}

export interface SlideOptions {
  children?: SlideChild[];
  background?: BackgroundOptions;
  notes?: string;
  transition?: TransitionOptions;
  headerFooter?: SlideHeaderFooterOptions;
  comments?: SlideCommentOptions[];
  layout?: SlideLayoutType | string;
  master?: string;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  controls?: ControlOptions[];
  customerData?: { rId: string }[];
  slideSync?: SlideSyncOptions;
  animations?: SlideAnimation[];
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
  colorMru?: string[];
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
  /** Custom properties (docProps/custom.xml); omitted from the package when empty */
  customProperties?: CustomPropertyOptions[];
}
