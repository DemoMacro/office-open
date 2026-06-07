import type { AnimationOptions } from "@file/animation/types";
import type { BackgroundOptions } from "@file/background";
import type { CorePropertiesOptions } from "@file/core-properties";
import type { HandoutMasterOptions } from "@file/handout-master";
import type { SlideHeaderFooterOptions } from "@file/header-footer";
import type { NotesMasterOptions } from "@file/notes-master";
import type {
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
} from "@file/presentation";
import type {
  WebPropertiesOptions,
  PrintPropertiesOptions,
  HtmlPublishPropertiesOptions,
} from "@file/presentation-properties";
import type { ShapeOptions } from "@file/shape/shape";
import type { SlideLayoutType } from "@file/slide-layout";
import type { MasterPlaceholderPosition, SlideMasterOptions } from "@file/slide-master";
import type { ControlOptions } from "@file/slide/slide";
import type { SlideChild } from "@file/slide/slide-child";
import type { SlideSyncOptions } from "@file/slide/slide-sync-properties";
import type { ThemeOptions } from "@file/theme";
import type { TransitionOptions } from "@file/transition";
import type { ViewPropertiesOptions } from "@file/view-properties";
import type { BaseXmlComponent } from "@file/xml-components";
import type { TableStyleListOptions } from "@office-open/core";

// ── Public interfaces ──

export type MasterChild = BaseXmlComponent | { shape: ShapeOptions };

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
}
