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

export type SlideSize = "16:9" | "4:3" | { readonly width: number; readonly height: number };

export interface LayoutPlaceholderOptions {
  readonly title?: MasterPlaceholderPosition | false;
  readonly body?: MasterPlaceholderPosition | false;
  readonly subtitle?: MasterPlaceholderPosition | false;
  readonly date?: MasterPlaceholderPosition | false;
  readonly footer?: MasterPlaceholderPosition | false;
  readonly slideNumber?: MasterPlaceholderPosition | false;
}

export interface LayoutDefinition {
  readonly type?: SlideLayoutType;
  readonly name?: string;
  readonly matchingName?: string;
  readonly placeholders?: LayoutPlaceholderOptions;
  readonly children?: readonly MasterChild[];
}

export interface MasterDefinition extends SlideMasterOptions {
  readonly name?: string;
  readonly theme?: ThemeOptions;
  readonly layouts?: readonly LayoutDefinition[];
}

export interface SlideCommentOptions {
  readonly author: string;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly initials?: string;
  readonly date?: string;
}

export interface SlideAnimation {
  readonly shapeId: number;
  readonly options: AnimationOptions;
}

export interface SlideOptions {
  readonly children?: readonly SlideChild[];
  readonly background?: BackgroundOptions;
  readonly notes?: string;
  readonly transition?: TransitionOptions;
  readonly headerFooter?: SlideHeaderFooterOptions;
  readonly comments?: readonly SlideCommentOptions[];
  readonly layout?: SlideLayoutType | string;
  readonly master?: string;
  readonly showMasterShapes?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly controls?: readonly ControlOptions[];
  readonly customerData?: readonly { readonly rId: string }[];
  readonly slideSync?: SlideSyncOptions;
  readonly animations?: readonly SlideAnimation[];
}

export interface ShowOptions {
  readonly loop?: boolean;
  readonly type?: "present" | "browse" | "kiosk";
  readonly showScrollbar?: boolean;
  readonly restart?: number;
  readonly showNarration?: boolean;
  readonly showAnimation?: boolean;
  readonly useTimings?: boolean;
  readonly slideRange?: { readonly start: number; readonly end: number };
  readonly penColor?: string;
}

export interface PresentationOptions extends CorePropertiesOptions {
  readonly size?: SlideSize;
  readonly masters?: readonly MasterDefinition[];
  readonly slides?: readonly SlideOptions[];
  readonly show?: ShowOptions;
  readonly view?: ViewPropertiesOptions;
  readonly includeHandoutMaster?: boolean;
  readonly includeNotesMaster?: boolean;
  readonly handoutMasterOptions?: HandoutMasterOptions;
  readonly notesMasterOptions?: NotesMasterOptions;
  readonly tableStyles?: TableStyleListOptions;
  readonly web?: WebPropertiesOptions;
  readonly print?: PrintPropertiesOptions;
  readonly htmlPublish?: HtmlPublishPropertiesOptions;
  readonly serverZoom?: string;
  readonly firstSlideNum?: number;
  readonly showSpecialPlsOnTitleSld?: boolean;
  readonly rtl?: boolean;
  readonly removePersonalInfoOnSave?: boolean;
  readonly compatMode?: boolean;
  readonly strictFirstAndLastChars?: boolean;
  readonly embedTrueTypeFonts?: boolean;
  readonly saveSubsetFonts?: boolean;
  readonly autoCompressPictures?: boolean;
  readonly bookmarkIdSeed?: number;
  readonly conformance?: "strict" | "transitional";
  readonly photoAlbum?: PhotoAlbumOptions;
  readonly modifyVerifier?: ModifyVerifierOptions;
  readonly embeddedFonts?: EmbeddedFontOptions[];
  readonly customShows?: CustomShowOptions[];
  readonly kinsoku?: KinsokuOptions[];
  readonly customerData?: CustomerDataOptions;
  readonly colorMru?: readonly string[];
}
