// ── Parts: standalone OOXML part types ──

export type { ViewPropertiesOptions } from "./view-properties";
export type {
  PresentationPartOptions as IPresentationXmlOptions,
  PhotoAlbumOptions,
  ModifyVerifierOptions,
  EmbeddedFontOptions,
  CustomShowOptions,
  KinsokuOptions,
  CustomerDataOptions,
} from "./presentation";
export type { CorePropertiesOptions } from "@office-open/core";
export { type SlideChild } from "./slide/slide-child";
export {
  type SlideMasterOptions,
  type MasterPlaceholderOptions,
  type MasterPlaceholderPosition,
} from "./slide-master";
export {
  DEFAULT_TEXT_STYLES,
  type TextListStyleOptions,
  type TextStylesOptions,
} from "./descriptors/text-list-style";
export { buildLayoutXml, buildCustomLayoutXml, type SlideLayoutType } from "./slide-layout";
export { DEFAULT_NOTES_STYLE, type NotesMasterOptions } from "./notes-master";
export { buildHandoutMasterXml } from "./handout-master";
export { buildNotesSlideXml } from "./notes-slide";
export { type NotesSlideOptions } from "./descriptors/notes-slide";
export type { ChartOptions } from "./chart-frame";
export type { SmartArtOptions } from "./smartart";
export type { LockedCanvasFrameOptions } from "./locked-canvas-frame";

// ── Descriptors ──
export { themeDesc } from "./descriptors/theme";
export { slideLayoutDesc } from "./descriptors/slide-layout";
export { timingDesc } from "./descriptors/animation";
export { commentAuthorsDesc, slideCommentsDesc } from "./descriptors/comments";
export { backgroundDesc } from "./descriptors/background";
export { presentationPropertiesDesc } from "./descriptors/presentation-properties";
export { slideDesc } from "./descriptors/slide";
export { slideMasterDesc } from "./descriptors/slide-master";
