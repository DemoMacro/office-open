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
  ViewWrapper,
} from "./presentation";
export type { CorePropertiesOptions } from "@office-open/core";
export { ContentTypes } from "./content-types";
export { type SlideChild } from "./slide/slide-child";
export {
  buildSlideMasterXml,
  type SlideMasterOptions,
  type MasterPlaceholderOptions,
  type MasterPlaceholderPosition,
} from "./slide-master";
export { buildLayoutXml, buildCustomLayoutXml, type SlideLayoutType } from "./slide-layout";
export { buildNotesMasterXml } from "./notes-master";
export { buildHandoutMasterXml } from "./handout-master";
export { buildNotesSlideXml, type NotesSlideOptions } from "./notes-slide";
export type { ChartOptions } from "./chart-frame";
export type { SmartArtOptions } from "./smartart";
export type { LockedCanvasFrameOptions } from "./locked-canvas-frame";

// ── Descriptors ──
export { themeDesc } from "./descriptors/theme";
export { slideLayoutDesc } from "./descriptors/slide-layout";
export { timingDesc } from "./descriptors/animation";
export { commentAuthorsDesc, slideCommentsDesc } from "./descriptors/comments";
export { backgroundDesc } from "./descriptors/background";
export { presPropsDesc } from "./descriptors/presentation-properties";
export { slideDesc } from "./descriptors/slide";
export { slideMasterDesc } from "./descriptors/slide-master";
