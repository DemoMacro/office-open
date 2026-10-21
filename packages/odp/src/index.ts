export {
  generatePresentation,
  generatePresentationStream,
  generatePresentationSync,
  parsePresentation,
  parsePresentationSync,
} from "./presentation";
export { patchPresentation, type OdpPatchCommand, type OdpPatchOptions } from "./patch";
export { OdpGenerateError, OdpParseError } from "./error";
export type { OdpPresentationOptions, PresentationOptions, SlideOptions } from "./options";
export type { OdpEmbeddedObjectOptions, OdpEmbeddedObjectFrameOptions } from "./embedded-objects";
export type { OdpOpaqueSlideNodeOptions } from "./opaque-node";
export type { PackageMemberOptions } from "@office-open/odf";
