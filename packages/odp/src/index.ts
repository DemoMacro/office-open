export {
  generatePresentation,
  generatePresentationStream,
  generatePresentationSync,
  parsePresentation,
  parsePresentationSync,
} from "./presentation";
export { patchPresentation, type OdpPatchCommand, type OdpPatchOptions } from "./patch";
export { OdpGenerateError, OdpParseError } from "./error";
export type {
  OdpPresentationOptions,
  OdpSlideOptions,
  PresentationOptions,
  SlideOptions,
} from "./options";
export type { OdpEmbeddedObjectOptions, OdpEmbeddedObjectFrameOptions } from "./embedded-objects";
export type {
  OdpAppletOptions,
  OdpCubeOptions,
  OdpExtrudeOptions,
  OdpFloatingFrameOptions,
  OdpFrameContentOptions,
  OdpFrameOptions,
  OdpGluePointOptions,
  OdpImageReferenceOptions,
  OdpLightOptions,
  OdpMeasureOptions,
  OdpObjectOleOptions,
  OdpObjectReferenceOptions,
  OdpPageThumbnailOptions,
  OdpParameterOptions,
  OdpPluginOptions,
  OdpRotateOptions,
  OdpSceneObjectOptions,
  OdpSceneOptions,
  OdpSlideChild,
  OdpSlideNodeBaseOptions,
  OdpSlideNodeOptions,
  OdpSlideNodePayload,
  OdpSphereOptions,
  OdpTextBoxOptions,
  OdpVector3DOptions,
} from "./slide-nodes";
export type { PackageMemberOptions } from "@office-open/odf";
