import type { PictureOptions } from "@shared/picture";

export interface ControlOptions {
  /** Shape id (AG_Ole `@spid`, ST_ShapeID token, e.g. "_x0000_s1026"). */
  shapeId?: string;
  name?: string;
  showAsIcon?: boolean;
  rId?: string;
  imageWidth?: number;
  imageHeight?: number;
  /** Fallback picture displayed by producers that do not understand the control. */
  picture?: PictureOptions;
  /** Markup-compatibility choice/fallback wrapper emitted by modern ActiveX producers. */
  alternateContent?: {
    requires?: string;
    requiresPrefix?: string;
    requiresNamespace?: string;
    choice: ControlOptions;
    fallback: ControlOptions;
  };
}
