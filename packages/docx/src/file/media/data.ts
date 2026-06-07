import type { SourceRectangleOptions } from "@file/drawing/inline/graphic/graphic-data/pic/blip/source-rectangle";
import type { EffectListOptions } from "@file/drawing/inline/graphic/graphic-data/pic/shape-properties/effects/effect-list";
import type { OutlineOptions } from "@file/drawing/inline/graphic/graphic-data/pic/shape-properties/outline/outline";
import type {
  ChildExtent,
  ChildOffset,
} from "@file/drawing/inline/graphic/graphic-data/wpg/wpg-group";
import type { WpsShapeCoreOptions } from "@file/drawing/inline/graphic/graphic-data/wps";
import type { FillOptions } from "@office-open/core/drawingml";

export interface MediaDataTransformation {
  offset?: {
    pixels: {
      x: number;
      y: number;
    };
    emus?: {
      x: number;
      y: number;
    };
  };
  pixels: {
    /** Width in pixels */
    x: number;
    /** Height in pixels */
    y: number;
  };
  /** Display dimensions in EMUs (English Metric Units) */
  emus: {
    /** Width in EMUs (1 inch = 914400 EMUs) */
    x: number;
    /** Height in EMUs (1 inch = 914400 EMUs) */
    y: number;
  };
  /** Optional flip transformations */
  flip?: {
    /** Whether to flip the image vertically */
    vertical?: boolean;
    /** Whether to flip the image horizontally */
    horizontal?: boolean;
  };
  /** Optional rotation angle in degrees */
  rotation?: number;
}

/**
 * Core properties shared by all media data types.
 */
interface CoreMediaData {
  /** File name for the media in the package */
  fileName: string;
  /** Transformation settings for display */
  transformation: MediaDataTransformation;
  /** Raw image data */
  data: Uint8Array;
  /** Source rectangle for image cropping */
  srcRect?: SourceRectangleOptions;
}

/**
 * Regular raster image formats.
 */
interface RegularMediaData {
  /** Image format type */
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
}

/**
 * SVG image format with fallback support.
 */
interface SvgMediaData {
  /** SVG image type */
  type: "svg";
  /**
   * Fallback image for Word processors that do not support SVG.
   * This ensures the document displays correctly in all viewers.
   */
  fallback: RegularMediaData & CoreMediaData;
}

export interface WpsMediaData {
  type: "wps";
  transformation: MediaDataTransformation;
  data: WpsShapeCoreOptions;
}

export interface WpgCommonMediaData {
  outline?: OutlineOptions;
  fill?: FillOptions;
}

export type IGroupChildMediaData = (WpsMediaData | IMediaData) & WpgCommonMediaData;

export interface WpgMediaData {
  type: "wpg";
  transformation: MediaDataTransformation;
  children: IGroupChildMediaData[];
  /** Child coordinate offset */
  chOff?: ChildOffset;
  /** Child coordinate extent */
  chExt?: ChildExtent;
  /** Group fill */
  fill?: FillOptions;
  /** Group effects */
  effects?: EffectListOptions;
}

/**
 * Chart media data — references a chart part via placeholder.
 */
export interface ChartMediaData {
  type: "chart";
  transformation: MediaDataTransformation;
  chartKey: string;
}

/**
 * SmartArt media data — references a diagram data part via placeholder.
 */
export interface SmartArtMediaData {
  type: "smartart";
  transformation: MediaDataTransformation;
  smartArtKey: string;
}

export type IExtendedMediaData =
  | IMediaData
  | WpsMediaData
  | WpgMediaData
  | ChartMediaData
  | SmartArtMediaData;

export type IMediaData = (RegularMediaData | SvgMediaData) & CoreMediaData;

// Needed because of: https://github.com/s-panferov/awesome-typescript-loader/issues/432
/**
 * @ignore
 */
export const WORKAROUND2 = "";
