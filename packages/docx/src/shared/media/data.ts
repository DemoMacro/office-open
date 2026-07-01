import type { FillOptions } from "@office-open/core/drawingml";
import type { GroupShapeLocksOptions } from "@parts/drawing/descriptor";
import type { SourceRectangleOptions } from "@parts/drawing/inline/graphic/graphic-data/pic/blip/source-rectangle";
import type { EffectListOptions } from "@parts/drawing/inline/graphic/graphic-data/pic/effects/effect-list";
import type { OutlineOptions } from "@parts/drawing/inline/graphic/graphic-data/pic/outline/outline";
import type {
  ChildExtent,
  ChildOffset,
} from "@parts/drawing/inline/graphic/graphic-data/wpg/wpg-group";
import type { WpsShapeCoreOptions } from "@parts/drawing/inline/graphic/graphic-data/wps";

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
  /**
   * Effect extent (wp:effectExtent) in raw EMUs — round-tripped verbatim from
   * the source wrapper. When absent (generation path), the descriptor computes
   * it from the shape's effects.
   */
  effectExtent?: { l: number; t: number; r: number; b: number };
}

/**
 * Round-trip of `pic:cNvPr` — the picture non-visual id/name/description.
 * All fields optional to mirror the source element: `description` is omitted when
 * absent rather than emitted as an empty attribute (Word never writes it empty).
 */
export interface NonVisualPropertiesOptions {
  id?: number;
  name?: string;
  description?: string;
  /**
   * From the sibling pic:cNvPicPr. Omitted = Word's default (true); only
   * `false` is emitted as preferRelativeResize="0" because Word never writes
   * the default value explicitly.
   */
  preferRelativeResize?: boolean;
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
  sourceRectangle?: SourceRectangleOptions;
  /** Picture non-visual properties (pic:cNvPr) for round-trip fidelity */
  nonVisualProperties?: NonVisualPropertiesOptions;
  /**
   * Blip extension `a14:useLocalDpi` (val="0" = use document DPI, not a local
   * override). Word emits this as a rendering hint on a:blip; carried verbatim
   * for round-trip fidelity. Omitted when absent (Word's default behavior).
   */
  useLocalDpi?: boolean;
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
export interface SvgMediaData {
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

export type GroupChildMediaData = (WpsMediaData | MediaData | WpgMediaData) & WpgCommonMediaData;

export interface WpgMediaData {
  type: "wpg";
  transformation: MediaDataTransformation;
  children: GroupChildMediaData[];
  /** Child coordinate offset */
  childOffset?: ChildOffset;
  /** Child coordinate extent */
  childExtent?: ChildExtent;
  /** Group fill */
  fill?: FillOptions;
  /** Group effects */
  effects?: EffectListOptions;
  /** Group shape locks (wpg:cNvGrpSpPr/a:grpSpLocks) for round-trip. */
  groupShapeLocks?: GroupShapeLocksOptions;
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

export type ExtendedMediaData =
  | MediaData
  | WpsMediaData
  | WpgMediaData
  | ChartMediaData
  | SmartArtMediaData;

export type MediaData = (RegularMediaData | SvgMediaData) & CoreMediaData;
