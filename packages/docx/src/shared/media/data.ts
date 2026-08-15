import type { BaseMediaEntry } from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core";
import type {
  BlackWhiteMode,
  EffectListOptions,
  FillOptions,
  NonVisualContentPartPropertiesOptions,
  NonVisualDrawingPropertiesOptions,
  OutlineOptions,
  SourceRectangleOptions,
} from "@office-open/core/drawing";
import type { GraphicFrameLocksOptions, GroupShapeLocksOptions } from "@parts/drawing/descriptor";
import type {
  ChildExtent,
  ChildOffset,
} from "@parts/drawing/inline/graphic/graphic-data/wpg/wpg-group";
import type { ShapeCoreOptions } from "@parts/drawing/inline/graphic/graphic-data/wps";

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
export interface NonVisualPropertiesOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
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
interface CoreMediaData extends BaseMediaEntry {
  /** Transformation settings for display */
  transformation: MediaDataTransformation;
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

export interface ShapeMediaData {
  type: "wps";
  transformation: MediaDataTransformation;
  data: ShapeCoreOptions;
}

export interface GroupCommonMediaData {
  outline?: OutlineOptions;
  fill?: FillOptions;
}

export type GroupChildMediaData = (
  | ShapeMediaData
  | MediaData
  | GroupMediaData
  | ChartMediaData
  | ContentPartMediaData
) &
  GroupCommonMediaData;

export interface GroupMediaData {
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
 *
 * Top-level drawings always carry a chartKey assigned at the `{ chart }` sugar
 * dispatch; charts nested in a wpg group may instead carry `chartOptions`,
 * in which case the group dispatch assigns the key and registers the part.
 */
export interface ChartMediaData {
  type: "chart";
  transformation: MediaDataTransformation;
  chartKey?: string;
  /**
   * Chart payload for a group-nested chart (wpg:graphicFrame). Present on the
   * round-trip path and for fresh group children declared as options.
   * ChartSpaceOptions, not the top-level ChartOptions — the group child
   * carries the transformation.
   */
  chartOptions?: ChartSpaceOptions;
  /** Non-visual properties (wpg:cNvPr) for round-trip fidelity. */
  nonVisualProperties?: NonVisualPropertiesOptions;
  /** Graphic frame locks (wpg:cNvFrPr/a:graphicFrameLocks) for round-trip. */
  graphicFrameLocks?: GraphicFrameLocksOptions | null;
}

/**
 * Content part options (wp:contentPart / wpg:contentPart) — references an
 * opaque part via r:id. The relationship is not re-registered on generate;
 * the source r:id is passed through verbatim.
 */
export interface ContentPartOptions {
  /** Relationship id from the source document (r:id, round-trip passthrough). */
  referenceId: string;
  /** Placement transform (xfrm) */
  transformation?: MediaDataTransformation;
  /** Non-visual content part properties (cNvPr + cNvContentPartPr). */
  nonVisualProperties?: NonVisualPropertiesOptions & {
    contentPart?: NonVisualContentPartPropertiesOptions;
  };
  /** Black-and-white mode (@bwMode). */
  blackWhiteMode?: BlackWhiteMode;
}

/** Group-child / drawing-data form of {@link ContentPartOptions}. */
export interface ContentPartMediaData extends ContentPartOptions {
  type: "contentPart";
  transformation: MediaDataTransformation;
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
  | ShapeMediaData
  | GroupMediaData
  | ChartMediaData
  | SmartArtMediaData
  | ContentPartMediaData;

export type MediaData = (RegularMediaData | SvgMediaData) & CoreMediaData;
