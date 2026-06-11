import type { HeightRule } from "@parts/table";
/**
 * Frame properties module for paragraph text frames in WordprocessingML documents.
 *
 * Frames allow paragraphs to be positioned absolutely on the page, enabling text wrapping
 * and drop cap effects. They are commonly used for floating text boxes and decorative elements.
 *
 * Reference: http://officeopenxml.com/WPparagraph-textFrames.php
 *
 * @module
 */
import type { HorizontalPositionAlign, VerticalPositionAlign } from "@shared/constants";

/**
 * Drop cap types for paragraph frames.
 */
export const DropCapType = {
  /** No drop cap effect */
  NONE: "none",
  /** Drop cap that drops down into the paragraph text */
  DROP: "drop",
  /** Drop cap that extends into the margin */
  MARGIN: "margin",
} as const;

/**
 * Frame anchor types specifying what the frame should be anchored relative to.
 */
export const FrameAnchorType = {
  /** Anchor relative to the page margin */
  MARGIN: "margin",
  /** Anchor relative to the page edge */
  PAGE: "page",
  /** Anchor relative to the text column */
  TEXT: "text",
} as const;

/**
 * Text wrapping types for frames.
 */
export const FrameWrap = {
  /** Wrap text around the frame on all sides */
  AROUND: "around",
  /** Automatic wrapping based on available space */
  AUTO: "auto",
  /** No text wrapping */
  NONE: "none",
  /** Do not allow text beside the frame */
  NOT_BESIDE: "notBeside",
  /** Allow text to flow through the frame */
  THROUGH: "through",
  /** Wrap text tightly around the frame */
  TIGHT: "tight",
} as const;

/**
 * Base options shared by all frame types.
 */
interface BaseFrameOptions {
  /** Lock the anchor position to prevent it from moving */
  anchorLock?: boolean;
  /** Drop cap effect type */
  dropCap?: (typeof DropCapType)[keyof typeof DropCapType];
  /** Frame width in twips */
  width: number;
  /** Frame height in twips */
  height: number;
  /** Text wrapping behavior around the frame */
  wrap?: (typeof FrameWrap)[keyof typeof FrameWrap];
  /** Number of lines for drop cap effect */
  lines?: number;
  /** Anchor reference points for horizontal and vertical positioning */
  anchor: {
    /** Horizontal anchor reference point */
    horizontal: (typeof FrameAnchorType)[keyof typeof FrameAnchorType];
    /** Vertical anchor reference point */
    vertical: (typeof FrameAnchorType)[keyof typeof FrameAnchorType];
  };
  /** Spacing between frame and surrounding text in twips */
  space?: {
    /** Horizontal spacing in twips */
    horizontal: number;
    /** Vertical spacing in twips */
    vertical: number;
  };
  /** Height rule determining how frame height is calculated */
  rule?: (typeof HeightRule)[keyof typeof HeightRule];
}

/**
 * Options for frames positioned using absolute X/Y coordinates.
 */
export type IXYFrameOptions = {
  /** Must be "absolute" for coordinate-based positioning */
  type: "absolute";
  /** Absolute X and Y coordinates in twips */
  position: {
    /** Horizontal position in twips from the anchor point */
    x: number;
    /** Vertical position in twips from the anchor point */
    y: number;
  };
} & BaseFrameOptions;

/**
 * Options for frames positioned using alignment values.
 */
export type IAlignmentFrameOptions = {
  /** Must be "alignment" for alignment-based positioning */
  type: "alignment";
  /** Horizontal and vertical alignment values */
  alignment: {
    /** Horizontal alignment relative to the anchor */
    x: (typeof HorizontalPositionAlign)[keyof typeof HorizontalPositionAlign];
    /** Vertical alignment relative to the anchor */
    y: (typeof VerticalPositionAlign)[keyof typeof VerticalPositionAlign];
  };
} & BaseFrameOptions;

/**
 * Union type for all frame positioning options.
 */
export type IFrameOptions = IXYFrameOptions | IAlignmentFrameOptions;
