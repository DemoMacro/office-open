/**
 * Position/size conversion utilities for PPTX components.
 * @module
 */
import { convertPixelsToEmu } from "@office-open/core";

/**
 * Pixel position options accepted by most positioned components.
 */
export interface PositionOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
}

/**
 * Converts pixel position/size values to EMU, defaulting missing values to 0.
 *
 * Used by components that always need a transform (ChartFrame, Picture,
 * TableFrame, SmartArtFrame, MediaFrameBase, etc.).
 */
export function emuPosition(opts: PositionOptions): {
  x: number;
  y: number;
  width: number;
  height: number;
} {
  return {
    x: convertPixelsToEmu(opts.x ?? 0),
    y: convertPixelsToEmu(opts.y ?? 0),
    width: convertPixelsToEmu(opts.width ?? 0),
    height: convertPixelsToEmu(opts.height ?? 0),
  };
}

/**
 * Converts pixel position/size values to EMU, preserving undefined.
 *
 * Used by Shape which may omit position to let PowerPoint auto-layout.
 */
export function emuPositionOptional(opts: PositionOptions): {
  x: number | undefined;
  y: number | undefined;
  width: number | undefined;
  height: number | undefined;
} {
  return {
    x: opts.x !== undefined ? convertPixelsToEmu(opts.x) : undefined,
    y: opts.y !== undefined ? convertPixelsToEmu(opts.y) : undefined,
    width: opts.width !== undefined ? convertPixelsToEmu(opts.width) : undefined,
    height: opts.height !== undefined ? convertPixelsToEmu(opts.height) : undefined,
  };
}
