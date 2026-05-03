/**
 * Media transformation utilities for DrawingML.
 *
 * Converts user-facing transformation options (pixels) to internal
 * transformation data (pixels + EMUs).
 *
 * @module
 */
import { convertPixelsToEmu } from "../../converters";

/**
 * Internal media data transformation with both pixel and EMU values.
 */
export interface IMediaDataTransformation {
    readonly offset?: {
        readonly pixels: {
            readonly x: number;
            readonly y: number;
        };
        readonly emus?: {
            readonly x: number;
            readonly y: number;
        };
    };
    readonly pixels: {
        /** Width in pixels */
        readonly x: number;
        /** Height in pixels */
        readonly y: number;
    };
    /** Display dimensions in EMUs (English Metric Units) */
    readonly emus: {
        /** Width in EMUs (1 inch = 914400 EMUs) */
        readonly x: number;
        /** Height in EMUs (1 inch = 914400 EMUs) */
        readonly y: number;
    };
    /** Optional flip transformations */
    readonly flip?: {
        /** Whether to flip the image vertically */
        readonly vertical?: boolean;
        /** Whether to flip the image horizontally */
        readonly horizontal?: boolean;
    };
    /** Optional rotation angle in degrees */
    readonly rotation?: number;
}

/**
 * Transformation options for media display.
 *
 * Specifies how an image should be transformed when displayed in the document.
 */
export interface IMediaTransformation {
    readonly offset?: {
        readonly top?: number;
        readonly left?: number;
    };
    readonly width: number;
    /** Display height in pixels */
    readonly height: number;
    /** Optional flip transformations */
    readonly flip?: {
        /** Whether to flip the image vertically */
        readonly vertical?: boolean;
        /** Whether to flip the image horizontally */
        readonly horizontal?: boolean;
    };
    /** Optional rotation angle in degrees */
    readonly rotation?: number;
}

/**
 * Converts user-facing transformation options (pixels) to internal
 * transformation data (pixels + EMUs).
 *
 * @param options - User-facing transformation in pixels
 * @returns Internal transformation data with both pixel and EMU values
 */
export const createTransformation = (options: IMediaTransformation): IMediaDataTransformation => ({
    emus: {
        x: convertPixelsToEmu(options.width),
        y: convertPixelsToEmu(options.height),
    },
    flip: options.flip,
    offset: {
        emus: {
            x: convertPixelsToEmu(options.offset?.left ?? 0),
            y: convertPixelsToEmu(options.offset?.top ?? 0),
        },
        pixels: {
            x: Math.round(options.offset?.left ?? 0),
            y: Math.round(options.offset?.top ?? 0),
        },
    },
    pixels: {
        x: Math.round(options.width),
        y: Math.round(options.height),
    },
    rotation: options.rotation ? options.rotation * 60_000 : undefined,
});
