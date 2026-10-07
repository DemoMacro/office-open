/**
 * Document background module for WordprocessingML documents.
 *
 * This module provides functionality for setting document background colors,
 * theme-based backgrounds, and background images.
 *
 * Reference: http://officeopenxml.com/WPdocument.php
 *
 * @module
 */
import type { DataType } from "@office-open/core";
import type { ThemeColor } from "@office-open/core";
import type { HexColorOrAuto, UcharHexNumber } from "@office-open/core";
import type { VmlBackgroundOptions } from "@office-open/core";

/**
 * Image options for document background.
 *
 * Specifies the image data and type for a background image.
 */
export interface BackgroundImageOptions {
  /** Raw image data (Uint8Array, base64 string, etc.) */
  data: DataType;
  /** Image format type */
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
}

/** Options for the document background (w:background). */
export interface DocumentBackgroundOptions {
  /** Background color, "auto" or hex format (e.g., "FF0000" for red) */
  color?: HexColorOrAuto;
  /** Theme color name (w:themeColor, ST_ThemeColor — e.g. "accent1", "dark1") */
  themeColor?: ThemeColor;
  /** Theme shade value (darkens the theme color) */
  themeShade?: UcharHexNumber;
  /** Theme tint value (lightens the theme color) */
  themeTint?: UcharHexNumber;
  /** Background image rendered as a full-page VML fill */
  image?: BackgroundImageOptions;
  /** Typed VML child of `w:background`; fill relationship refs use placeholders. */
  vmlBackground?: VmlBackgroundOptions;
  /** Media referenced by `rawXml` placeholders, registered on generate. */
  rawMedia?: BackgroundRawMediaOptions[];
}

/** Media item referenced by a raw-XML document background. */
export interface BackgroundRawMediaOptions {
  /** Placeholder key matching the `{fileName}` token inside the raw background XML. */
  fileName: string;
  /** Raw image data: Uint8Array, ArrayBuffer, or a base64 data URL. */
  data: DataType;
  /** Image format type. */
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
}
