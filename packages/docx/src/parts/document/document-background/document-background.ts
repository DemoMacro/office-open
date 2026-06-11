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

/**
 * Options for creating a document background.
 *
 * @see {@link DocumentBackground}
 */
export interface DocumentBackgroundOptions {
  /** Background color in hex format (e.g., "FF0000" for red) */
  color?: string;
  /** Theme color name (e.g., "accent1", "dark1") */
  themeColor?: string;
  /** Theme shade value (darkens the theme color) */
  themeShade?: string;
  /** Theme tint value (lightens the theme color) */
  themeTint?: string;
  /** Background image rendered as a full-page VML fill */
  image?: BackgroundImageOptions;
}
