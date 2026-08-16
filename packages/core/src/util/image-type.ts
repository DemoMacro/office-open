/**
 * Media path extension → image type token.
 *
 * The shared normalization for every package's image-kind media: `jpeg`
 * collapses to `jpg`, `tiff` to `tif`, unknown extensions fall back to `png`.
 * Format packages narrow the superset to their own options union at the
 * consumption site.
 *
 * @module
 */

/** Image type tokens (extension-derived, superset shared by all packages). */
export type ImageFileType = "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";

/** Map a media path's file extension to an image type token. */
export function imageTypeFromPath(path: string): ImageFileType {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";
  switch (ext) {
    case "jpg":
    case "jpeg":
      return "jpg";
    case "png":
      return "png";
    case "gif":
      return "gif";
    case "bmp":
      return "bmp";
    case "tif":
    case "tiff":
      return "tif";
    case "ico":
      return "ico";
    case "emf":
      return "emf";
    case "wmf":
      return "wmf";
    default:
      return "png";
  }
}
