/**
 * Hyperlink types for WordprocessingML documents.
 *
 * @module
 */

/**
 * Hyperlink type enumeration.
 * @publicApi
 */
export const HyperlinkType = {
  /** Internal hyperlink to a bookmark within the document */
  INTERNAL: "INTERNAL",
  /** External hyperlink to a URL outside the document */
  EXTERNAL: "EXTERNAL",
} as const;

/**
 * Options for creating an internal hyperlink.
 */
export interface InternalHyperlinkOptions {
  /** Name of the bookmark to link to within the document */
  anchor: string;
  /** Screen tip text shown when hovering over the hyperlink */
  tooltip?: string;
}

/**
 * Options for creating an external hyperlink.
 */
export interface ExternalHyperlinkOptions {
  /** URL to link to outside the document */
  link: string;
  /** Screen tip text shown when hovering over the hyperlink */
  tooltip?: string;
  /** Target frame for the hyperlink (e.g., "_blank", "_self") */
  tgtFrame?: string;
}
