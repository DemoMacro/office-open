/**
 * Page text direction module for WordprocessingML section properties.
 *
 * Defines text flow direction for pages in a section.
 *
 * Reference: http://officeopenxml.com/WPsectionPr.php
 *
 * @module
 */

/**
 * Specifies the text flow direction for pages in a section.
 *
 * This controls whether text flows horizontally (left-to-right) or
 * vertically (top-to-bottom), commonly used for East Asian languages.
 */
export const PageTextDirectionType = {
  /** Left-to-right, top-to-bottom (standard Western text flow) */
  LEFT_TO_RIGHT_TOP_TO_BOTTOM: "lrTb",
  /** Top-to-bottom, right-to-left (vertical East Asian text flow) */
  TOP_TO_BOTTOM_RIGHT_TO_LEFT: "tbRl",
} as const;
