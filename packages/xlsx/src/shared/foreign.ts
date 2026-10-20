/**
 * Foreign SpreadsheetML namespace option shapes shared by rebuilt parts.
 *
 * @module
 */

/** Foreign-namespace attribute on a rebuilt element; round-trip only. */
export interface ForeignAttributeOptions {
  /** Prefixed attribute name, e.g. "hs:extension". */
  name: string;
  /** Attribute value. */
  value: string;
}
