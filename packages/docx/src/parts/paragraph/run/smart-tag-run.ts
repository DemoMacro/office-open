/**
 * Smart tag run types for WordprocessingML documents.
 *
 * @module
 */

export interface SmartTagRunOptions {
  /** Namespace URI of the smart tag */
  uri?: string;
  /** Element name within the namespace */
  element: string;
  /** Attributes for w:smartTagPr */
  properties?: {
    uri: string;
    name: string;
    val: string;
  }[];
}
