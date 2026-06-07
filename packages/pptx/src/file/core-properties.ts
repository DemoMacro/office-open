/**
 * Core Properties types for PresentationML documents.
 *
 * @module
 */

export interface CorePropertiesOptions {
  readonly title?: string;
  readonly subject?: string;
  readonly creator?: string;
  readonly keywords?: string;
  readonly description?: string;
  readonly lastModifiedBy?: string;
  readonly revision?: number;
}
