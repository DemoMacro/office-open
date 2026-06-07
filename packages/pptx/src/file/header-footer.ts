/**
 * Slide header/footer types.
 *
 * @module
 */

export interface SlideHeaderFooterOptions {
  readonly slideNumber?: boolean;
  readonly dateTime?: boolean;
  readonly footer?: string | boolean;
  readonly header?: boolean;
}
