/**
 * Structured failure for malformed PresentationML that belongs to a modeled
 * part.
 *
 * @module
 */

export class PptxParseError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly reason: string,
  ) {
    super(message);
    this.name = "PptxParseError";
  }
}
