/** Structured error raised when a DOCX package reference cannot be resolved. */
export class DocxParseError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly reason: string,
    readonly targetPath?: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "DocxParseError";
  }
}
