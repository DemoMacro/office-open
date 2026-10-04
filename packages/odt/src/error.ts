/** Structured error raised when an ODT package or element cannot be parsed. */
export class OdtParseError extends Error {
  constructor(
    message: string,
    readonly part: string = "",
    readonly path: string = "",
    readonly name: string = "",
    readonly reason: string = "",
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
