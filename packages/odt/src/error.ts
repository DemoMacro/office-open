/** Structured error raised when an ODT package or element cannot be parsed. */
export class OdtParseError extends Error {
  constructor(
    message: string,
    readonly part: string = "",
    readonly path: string = "",
    readonly elementName: string = "",
    readonly reason: string = "",
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OdtParseError";
  }
}

/** Structured error raised when canonical options cannot be represented in ODF. */
export class OdtGenerateError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly elementName: string,
    readonly reason: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OdtGenerateError";
  }
}
