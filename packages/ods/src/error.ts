/** Structured error raised when an ODS package or element cannot be parsed. */
export class OdsParseError extends Error {
  constructor(
    message: string,
    readonly part: string = "",
    readonly path: string = "",
    readonly elementName: string = "",
    readonly reason: string = "",
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OdsParseError";
  }
}

/** Structured error raised when canonical workbook options cannot become ODS. */
export class OdsGenerateError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly elementName: string,
    readonly reason: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OdsGenerateError";
  }
}
