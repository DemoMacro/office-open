/** Structured error raised when an XLSX package or XML element cannot be parsed. */
export class XlsxParseError extends Error {
  constructor(
    readonly part: string,
    readonly path: string,
    readonly elementName: string,
    readonly reason: string,
  ) {
    super(`${part}${path}: ${elementName}: ${reason}`);
    this.name = "XlsxParseError";
  }
}

/** Shared structured fields for part-specific XLSX parse failures. */
export abstract class XlsxPartParseError extends Error {
  readonly part: string;
  readonly path: string;
  readonly name: string;
  readonly reason: string;

  constructor(part: string, path: string, elementName: string, reason: string, errorName: string) {
    super(`${part}${path}: ${elementName}: ${reason}`);
    this.part = part;
    this.path = path;
    this.name = errorName;
    this.reason = reason;
  }
}
