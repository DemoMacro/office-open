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
