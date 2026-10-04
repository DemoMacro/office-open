/** Legacy XLS input cannot be decoded safely or is not supported. */
export class XlsParseError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "XlsParseError";
  }
}

export class LegacyExcelError extends XlsParseError {}
