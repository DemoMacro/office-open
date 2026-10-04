/** Structured location for a legacy XLS parse failure. */
export interface XlsParseContext {
  /** Binary format that produced the error. */
  readonly format: "xls";
  /** Coarse source: CFB container, workbook stream, table-like range, or data. */
  readonly part: "container" | "stream" | "table" | "data" | "document";
  /** Canonical CFB path or logical stream name when available. */
  readonly path?: string;
  /** Numeric BIFF record code when available. */
  readonly recordType?: number | string;
  /** Human-readable BIFF record name when available. */
  readonly recordName?: string;
  /** Stream-relative byte offset at which validation failed. */
  readonly offset?: number;
  /** Required or declared field/record length in bytes. */
  readonly length?: number;
  /** Inclusive stream-relative byte range, when a malformed range is known. */
  readonly byteRange?: readonly [number, number];
  /** Stable machine-readable failure category. */
  readonly reason: string;
}

/** Legacy XLS input cannot be decoded safely or is not supported. */
export class XlsParseError extends Error {
  /** Canonical location describing the rejected XLS construct. */
  public readonly context: XlsParseContext;

  public constructor(message: string, context?: Partial<Omit<XlsParseContext, "format">>) {
    super(message);
    this.name = "XlsParseError";
    this.context = {
      format: "xls",
      part: "document",
      path: "Workbook",
      reason: message,
      ...context,
    } as XlsParseContext;
  }
}

export class LegacyExcelError extends XlsParseError {}
