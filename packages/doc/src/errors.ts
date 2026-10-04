/** Structured location for a legacy DOC parse failure. */
export interface DocParseContext {
  /** Binary format that produced the error. */
  readonly format: "doc";
  /** Coarse OOXML-style source: container, stream, table, or embedded data. */
  readonly part: "container" | "stream" | "table" | "data" | "document";
  /** Canonical CFB path or logical stream name when available. */
  readonly path?: string;
  /** Numeric record type or human-readable record name when available. */
  readonly recordType?: number | string;
  /** Human-readable record name when a numeric type is not available. */
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

/** Legacy DOC input cannot be decoded safely or is not supported. */
export class DocParseError extends Error {
  /** Canonical location describing the rejected DOC construct. */
  public readonly context: DocParseContext;

  public constructor(message: string, context?: Partial<Omit<DocParseContext, "format">>) {
    super(message);
    this.name = "DocParseError";
    this.context = {
      format: "doc",
      part: "document",
      path: "document",
      reason: message,
      ...context,
    } as DocParseContext;
  }
}
