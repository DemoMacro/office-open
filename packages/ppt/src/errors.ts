/** Structured location for a legacy PPT parse failure. */
export interface PptParseContext {
  /** Binary format that produced the error. */
  readonly format: "ppt";
  /** Coarse source: CFB container, PowerPoint stream, persisted record, or data. */
  readonly part: "container" | "stream" | "table" | "data" | "record";
  /** Canonical CFB path or logical stream name when available. */
  readonly path?: string;
  /** Numeric PowerPoint record type when available. */
  readonly recordType?: number | string;
  /** Human-readable PowerPoint record name when available. */
  readonly recordName?: string;
  /** Escher OPT property id when the failure is scoped to one property. */
  readonly propertyId?: number;
  /** Stream-relative byte offset at which validation failed. */
  readonly offset?: number;
  /** Required or declared field/record length in bytes. */
  readonly length?: number;
  /** Inclusive stream-relative byte range, when a malformed range is known. */
  readonly byteRange?: readonly [number, number];
  /** Stable machine-readable failure category. */
  readonly reason: string;
}

export class PptParseError extends Error {
  /** Canonical location describing the rejected PPT construct. */
  public readonly context: PptParseContext;

  public constructor(message: string, context?: Partial<Omit<PptParseContext, "format">>) {
    super(message);
    this.name = "PptParseError";
    this.context = {
      format: "ppt",
      part: "record",
      path: "PowerPoint Document",
      reason: message,
      ...context,
    } as PptParseContext;
  }
}
