export interface RtfParseContext {
  /** Logical RTF destination or container that rejected the construct. */
  part: string;
  /** Source-relative path such as `destination/shpinst`. */
  path: string;
  /** Canonical construct or RTF control name. */
  name: string;
  /** Why the decoded construct has no canonical equivalent. */
  reason: string;
}

export interface RtfGenerateContext {
  /** RTF writer block that rejected the canonical options. */
  part: string;
  /** Canonical options path, such as `sections[0].children[1].table`. */
  path: string;
  /** Canonical field or child-variant name that cannot be represented. */
  name: string;
  /** Why the field cannot be written without loss. */
  reason: string;
}

export class RtfGenerateError extends Error {
  readonly context: RtfGenerateContext;

  constructor(message: string, context: RtfGenerateContext) {
    super(message);
    this.name = "RtfGenerateError";
    this.context = context;
  }
}

export class RtfParseError extends Error {
  readonly context?: RtfParseContext;

  constructor(message: string, position: number, source: string, context?: RtfParseContext) {
    const before = source.slice(0, position);
    const line = (before.match(/\r\n|\r|\n/g)?.length ?? 0) + 1;
    const lastLine = Math.max(before.lastIndexOf("\n"), before.lastIndexOf("\r"));
    const column = position - lastLine;
    super(`Invalid RTF at ${line}:${column}: ${message}`);
    this.name = "RtfParseError";
    this.context = context;
  }
}
