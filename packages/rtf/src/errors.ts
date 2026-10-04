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
