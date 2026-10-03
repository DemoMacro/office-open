export class RtfParseError extends Error {
  constructor(message: string, position: number, source: string) {
    const before = source.slice(0, position);
    const line = (before.match(/\r\n|\r|\n/g)?.length ?? 0) + 1;
    const lastLine = Math.max(before.lastIndexOf("\n"), before.lastIndexOf("\r"));
    const column = position - lastLine;
    super(`Invalid RTF at ${line}:${column}: ${message}`);
    this.name = "RtfParseError";
  }
}
