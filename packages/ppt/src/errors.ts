export class PptParseError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "PptParseError";
  }
}
