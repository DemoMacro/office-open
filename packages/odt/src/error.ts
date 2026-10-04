/** Structured error raised when an ODT package cannot be parsed. */
export class OdtParseError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = new.target.name;
  }
}
