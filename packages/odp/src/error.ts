/** Structured error raised when an ODP package cannot be parsed. */
export class OdpParseError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = new.target.name;
  }
}
