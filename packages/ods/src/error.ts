/** Structured error raised when an ODS package cannot be parsed. */
export class OdsParseError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = new.target.name;
  }
}
