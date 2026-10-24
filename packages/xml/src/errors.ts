/** Structured error thrown when input is not well-formed XML. */
export class XmlParseError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "XmlParseError";
  }
}
