/** Structured error raised for an ODF element that cannot be mapped canonically. */
export class OdfSchemaError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly name: string,
    readonly reason: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
