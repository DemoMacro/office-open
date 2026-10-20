/** Structured error for malformed or inconsistent Compound File Binary data. */
export class CfbError extends Error {
  constructor(
    message: string,
    readonly reason = "invalid-cfb",
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "CfbError";
  }
}
