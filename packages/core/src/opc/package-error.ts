/** Structured error for unreadable or inconsistent OPC ZIP packages. */
export class OpcPackageError extends Error {
  constructor(
    message: string,
    readonly part: string | undefined,
    readonly reason: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OpcPackageError";
  }
}
