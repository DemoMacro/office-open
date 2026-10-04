/** Base error for invalid ODF schema structures. */
export class OdfSchemaError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = new.target.name;
  }
}
