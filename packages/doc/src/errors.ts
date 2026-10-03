/** Legacy DOC input cannot be decoded safely or is not supported. */
export class DocParseError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "DocParseError";
  }
}
