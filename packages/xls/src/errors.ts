/** Legacy XLS input cannot be decoded safely or is not supported. */
export class LegacyExcelError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "LegacyExcelError";
  }
}
