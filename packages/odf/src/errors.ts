/** Base error for invalid OpenDocument Container Format packages. */
export class OcfError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = new.target.name;
  }
}

/** The mimetype entry does not identify the expected ODF document family. */
export class OcfMimeTypeError extends OcfError {
  constructor(
    readonly expected: string,
    readonly actual: string,
  ) {
    super(`Unexpected ODF MIME type: ${actual || "missing"}`);
  }
}

/** META-INF/manifest.xml is missing or structurally invalid. */
export class OcfManifestError extends OcfError {}

/** A required XML part is absent or cannot be parsed. */
export class OdfXmlError extends OcfError {
  constructor(
    readonly path: string,
    options?: ErrorOptions,
  ) {
    super(`ODF package is missing or cannot read ${path}`, options);
  }
}

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
