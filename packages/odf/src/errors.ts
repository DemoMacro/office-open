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
export class OcfManifestError extends OcfError {
  constructor(message: string, options?: ErrorOptions & { fullPath?: string }) {
    super(message, options === undefined ? undefined : { cause: options.cause });
    if (options?.fullPath !== undefined) this.fullPath = options.fullPath;
  }

  /** Entry path that caused a manifest conflict, when applicable. */
  readonly fullPath?: string;
}

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
