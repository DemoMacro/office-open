/**
 * Content Types components for OPC (Open Packaging Convention).
 *
 * Provides Default and Override elements used in [Content_Types].xml
 * to map file extensions and part paths to MIME content types.
 *
 * Reference: ECMA-376 Part 2, Section 10.1
 *
 * @module
 */

export interface DefaultAttributes {
  readonly contentType: string;
  readonly extension?: string;
}

/**
 * Creates a Default element mapping a file extension to a MIME content type.
 */
export const createDefault = (contentType: string, extension?: string): string =>
  extension !== undefined
    ? `<Default ContentType="${contentType}" Extension="${extension}"/>`
    : `<Default ContentType="${contentType}"/>`;

export interface OverrideAttributes {
  readonly contentType: string;
  readonly partName?: string;
}

/**
 * Creates an Override element mapping a specific part path to a MIME content type.
 */
export const createOverride = (contentType: string, partName?: string): string =>
  partName !== undefined
    ? `<Override ContentType="${contentType}" PartName="${partName}"/>`
    : `<Override ContentType="${contentType}"/>`;
