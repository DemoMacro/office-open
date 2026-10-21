/** Word revision insert/delete markers decoded from FRD ranges. */
export interface LegacyRevisionRange {
  /** Start character position in the main document. */
  start: number;
  /** End character position, exclusive. */
  end: number;
  /** True for inserted text; false for deleted text. */
  inserted: boolean;
}

/** A named bookmark range decoded from PlcfBkf and PlcfBkl. */
export interface LegacyBookmark {
  id: number;
  name: string;
  start: number;
  end: number;
}

/** Main-document field instruction decoded from PlcffldMom. */
export interface LegacyField {
  start: number;
  end: number;
  instruction: string;
}

/** Picture bytes and dimensions decoded from the Escher BStore container. */
export interface LegacyPictureData {
  type: "png" | "jpeg" | "gif" | "tiff" | "bmp" | "ico" | "emf" | "wmf" | "unknown";
  data: Uint8Array;
  width: number;
  height: number;
}

/** A picture shape resolved from the OfficeArt drawing group by its shape id. */
export type LegacyPictureByShapeId = Map<number, LegacyPictureData>;

/** Password for password-protected legacy DOC containers. */
export interface LegacyParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}
