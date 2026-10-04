import type { DocumentOptions } from "@office-open/docx";

import { DocParseError } from "./errors";
import { parseInternal } from "./mappers/document";

export interface DocParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

/** Input is not a byte sequence, so no CFB path can be established. */
export function parseDocument(data: Uint8Array, options?: DocParseOptions): DocumentOptions {
  if (!(data instanceof Uint8Array)) {
    throw new DocParseError("DOC input must be a Uint8Array", {
      part: "container",
      path: "/",
      reason: "invalid-input-type",
    });
  }
  return parseInternal(data, options?.password);
}
