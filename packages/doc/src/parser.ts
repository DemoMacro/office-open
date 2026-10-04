import type { DocumentOptions } from "@office-open/docx";

import { DocParseError } from "./errors";
import { parseInternal } from "./mappers/document";

export interface DocParseOptions {
  /** Password used to verify and decrypt legacy Office RC4 containers. */
  password?: string;
}

export function parseDocument(data: Uint8Array, options?: DocParseOptions): DocumentOptions {
  if (!(data instanceof Uint8Array)) throw new DocParseError("DOC input must be a Uint8Array");
  return parseInternal(data, options?.password);
}
