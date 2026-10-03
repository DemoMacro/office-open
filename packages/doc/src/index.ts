import type { DocumentOptions } from "@office-open/docx";

import { DocParseError } from "./errors";
import { parseInternal } from "./parser";

export { DocParseError };

/** Parse a legacy Microsoft Word 97-2003 binary document. */
export function parseDocument(data: Uint8Array): DocumentOptions {
  if (!(data instanceof Uint8Array)) throw new DocParseError("DOC input must be a Uint8Array");
  return parseInternal(data);
}
