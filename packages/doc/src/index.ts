import { DocParseError } from "./errors";
import { parseInternal } from "./parser";
import type { LegacyDocumentOptions } from "./types";

export { DocParseError };

export type {
  LegacyBookmark,
  LegacyDocumentOptions,
  LegacyField,
  LegacyPictureData,
  LegacyRevisionRange,
} from "./types";

/** Parse a legacy Microsoft Word 97-2003 binary document. */
export function parseDocument(data: Uint8Array): LegacyDocumentOptions {
  if (!(data instanceof Uint8Array)) throw new DocParseError("DOC input must be a Uint8Array");
  return parseInternal(data);
}
