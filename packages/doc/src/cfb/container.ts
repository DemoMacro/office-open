import { parseDocumentSummaryInformation, parseSummaryInformation } from "@office-open/core/cfb";
import type { CompoundFileReader } from "@office-open/core/cfb";
import type { DocumentOptions } from "@office-open/docx";

import { DocParseError } from "../errors";

export function readSummaryInformation(
  reader: CompoundFileReader,
): Omit<DocumentOptions, "sections"> {
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  const documentPath = `${String.fromCharCode(5)}DocumentSummaryInformation`;
  try {
    return {
      ...(reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {}),
      ...(reader.entry(documentPath)
        ? parseDocumentSummaryInformation(reader.read(documentPath))
        : {}),
    };
  } catch (error) {
    throw new DocParseError(
      `Invalid Word document summary information: ${(error as Error).message}`,
      { part: "stream", path, reason: "invalid-metadata" },
      { cause: error },
    );
  }
}
