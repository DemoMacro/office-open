import { parseSummaryInformation } from "@office-open/core";
import type { CompoundFileReader } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";

import { DocParseError } from "../errors";

export function readSummaryInformation(
  reader: CompoundFileReader,
): Omit<DocumentOptions, "sections"> {
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  try {
    return reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {};
  } catch (error) {
    throw new DocParseError(
      `Invalid Word document summary information: ${(error as Error).message}`,
      { part: "stream", path, reason: "invalid-metadata" },
    );
  }
}
