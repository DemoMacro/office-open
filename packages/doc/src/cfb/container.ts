import { parseSummaryInformation } from "@office-open/core";
import type { CompoundFileReader } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";

export function readSummaryInformation(
  reader: CompoundFileReader,
): Omit<DocumentOptions, "sections"> {
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  return reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {};
}
