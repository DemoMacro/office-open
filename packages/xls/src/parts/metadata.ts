import {
  CompoundFileReader,
  parseDocumentSummaryInformation,
  parseSummaryInformation,
} from "@office-open/core";
import type { WorkbookOptions } from "@office-open/xlsx";

export function readSummaryInformation(data: Uint8Array): WorkbookOptions {
  try {
    const reader = new CompoundFileReader(data);
    const path = `${String.fromCharCode(5)}SummaryInformation`;
    const documentPath = `${String.fromCharCode(5)}DocumentSummaryInformation`;
    const result = reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {};
    if (reader.entry(documentPath)) {
      const values = parseDocumentSummaryInformation(reader.read(documentPath));
      if (values.category !== undefined) result.category = values.category;
    }
    return result;
  } catch {
    return {};
  }
}
