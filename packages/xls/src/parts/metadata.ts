import {
  CompoundFileReader,
  parseDocumentSummaryInformation,
  parseSummaryInformation,
} from "@office-open/core/cfb";
import type { WorkbookOptions } from "@office-open/xlsx";

import { XlsParseError } from "../errors";

export function readSummaryInformation(data: Uint8Array): WorkbookOptions {
  if (!isCompoundFile(data)) return {};
  const reader = new CompoundFileReader(data);
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  const documentPath = `${String.fromCharCode(5)}DocumentSummaryInformation`;
  const result = reader.entry(path)
    ? readSummaryStream(path, reader.read(path), parseSummaryInformation)
    : {};
  if (reader.entry(documentPath)) {
    const values = readSummaryStream(
      documentPath,
      reader.read(documentPath),
      parseDocumentSummaryInformation,
    );
    if (values.category !== undefined) result.category = values.category;
  }
  return result;
}

function isCompoundFile(data: Uint8Array): boolean {
  return (
    data.byteLength >= 8 &&
    data[0] === 0xd0 &&
    data[1] === 0xcf &&
    data[2] === 0x11 &&
    data[3] === 0xe0 &&
    data[4] === 0xa1 &&
    data[5] === 0xb1 &&
    data[6] === 0x1a &&
    data[7] === 0xe1
  );
}

function readSummaryStream(
  path: string,
  data: Uint8Array,
  parse: (data: Uint8Array) => WorkbookOptions,
): WorkbookOptions {
  try {
    return parse(data);
  } catch (error) {
    throw new XlsParseError("Invalid legacy XLS file: summary information stream is malformed", {
      part: "stream",
      path,
      reason: "malformed-summary-information",
      cause: error,
    });
  }
}
