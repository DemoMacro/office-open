import { CompoundFileReader } from "@office-open/core";

import { LegacyExcelError } from "../errors";

export function readCompoundFile(data: Uint8Array): CompoundFileReader {
  return new CompoundFileReader(data);
}

export function readRequiredStream(reader: CompoundFileReader, path: string): Uint8Array {
  const entry = reader.entry(path);
  if (!entry || entry.type !== "stream") {
    throw new LegacyExcelError(`Invalid legacy XLS file: missing or unreadable ${path} stream`);
  }
  return reader.read(entry.path);
}
