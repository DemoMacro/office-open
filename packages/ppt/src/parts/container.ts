import { CompoundFileReader, parseSummaryInformation } from "@office-open/core";
import type { PresentationOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";

export function readSummaryInformation(reader: CompoundFileReader): PresentationOptions {
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  return reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {};
}

export function assertCfbSignature(data: Uint8Array): void {
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (data.byteLength < signature.length || signature.some((byte, index) => data[index] !== byte)) {
    throw new PptParseError("Input is not a legacy PowerPoint CFB container");
  }
}

export function readStream(reader: CompoundFileReader, path: string): Uint8Array {
  try {
    return reader.read(path);
  } catch {
    throw new PptParseError(
      `Invalid legacy PowerPoint CFB container: missing or unreadable "${path}" stream`,
    );
  }
}
