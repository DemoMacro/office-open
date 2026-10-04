import { CompoundFileReader, parseSummaryInformation } from "@office-open/core";
import type { PresentationOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";

export function readSummaryInformation(reader: CompoundFileReader): PresentationOptions {
  const path = `${String.fromCharCode(5)}SummaryInformation`;
  try {
    return reader.entry(path) ? parseSummaryInformation(reader.read(path)) : {};
  } catch (error) {
    throw new PptParseError(
      `Invalid legacy PowerPoint summary information: ${(error as Error).message}`,
      { part: "stream", path, reason: "invalid-metadata" },
    );
  }
}

export function assertCfbSignature(data: Uint8Array): void {
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (data.byteLength < signature.length || signature.some((byte, index) => data[index] !== byte)) {
    throw new PptParseError("Input is not a legacy PowerPoint CFB container", {
      part: "container",
      path: "/",
      offset: 0,
      length: signature.length,
      reason: "invalid-container-signature",
    });
  }
}

export function readStream(reader: CompoundFileReader, path: string): Uint8Array {
  try {
    return reader.read(path);
  } catch {
    throw new PptParseError(
      `Invalid legacy PowerPoint CFB container: missing or unreadable "${path}" stream`,
      {
        part: "stream",
        path,
        reason: "missing-required-stream",
      },
    );
  }
}
