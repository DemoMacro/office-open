import { CompoundFileReader, parseSummaryInformation } from "@office-open/core/cfb";
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
      { cause: error },
    );
  }
}

export function assertCfbSignature(data: Uint8Array): void {
  if (
    data.byteLength >= 8 &&
    data[0] === 0xed &&
    data[1] === 0xde &&
    data[2] === 0xad &&
    data[3] === 0x0b
  ) {
    throw new PptParseError("Unsupported legacy PowerPoint 2/3 file", {
      part: "container",
      path: "/",
      offset: 0,
      length: 8,
      reason: "unsupported-legacy-format",
    });
  }
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

export function isLegacyPp40Reader(reader: CompoundFileReader): boolean {
  return reader.entry("PP40") !== undefined && reader.entry("PowerPoint Document") === undefined;
}

export function readStream(reader: CompoundFileReader, path: string): Uint8Array {
  if (!reader.entry(path)) {
    throw new PptParseError(`Invalid legacy PowerPoint CFB container: missing "${path}" stream`, {
      part: "stream",
      path,
      reason: "missing-required-stream",
    });
  }
  try {
    return reader.read(path);
  } catch (cause) {
    throw new PptParseError(
      `Invalid legacy PowerPoint CFB container: unreadable "${path}" stream`,
      {
        part: "stream",
        path,
        reason: "invalid-stream",
      },
      { cause },
    );
  }
}
