import type { SectionChild } from "@office-open/docx";

import { DocParseError } from "../errors";
import { readUint16, readUint32, requireRange } from "./binary";
import { FLAG_ENCRYPTED } from "./constants";
import { WINDOWS_1252_HIGH } from "./text";
export interface LegacyFib {
  fcMin: number;
  fcMac: number;
  ccpText: number;
}

export function parseLegacyFib(word: Uint8Array): LegacyFib {
  const message = "Invalid legacy Word document: FIB is truncated";
  requireRange(word, 0, 88, message);
  const nFib = readUint16(word, 2, message);
  if (
    readUint16(word, 0, message) !== 0xa5dc ||
    nFib < 101 ||
    nFib > 105 ||
    (readUint16(word, 10, message) & FLAG_ENCRYPTED) !== 0
  ) {
    throw new DocParseError("Invalid Word document: unsupported FIB signature", {
      part: "stream",
      path: "WordDocument",
      recordName: "FIB",
      offset: 0,
      length: 2,
      reason: "unsupported-required-structure",
    });
  }
  return {
    fcMin: readUint32(word, 24, message),
    fcMac: readUint32(word, 28, message),
    ccpText: readUint32(word, 52, message),
  };
}

export function parseLegacyText(word: Uint8Array, fib: LegacyFib): string {
  requireRange(
    word,
    fib.fcMin,
    fib.fcMac - fib.fcMin,
    "Invalid legacy Word document: text range is outside the stream",
  );
  let result = "";
  for (let index = fib.fcMin; index < fib.fcMac; index += 1) {
    const byte = word[index]!;
    result += WINDOWS_1252_HIGH.get(byte) ?? String.fromCharCode(byte);
  }
  return result;
}

export function legacyChildren(text: string): SectionChild[] {
  const children: SectionChild[] = [];
  for (const value of text.split(/\r/)) {
    children.push({ paragraph: value });
  }
  return children;
}
