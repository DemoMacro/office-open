import type { CompoundFileReader } from "@office-open/core/cfb";
import { describe, expect, it } from "vitest";

import { readSummaryInformation } from "./container";

function documentSummaryInformation(category: string): Uint8Array {
  const value = new Uint8Array(80);
  const view = new DataView(value.buffer);
  view.setUint16(0, 0xfffe, true);
  view.setUint32(24, 1, true);
  view.setUint32(44, 48, true);
  view.setUint32(48, 32, true);
  view.setUint32(52, 1, true);
  view.setUint32(56, 2, true);
  view.setUint32(60, 16, true);
  view.setUint32(64, 30, true);
  view.setUint32(68, category.length + 1, true);
  for (let index = 0; index < category.length; index += 1) {
    value[72 + index] = category.charCodeAt(index);
  }
  return value;
}

describe("DOC summary information", () => {
  it("reads category from the separate document summary stream", () => {
    const path = `${String.fromCharCode(5)}DocumentSummaryInformation`;
    const stream = documentSummaryInformation("Reports");
    const reader = {
      entry: (name: string) => (name === path ? {} : undefined),
      read: (name: string) => (name === path ? stream : new Uint8Array()),
    } as unknown as CompoundFileReader;

    expect(readSummaryInformation(reader)).toEqual({ category: "Reports" });
  });
});
