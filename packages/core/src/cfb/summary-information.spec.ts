import { describe, expect, it } from "vitest";

import { parseDocumentSummaryInformation, parseSummaryInformation } from "./summary-information";

const SUMMARY_PROPERTY_SET = [
  0xf2, 0x9f, 0x85, 0xe0, 0x4f, 0xf9, 0x10, 0x68, 0xab, 0x91, 0x08, 0x00, 0x2b, 0x27, 0xb3, 0x29,
] as const;

const DOCUMENT_SUMMARY_PROPERTY_SET = [
  0xd5, 0xcd, 0xd5, 0x02, 0x2e, 0x9c, 0x10, 0x16, 0x39, 0xc1, 0x00, 0x00, 0x6b, 0x92, 0x42, 0x62,
] as const;

function byteNumber(value: number): readonly number[] {
  return Array.from({ length: 4 }, (_, index) => (value >>> (index * 8)) & 0xff);
}

function zeroBytes(length: number): readonly number[] {
  return Array.from({ length }, () => 0);
}

function uint16(value: number): readonly number[] {
  return [value & 0xff, (value >>> 8) & 0xff];
}

function guid(guid: readonly number[]): readonly number[] {
  return [...guid];
}

function propertyValue(
  id: number,
  value: readonly number[],
): {
  id: number;
  value: readonly number[];
} {
  return { id, value };
}

function lpString(value: string): readonly number[] {
  return [
    30,
    0,
    0,
    0,
    ...byteNumber(value.length),
    ...Array.from(value, (character) => character.charCodeAt(0)),
  ];
}

function fileTime(iso: string): readonly number[] {
  const time = Date.parse(iso) * 10_000 + 116444736000000000;
  return [64, 0, 0, 0, ...byteNumber(time), ...byteNumber(Math.floor(time / 2 ** 32))];
}

function propertySet(values: readonly { id: number; value: readonly number[] }[]): Uint8Array {
  const padded = values.map((item) => ({
    id: item.id,
    value: [...item.value, ...Array.from({ length: (4 - (item.value.length % 4)) % 4 }, () => 0)],
  }));
  const sectionSize = 8 + padded.reduce((total, item) => total + 8 + item.value.length, 0);
  const descriptors: number[] = [];
  const bodies: number[] = [];
  let offset = 8 + padded.length * 8;
  for (const item of padded) {
    descriptors.push(...byteNumber(item.id), ...byteNumber(offset));
    bodies.push(...item.value);
    offset += item.value.length;
  }
  const section = [
    ...byteNumber(sectionSize),
    ...byteNumber(padded.length),
    ...descriptors,
    ...bodies,
  ];
  return new Uint8Array(section);
}

describe("CFB summary information", () => {
  it("reads summary properties and file times", () => {
    const section = propertySet([
      propertyValue(2, lpString("Title")),
      propertyValue(4, lpString("Author")),
      propertyValue(9, [3, 0, 0, 0, 7, 0, 0, 0]),
      propertyValue(15, fileTime("2026-01-02T03:04:05.000Z")),
    ]);
    const stream = new Uint8Array([
      ...uint16(0xfffe),
      ...uint16(0),
      ...byteNumber(0),
      ...zeroBytes(16),
      ...byteNumber(1),
      ...guid(SUMMARY_PROPERTY_SET),
      ...byteNumber(48),
      ...section,
    ]);
    expect(parseSummaryInformation(stream)).toEqual({
      title: "Title",
      creator: "Author",
      revision: 7,
      created: "2026-01-02T03:04:05Z",
    });
  });

  it("reads a separate document summary category", () => {
    const section = propertySet([propertyValue(2, lpString("Reports"))]);
    const stream = new Uint8Array([
      ...uint16(0xfffe),
      ...uint16(0),
      ...byteNumber(0),
      ...zeroBytes(16),
      ...byteNumber(1),
      ...guid(DOCUMENT_SUMMARY_PROPERTY_SET),
      ...byteNumber(48),
      ...section,
    ]);

    expect(parseDocumentSummaryInformation(stream)).toEqual({ category: "Reports" });
  });

  it("reads both summary property sets", () => {
    const summarySection = propertySet([
      propertyValue(2, lpString("Title")),
      propertyValue(4, lpString("Author")),
    ]);
    const documentSection = propertySet([propertyValue(2, lpString("Reports"))]);
    const summaryOffset = 68;
    const documentOffset = summaryOffset + summarySection.byteLength;
    const stream = new Uint8Array([
      ...uint16(0xfffe),
      ...uint16(0),
      ...byteNumber(0),
      ...zeroBytes(16),
      ...byteNumber(2),
      ...guid(SUMMARY_PROPERTY_SET),
      ...byteNumber(summaryOffset),
      ...guid(DOCUMENT_SUMMARY_PROPERTY_SET),
      ...byteNumber(documentOffset),
      ...summarySection,
      ...documentSection,
    ]);

    expect(parseSummaryInformation(stream)).toEqual({
      title: "Title",
      creator: "Author",
      category: "Reports",
    });
  });
});
