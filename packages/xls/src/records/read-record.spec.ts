import { expect, it } from "vitest";

import { readRecord, recordHeaderSize } from ".";

it("reads a BIFF8 record boundary and body", () => {
  const bytes = new Uint8Array(14);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 0x0203, true);
  view.setUint16(2, 6, true);
  view.setUint32(4, 1, true);

  const record = readRecord(bytes, 0, 8);

  expect(record.code).toBe(0x0203);
  expect(record.start).toBe(0);
  expect(record.body).toEqual(bytes.subarray(4, 10));
  expect(recordHeaderSize(8)).toBe(4);
});
