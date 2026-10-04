import { expect, it } from "vitest";

import { RecordType, readRecordTree } from ".";

it("reads a container record and its atom child", () => {
  const bytes = new Uint8Array(36);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 15, true);
  view.setUint16(2, RecordType.document, true);
  view.setUint32(4, 28, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, RecordType.documentAtom, true);
  view.setUint32(12, 20, true);
  view.setInt32(16, 3200, true);
  view.setInt32(20, 2400, true);

  const [document] = readRecordTree(view, 0, bytes.byteLength);

  expect(document?.type).toBe(RecordType.document);
  expect(document?.children).toHaveLength(1);
  expect(document?.children[0]?.type).toBe(RecordType.documentAtom);
  expect(document?.children[0]?.length).toBe(20);
});
