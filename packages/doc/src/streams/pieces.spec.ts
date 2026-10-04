import { expect, it } from "vitest";

import { parsePieceTable } from "./pieces";

it("parses a compressed piece-table boundary", () => {
  const table = new Uint8Array(33);
  const view = new DataView(table.buffer);
  table[0] = 2;
  view.setUint32(1, 28, true);
  view.setUint32(5, 0, true);
  view.setUint32(9, 2, true);
  view.setUint32(13, 4, true);
  view.setUint32(19, 0x40000000, true);
  view.setUint32(27, 0x40000008, true);
  const pieces = parsePieceTable(table, { offset: 0, length: table.byteLength });

  expect(pieces).toEqual([
    { cpStart: 0, cpEnd: 2, fc: 0, compressed: true },
    { cpStart: 2, cpEnd: 4, fc: 4, compressed: true },
  ]);
});
