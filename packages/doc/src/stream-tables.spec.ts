import { describe, expect, it } from "vitest";

import { DocParseError } from "./errors";
import { parseHeaderStreams, parseTextboxBoundaries } from "./mappers/stream-tables";

describe("DOC story boundary tables", () => {
  it("rejects non-monotonic header CP boundaries", () => {
    const table = Uint8Array.from([3, 0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 0, 0]);

    expect(() => parseHeaderStreams(table, { offset: 0, length: 16 }, 4)).toThrow(DocParseError);
  });

  it("parses the FTXBXS shape identifier", () => {
    const table = new Uint8Array(48);
    const view = new DataView(table.buffer);
    view.setUint32(0, 0, true);
    view.setUint32(4, 3, true);
    view.setUint32(8, 6, true);
    view.setUint32(26, 7001, true);

    expect(parseTextboxBoundaries(table, { offset: 0, length: 48 }, 6)).toEqual([
      { start: 0, end: 3, shapeId: 7001 },
    ]);
  });
});
