import { describe, expect, it } from "vitest";

import type { Piece } from "../streams/pieces";
import { parseParagraphFkp } from "./fkp";

describe("parseParagraphFkp", () => {
  it("maps one PAPX FC range across compressed and Unicode pieces to one CP range", () => {
    const page = 512;
    const word = new Uint8Array(1024);
    const view = new DataView(word.buffer);
    view.setUint32(page, 100, true);
    view.setUint32(page + 4, 108, true);
    view.setUint8(page + 8, 20);
    view.setUint8(page + 40, 0);
    view.setUint8(page + 41, 4);
    view.setUint8(page + 42, 0);
    view.setUint8(page + 43, 0);
    word.set([0x03, 0x24, 0x01, 0, 0, 0], page + 44);
    view.setUint8(page + 511, 1);
    const pieces: Piece[] = [
      { cpStart: 0, cpEnd: 2, fc: 100, compressed: true },
      { cpStart: 2, cpEnd: 4, fc: 104, compressed: false },
    ];

    expect(parseParagraphFkp(word, page, pieces, 4)).toEqual([
      { cpStart: 0, cpEnd: 4, properties: { alignment: "center" } },
    ]);
  });
});
