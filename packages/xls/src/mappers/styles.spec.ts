import { describe, expect, it } from "vitest";

import { XlsParseError } from "../errors";
import { parseXf } from "./styles";

function biff5Xf(): Uint8Array {
  const body = new Uint8Array(16);
  const view = new DataView(body.buffer);
  view.setUint16(2, 164, true);
  view.setUint16(6, 0x0012, true);
  body[7] = 0x5c;
  view.setUint32(8, 0x00010008, true);
  return body;
}

describe("parseXf", () => {
  it("parses the BIFF5 alignment, fill and border layout", () => {
    const style = parseXf(biff5Xf(), 5, ["FF0000ff"]);
    expect(style).toMatchObject({
      fontIndex: 0,
      numberFormatId: 164,
      alignment: { horizontal: "center", vertical: "center" },
      fill: {
        type: "pattern",
        patternType: "solid",
        foregroundColor: { rgb: "FF0000ff" },
      },
    });
  });

  it("rejects truncated BIFF5 formats", () => {
    expect(() => parseXf(new Uint8Array(12), 5, [])).toThrow(XlsParseError);
  });
});
