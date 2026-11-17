import { describe, expect, it } from "vitest";

import { XlsParseError } from "../errors";
import { applyXfExt, parseXf } from "./styles";

function xfExtRecord(
  properties: readonly Uint8Array[],
  index = 0,
): {
  code: number;
  start: number;
  body: Uint8Array;
} {
  const length = 20 + properties.reduce((total, property) => total + property.byteLength, 0);
  const body = new Uint8Array(length);
  const view = new DataView(body.buffer);
  view.setUint16(0, 0x087d, true);
  view.setUint16(14, index, true);
  view.setUint16(18, properties.length, true);
  let offset = 20;
  for (const property of properties) {
    body.set(property, offset);
    offset += property.byteLength;
  }
  return { code: 0x087d, start: 0, body };
}

function fullColorProperty(type: number, colorType: number, value: number): Uint8Array {
  const property = new Uint8Array(20);
  const view = new DataView(property.buffer);
  view.setUint16(0, type, true);
  view.setUint16(2, 20, true);
  view.setUint16(4, colorType, true);
  view.setUint32(8, value, true);
  return property;
}

function byteProperty(type: number, value: number): Uint8Array {
  return new Uint8Array([type, 0, 5, 0, value]);
}

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

  it("projects XFExt colors and font scheme into the canonical style", () => {
    const style = parseXf(biff5Xf(), 5, []);
    const styles = new Map([[0, style]]);
    applyXfExt(
      xfExtRecord([
        fullColorProperty(0x0004, 2, 0xff996633),
        fullColorProperty(0x000d, 3, 1),
        byteProperty(0x000e, 2),
        new Uint8Array([0x0f, 0x00, 0x06, 0x00, 0x03, 0x00]),
      ]),
      styles,
      [],
    );
    expect(style.fill?.foregroundColor).toEqual({ rgb: "FF336699" });
    expect(style.font).toEqual({ color: { theme: 1 }, scheme: "minor" });
    expect(style.alignment?.indent).toBe(3);
  });

  it("rejects unknown XFExt properties", () => {
    const style = parseXf(biff5Xf(), 5, []);
    expect(() =>
      applyXfExt(xfExtRecord([new Uint8Array([0x10, 0, 4, 0])]), new Map([[0, style]]), []),
    ).toThrow(XlsParseError);
  });
});
