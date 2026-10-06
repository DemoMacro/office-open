import { describe, expect, it } from "vitest";

import { decodeXmlBytes } from "./xml-decode";

describe("decodeXmlBytes", () => {
  it("decodes UTF-8 XML", () => {
    expect(decodeXmlBytes(new TextEncoder().encode("<root/>"))).toBe("<root/>");
  });

  it("decodes UTF-16 XML by BOM", () => {
    const xml = '<?xml version="1.0" encoding="UTF-16"?><root/>';
    const bytes = [0xff, 0xfe];
    for (const character of xml) bytes.push(character.charCodeAt(0), character.charCodeAt(0) >> 8);
    const payload = Uint8Array.from(bytes);
    expect(decodeXmlBytes(payload)).toBe(xml);
  });

  it("decodes BOM-less UTF-16LE XML by declaration", () => {
    const xml = '<?xml version="1.0" encoding="UTF-16"?><root/>';
    const bytes: number[] = [];
    for (const character of xml) bytes.push(character.charCodeAt(0), character.charCodeAt(0) >> 8);
    const payload = Uint8Array.from(bytes);
    expect(decodeXmlBytes(payload)).toBe(xml);
  });
});
