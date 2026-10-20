import { parse } from "@office-open/xml";
import { describe, expect, it } from "vitest";

import { fillDesc } from "./fill-descriptors";

describe("fillDesc unknown blip extensions", () => {
  it("round-trips extension-list content", () => {
    const xml =
      '<a:blipFill xmlns:a="urn:a" rotWithShape="0">' +
      '<a:blip r:embed="rId1" xmlns:r="urn:r">' +
      '<a:extLst><a:ext uri="urn:test"><test:value/></a:ext></a:extLst>' +
      "</a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>";
    const options = fillDesc.parse(parse(xml), {
      resolveRelationship: () => undefined,
      getRaw: () => undefined,
    } as never);
    expect(options).toMatchObject({
      type: "blip",
      rotWithShape: false,
      blipExtensionXml: '<a:ext uri="urn:test"><test:value/></a:ext>',
    });
    expect(fillDesc.stringify(options!, {} as never)).toContain(
      '<a:extLst><a:ext uri="urn:test"><test:value/></a:ext></a:extLst>',
    );
  });
});
describe("fillDesc relative rectangles", () => {
  it("round-trips a gradient fill with an omitted stop list", () => {
    const options = fillDesc.parse(parse('<a:gradFill xmlns:a="urn:a"/>'), {} as never);
    expect(options).toEqual({ type: "gradient", options: {} });
    expect(fillDesc.stringify(options!, {} as never)).toBe("<a:gradFill/>");
  });

  it("preserves explicit zero gradient tile edges", () => {
    const xml =
      '<a:gradFill xmlns:a="urn:a"><a:gsLst>' +
      '<a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>' +
      '<a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs>' +
      '</a:gsLst><a:lin ang="0"/><a:tileRect l="0" t="0" r="0" b="0"/></a:gradFill>';

    const options = fillDesc.parse(parse(xml), {} as never);
    expect(options).toMatchObject({
      type: "gradient",
      options: { tileRectangle: { left: "0", top: "0", right: "0", bottom: "0" } },
    });
    expect(fillDesc.stringify(options!, {} as never)).toContain(
      '<a:tileRect l="0" t="0" r="0" b="0"/>',
    );
  });
});
