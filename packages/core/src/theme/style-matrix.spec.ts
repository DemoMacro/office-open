import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseFormatScheme } from "./style-matrix";

describe("format scheme", () => {
  it("ignores whitespace nodes between style-matrix children", () => {
    const xml =
      '<a:fmtScheme xmlns:a="urn:a" name="Test"><a:fillStyleLst>' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      "<a:noFill/>" +
      "<a:blipFill><a:blip/></a:blipFill>" +
      "</a:fillStyleLst></a:fmtScheme>";
    const el = parseXml(xml).elements?.[0];
    const parsed = parseFormatScheme(el, {} as never);
    expect(parsed?.fillStyles).toHaveLength(3);
    expect(parsed?.fillStyles?.[0]).toEqual({ type: "solid", color: { value: "phClr" } });
  });
});
