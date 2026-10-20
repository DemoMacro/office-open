import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { buildThemeXml } from "./build-theme-xml";
import { parseFormatScheme } from "./style-matrix";
import type { ThemeOptions } from "./theme-options";

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

describe("format scheme blip fills", () => {
  it("preserves an optional absent blip", () => {
    const xml =
      '<a:fmtScheme xmlns:a="urn:a" name="Test"><a:bgFillStyleLst>' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      '<a:blipFill rotWithShape="1"><a:tile tx="0" ty="0"/></a:blipFill>' +
      "</a:bgFillStyleLst></a:fmtScheme>";
    const el = parseXml(xml).elements?.[0];
    const parsed = parseFormatScheme(el, {} as never);
    expect(parsed?.backgroundFillStyles?.[2]).toMatchObject({
      type: "blip",
      noEmbed: true,
      omitBlip: true,
      tile: { tx: 0, ty: 0 },
      stretch: false,
    });
    const options: ThemeOptions = { formatScheme: parsed };
    const generatedXml = buildThemeXml(options, {} as never);
    expect(generatedXml).not.toMatch(/<a:blip(?:\s|>)/u);
    expect(generatedXml).toContain('<a:blipFill rotWithShape="1">');
  });
});
