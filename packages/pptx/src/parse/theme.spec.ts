import { xml2js } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseTheme, parseColorScheme, parseFontScheme } from "./theme";

const THEME_XML = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="333333"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

const THEME_SYSCLR_XML = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface="SimHei"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface="SimSun"/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

describe("parseTheme", () => {
  it("parses theme name", () => {
    const el = xml2js(THEME_XML).elements![0];
    const result = parseTheme(el);
    expect(result.name).to.equal("Custom Theme");
  });

  it("parses srgbClr colors", () => {
    const el = xml2js(THEME_XML).elements![0];
    const result = parseTheme(el);
    expect(result.colors).to.exist;
    expect(result.colors!.dark1).to.equal("333333");
    expect(result.colors!.light1).to.equal("FFFFFF");
    expect(result.colors!.accent1).to.equal("4472C4");
    expect(result.colors!.hyperlink).to.equal("0563C1");
  });

  it("parses sysClr colors (lastClr)", () => {
    const el = xml2js(THEME_SYSCLR_XML).elements![0];
    const result = parseTheme(el);
    expect(result.colors!.dark1).to.equal("000000");
    expect(result.colors!.light1).to.equal("FFFFFF");
  });

  it("parses font scheme", () => {
    const el = xml2js(THEME_XML).elements![0];
    const result = parseTheme(el);
    expect(result.fonts).to.exist;
    expect(result.fonts!.majorFont).to.equal("Calibri Light");
    expect(result.fonts!.minorFont).to.equal("Calibri");
  });

  it("parses asian fonts", () => {
    const el = xml2js(THEME_SYSCLR_XML).elements![0];
    const result = parseTheme(el);
    expect(result.fonts!.majorFontAsian).to.equal("SimHei");
    expect(result.fonts!.minorFontAsian).to.equal("SimSun");
  });
});

describe("parseColorScheme", () => {
  it("parses all 12 color entries", () => {
    const el = xml2js(THEME_XML).elements![0];
    const themeElements = el.elements![0];
    const clrScheme = themeElements.elements![0];
    const result = parseColorScheme(clrScheme);
    const keys = Object.keys(result);
    expect(keys.length).to.equal(12);
  });
});

describe("parseFontScheme", () => {
  it("parses major and minor fonts", () => {
    const el = xml2js(THEME_XML).elements![0];
    const themeElements = el.elements![0];
    const fontScheme = themeElements.elements![1];
    const result = parseFontScheme(fontScheme);
    expect(result.majorFont).to.equal("Calibri Light");
    expect(result.minorFont).to.equal("Calibri");
  });
});
