import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { settingsDesc } from "./descriptor";
import type { CompatibilityOptions, SettingsOptions } from "./settings";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: SettingsOptions): SettingsOptions {
  const xml = settingsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return settingsDesc.parse(el, readCtx);
}

describe("settingsDesc round-trip", () => {
  it("round-trips view", () => {
    const result = roundTrip({ view: "print" });
    expect(result.view).toBe("print");
  });

  it("round-trips zoom percent", () => {
    const result = roundTrip({ zoom: { percent: 150 } });
    const zoom = result.zoom!;
    expect(zoom.percent).toBe(150);
  });

  it("round-trips zoom val", () => {
    const result = roundTrip({ zoom: { val: "fullPage" } });
    const zoom = result.zoom!;
    expect(zoom.val).toBe("fullPage");
  });

  it("round-trips defaultTabStop", () => {
    const result = roundTrip({ defaultTabStop: 720 });
    expect(result.defaultTabStop).toBe(720);
  });

  it("round-trips characterSpacingControl", () => {
    const result = roundTrip({ characterSpacingControl: "doNotCompress" });
    expect(result.characterSpacingControl).toBe("doNotCompress");
  });

  it("round-trips trackRevisions", () => {
    const result = roundTrip({ trackRevisions: true });
    expect(result.trackRevisions).toBe(true);
  });

  it("round-trips updateFields", () => {
    const result = roundTrip({ updateFields: true });
    expect(result.updateFields).toBe(true);
  });

  it("round-trips compatibility version via compatSetting", () => {
    const result = roundTrip({ compatibility: { version: 15 } });
    expect((result.compatibility as CompatibilityOptions).version).toBe(15);
  });

  it("emits MS Office default compatSettings for fresh documents", () => {
    const xml = settingsDesc.stringify({}, writeCtx)!;
    expect(xml).toContain(
      'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"',
    );
    expect(xml).toContain('w:name="overrideTableStyleFontSizeAndJustification"');
    expect(xml).toContain('w:name="enableOpenTypeFeatures"');
    expect(xml).toContain('w:name="doNotFlipMirrorIndents"');
  });

  it("omits <w:compat> when compatibility is false", () => {
    const xml = settingsDesc.stringify({ compatibility: false }, writeCtx)!;
    expect(xml).not.toContain("<w:compat>");
    expect(roundTrip({ compatibility: false }).compatibility).toBe(false);
  });

  it("emits only user-specified compat fields without merging fresh defaults", () => {
    const xml = settingsDesc.stringify({ compatibility: { version: 14 } }, writeCtx)!;
    expect(xml).toContain("compatibilityMode");
    expect(xml).toContain('w:val="14"');
    expect(xml).not.toContain("enableOpenTypeFeatures");
    expect(xml).not.toContain("doNotFlipMirrorIndents");
  });

  it("round-trips docVars", () => {
    const result = roundTrip({
      docVars: [
        { name: "var1", val: "value1" },
        { name: "var2", val: "value2" },
      ],
    });
    const docVars = result.docVars!;
    expect(docVars).toHaveLength(2);
    expect(docVars[0]?.name).toBe("var1");
    expect(docVars[0]?.val).toBe("value1");
    expect(docVars[1]?.name).toBe("var2");
    expect(docVars[1]?.val).toBe("value2");
  });

  it("round-trips displayBackgroundShape", () => {
    const result = roundTrip({ displayBackgroundShape: true });
    expect(result.displayBackgroundShape).toBe(true);
  });

  it("round-trips embedTrueTypeFonts", () => {
    const result = roundTrip({ embedTrueTypeFonts: true });
    expect(result.embedTrueTypeFonts).toBe(true);
  });

  it("round-trips embedSystemFonts", () => {
    const result = roundTrip({ embedSystemFonts: true });
    expect(result.embedSystemFonts).toBe(true);
  });

  it("round-trips saveSubsetFonts", () => {
    const result = roundTrip({ saveSubsetFonts: true });
    expect(result.saveSubsetFonts).toBe(true);
  });

  it("round-trips evenAndOddHeaders", () => {
    const result = roundTrip({ evenAndOddHeaders: true });
    expect(result.evenAndOddHeaders).toBe(true);
  });

  it("round-trips themeFontLang with eastAsia", () => {
    const result = roundTrip({
      themeFontLang: { val: "en-US", eastAsia: "zh-CN" },
    });
    expect(result.themeFontLang?.val).toBe("en-US");
    expect(result.themeFontLang?.eastAsia).toBe("zh-CN");
  });

  it("round-trips compat compatSettings (newer flags)", () => {
    const result = roundTrip({
      compatibility: {
        compatSettings: [{ name: "differentiateMultirowTableHeaders", val: "1" }],
      },
    });
    expect((result.compatibility as CompatibilityOptions).compatSettings).toEqual([
      {
        name: "differentiateMultirowTableHeaders",
        val: "1",
        uri: "http://schemas.microsoft.com/office/word",
      },
    ]);
  });

  it("round-trips documentProtection edit", () => {
    const result = roundTrip({ documentProtection: { edit: "readOnly" } });
    expect(result.documentProtection?.edit).toBe("readOnly");
  });

  it("round-trips writeProtection recommended", () => {
    const result = roundTrip({ writeProtection: { recommended: true } });
    expect(result.writeProtection?.recommended).toBe(true);
  });

  it("round-trips combined options", () => {
    const result = roundTrip({
      view: "web",
      zoom: { percent: 100, val: "none" },
      defaultTabStop: 420,
      trackRevisions: true,
      characterSpacingControl: "compressPunctuation",
    });
    expect(result.view).toBe("web");
    const zoom = result.zoom!;
    expect(zoom.percent).toBe(100);
    expect(result.defaultTabStop).toBe(420);
    expect(result.trackRevisions).toBe(true);
  });

  it("round-trips empty settings", () => {
    const result = roundTrip({});
    // optional elements must not be invented when absent (faithful round-trip)
    expect(result.defaultTabStop).toBeUndefined();
    expect(result.characterSpacingControl).toBeUndefined();
  });

  it("drops unmapped non-standard elements (no verbatim rawXml fallback)", () => {
    // Parse is fully structured and aligned with generate; elements outside
    // SettingsOptions (like WPS's wpsCustomData) are dropped rather than
    // captured verbatim.
    const src =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData">' +
      '<w:defaultTabStop w:val="720"/>' +
      '<wpsCustomData:typoFeatureVersion wpsCustomData:val="1"/>' +
      "</w:settings>";
    const el = parseXml(src).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const parsed = settingsDesc.parse(el, readCtx);

    // Structured read still works.
    expect(parsed.defaultTabStop).toBe(720);
    // Unmapped non-standard element is dropped (no rawXml capture).
    expect(parsed.rawXml).toBeUndefined();

    // Re-emit uses structured generation (fixed SETTINGS_NS).
    const xml = settingsDesc.stringify(parsed, writeCtx)!;
    expect(xml).toContain('<w:defaultTabStop w:val="720"/>');
    expect(xml).not.toContain("wpsCustomData");
  });

  it("uses structured generation when rawXml is absent", () => {
    const xml = settingsDesc.stringify({ defaultTabStop: 720 }, writeCtx)!;
    expect(xml).toContain('<w:defaultTabStop w:val="720"/>');
  });

  it("round-trips shapeDefaults with structured o: content", () => {
    const result = roundTrip({
      hdrShapeDefaults: { shapedefaults: { ext: "edit", spidmax: 2050 } },
      shapeDefaults: {
        shapedefaults: {
          ext: "edit",
          spidmax: 3074,
          colormru: { ext: "edit", colors: "212121" },
        },
        shapelayout: { ext: "edit", idmap: { ext: "edit", data: "1" } },
      },
    });
    expect(result.hdrShapeDefaults).toEqual({ shapedefaults: { ext: "edit", spidmax: 2050 } });
    expect(result.shapeDefaults).toEqual({
      shapedefaults: { ext: "edit", spidmax: 3074, colormru: { ext: "edit", colors: "212121" } },
      shapelayout: { ext: "edit", idmap: { ext: "edit", data: "1" } },
    });
  });
});
