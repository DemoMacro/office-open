import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { settingsDesc } from "./descriptor";
import type { SettingsOptions } from "./settings";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

interface SettingsParseResult extends SettingsOptions {
  compatabilityModeVersion?: number;
  evenAndOddHeaderAndFooters?: boolean;
}

function roundTrip(opts: SettingsOptions) {
  const xml = settingsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return settingsDesc.parse(el, readCtx) as unknown as SettingsParseResult;
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

  it("round-trips compatibilityModeVersion via compatSetting", () => {
    const result = roundTrip({ compatibilityModeVersion: 15 });
    expect(result.compatabilityModeVersion).toBe(15);
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
    expect(docVars[0].name).toBe("var1");
    expect(docVars[0].val).toBe("value1");
    expect(docVars[1].name).toBe("var2");
    expect(docVars[1].val).toBe("value2");
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
    expect(result.evenAndOddHeaderAndFooters).toBe(true);
  });

  it("round-trips documentProtection with password", () => {
    // Verify stringify produces the XML correctly — parse doesn't support documentProtection
    const opts: SettingsOptions = {
      documentProtection: {
        edit: "readOnly",
        password: "test",
      },
    };
    const xml = settingsDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain("w:documentProtection");
    expect(xml).toContain('w:edit="readOnly"');
  });

  it("round-trips writeProtection", () => {
    const opts: SettingsOptions = {
      writeProtection: { recommended: true },
    };
    const xml = settingsDesc.stringify(opts, writeCtx)!;
    expect(xml).toContain("w:writeProtection");
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

  it("captures the full settings part verbatim with source namespaces", () => {
    // CT_Settings has ~100 element types; the verbatim rawXml capture preserves
    // all of them (including non-standard namespaces like WPS's wpsCustomData)
    // for byte-exact round-trip, while structured reads still run in parallel.
    const src =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData">' +
      '<w:defaultTabStop w:val="720"/>' +
      '<wpsCustomData:typoFeatureVersion wpsCustomData:val="1"/>' +
      "</w:settings>";
    const el = parseXml(src).elements![0];
    const parsed = settingsDesc.parse(el, readCtx) as unknown as SettingsParseResult;

    // Structured read still works alongside verbatim capture.
    expect(parsed.defaultTabStop).toBe(720);
    // Verbatim inner content + root attributes captured.
    expect(parsed.rawXml).toContain("wpsCustomData:typoFeatureVersion");
    expect(parsed.rootAttributes?.["xmlns:wpsCustomData"]).toBe(
      "http://www.wps.cn/officeDocument/2013/wpsCustomData",
    );

    // Re-emit verbatim: source namespace + custom element preserved.
    const xml = settingsDesc.stringify(parsed as unknown as SettingsOptions, writeCtx)!;
    expect(xml).toContain("xmlns:wpsCustomData");
    expect(xml).toContain("wpsCustomData:typoFeatureVersion");
  });

  it("uses structured generation when rawXml is absent", () => {
    const xml = settingsDesc.stringify({ defaultTabStop: 720 }, writeCtx)!;
    expect(xml).toContain('<w:defaultTabStop w:val="720"/>');
  });
});
