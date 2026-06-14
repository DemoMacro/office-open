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
  features?: { trackRevisions?: boolean; updateFields?: boolean };
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
    const features = result.features!;
    expect(features.trackRevisions).toBe(true);
  });

  it("round-trips updateFields", () => {
    const result = roundTrip({ updateFields: true });
    const features = result.features!;
    expect(features.updateFields).toBe(true);
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
    const features = result.features!;
    expect(features.trackRevisions).toBe(true);
  });

  it("round-trips empty settings", () => {
    const result = roundTrip({});
    // optional elements must not be invented when absent (faithful round-trip)
    expect(result.defaultTabStop).toBeUndefined();
    expect(result.characterSpacingControl).toBeUndefined();
  });
});
