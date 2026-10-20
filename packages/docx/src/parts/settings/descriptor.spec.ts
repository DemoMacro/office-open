import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { documentNamespaceDialect } from "../document/document-attributes";
import { mailMergeRecipientsDesc, settingsDesc } from "./descriptor";
import type { CompatibilityOptions, MailMergeRecipientsOptions, SettingsOptions } from "./settings";

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as Parameters<typeof documentNamespaceDialect>[0];

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

const unsupportedWriteCtx: WriteContext = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  addHyperlink: () => {},
};

function roundTrip(opts: SettingsOptions): SettingsOptions {
  const xml = settingsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return settingsDesc.parse(el, readCtx);
}

function roundTripXml(opts: SettingsOptions): string {
  return settingsDesc.stringify(roundTrip(opts), writeCtx)!;
}

describe("settingsDesc round-trip", () => {
  it("round-trips document identifier inline extension text", () => {
    const result = roundTrip({
      w14DocId: "646534E6",
      w14DocIdText: "Hello World!",
      w15DocId: "{37FF56B8-11AD-4A0A-8C3F-F4BD7E1B24BE}",
      w15DocIdText: "metadata",
    });
    expect(result.w14DocIdText).toBe("Hello World!");
    expect(result.w15DocIdText).toBe("metadata");
    expect(roundTripXml({ w14DocId: "A", w14DocIdText: "Hello World!" })).toContain(
      '<w14:docId w14:val="A">Hello World!</w14:docId>',
    );
  });

  it("round-trips w14:conflictMode", () => {
    const result = roundTrip({ w14ConflictMode: true });
    expect(result.w14ConflictMode).toBe(true);
  });

  it("round-trips mirror margins with the source on/off lexical form", () => {
    const result = roundTrip({ mirrorMargins: "off" });
    expect(result.mirrorMargins).toBe("off");
    expect(roundTripXml({ mirrorMargins: true })).toContain("<w:mirrorMargins/>");
    expect(roundTripXml({ mirrorMargins: false })).toContain('<w:mirrorMargins w:val="0"/>');
  });

  it("round-trips vendor tm settings in source order", () => {
    const vendor: SettingsOptions = {
      compatibility: false,
      tmPrefOne: 17,
      tmPrefTwo: 1,
      tmFormatPreference: 55063659,
      tmCommentsProperties: { place: 0, width: 3120, color: -1 },
      tmReviewProperties: {
        enabled: 0,
        show: 1,
        print: 0,
        revisionNumber: 11,
        insertionMark: 4,
        insertionColor: -1,
        deletionMark: 6,
        deletionColor: -1,
        formattingMark: 1,
        formattingColor: -1,
        lineMark: 1,
        lineColor: 0,
        tooltip: 1,
      },
      tmLastPosition: {
        page: 0,
        selection: 4,
        frameIndex: 0,
        caret: { paragraphIndex: 0, index: 0 },
        anchor: { paragraphIndex: 0, index: 0 },
        tableRectangle: { left: 0, top: 0, right: 0, bottom: 0 },
      },
      tmApplicationRevision: { date: 1533054491, val: 766 },
    };
    const result = roundTrip(vendor);
    expect(result).toEqual(vendor);

    const xml = roundTripXml(result);
    const tags = [
      "w:tmPrefOne",
      "w:tmPrefTwo",
      "w:tmFmtPref",
      "w:tmCommentsPr",
      "w:tmReviewPr",
      "w:tmLastPos",
      "w:tmAppRevision",
    ].map((tag) => xml.indexOf(tag));
    expect(tags).toEqual([...tags].sort((left, right) => left - right));
    expect(xml).toContain('<w:tmLastPosTblRect w:left="0" w:top="0" w:right="0" w:bottom="0"/>');
  });

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

  it("round-trips empty no-line-break attributes", () => {
    const result = roundTrip({ noLineBreaksBefore: { lang: "", val: "" } });
    expect(result.noLineBreaksBefore).toEqual({ lang: "", val: "" });
  });

  it("round-trips defaultTabStop", () => {
    const result = roundTrip({ defaultTabStop: 720 });
    expect(result.defaultTabStop).toBe(720);
  });

  it("round-trips defaultTabStop measure literal", () => {
    const result = roundTrip({ defaultTabStop: "36pt" });
    expect(result.defaultTabStop).toBe("36pt");
    expect(roundTripXml({ defaultTabStop: "36pt" })).toContain('<w:defaultTabStop w:val="36pt"/>');
  });

  it("round-trips hyphenation and drawing-grid measure literals", () => {
    const result = roundTrip({
      hyphenationZone: "21.25pt",
      drawingGridHorizontalSpacing: "14.20pt",
      drawingGridVerticalSpacing: 280,
    });
    expect(result.hyphenationZone).toBe("21.25pt");
    expect(result.drawingGridHorizontalSpacing).toBe("14.20pt");
    expect(result.drawingGridVerticalSpacing).toBe(280);
    expect(roundTripXml(result)).toContain('<w:hyphenationZone w:val="21.25pt"/>');
    expect(roundTripXml(result)).toContain('<w:drawingGridHorizontalSpacing w:val="14.20pt"/>');
  });

  it("parses revision save IDs as typed values", () => {
    const xml =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:rsids><w:rsidRoot w:val="00000000"/><w:rsid w:val="AF0B12C0"/></w:rsids>' +
      "</w:settings>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed settings has no root element");
    const result = settingsDesc.parse(el, readCtx);
    expect(result.rsids).toEqual({ rsidRoot: "00000000", rsids: ["AF0B12C0"] });
    const roundTripped = roundTrip(result);
    expect(roundTripped.rsids).toEqual({ rsidRoot: "00000000", rsids: ["AF0B12C0"] });
  });

  it("round-trips a bare rsid element without a value attribute", () => {
    const xml =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:rsids><w:rsid/><w:rsid w:val="AF0B12C0"/></w:rsids>' +
      "</w:settings>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed settings has no root element");
    const result = settingsDesc.parse(el, readCtx);
    expect(result.rsids).toEqual({ rsids: ["", "AF0B12C0"] });
    expect(settingsDesc.stringify(result, writeCtx)).toContain(
      '<w:rsids><w:rsid/><w:rsid w:val="AF0B12C0"/></w:rsids>',
    );
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

  it("preserves explicit compatibility false values and empty mail-merge strings", () => {
    const xml =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:compat><w:useFELayout w:val="0"/>' +
      '<w:doNotExpandShiftReturn w:val="0"/>' +
      '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>' +
      "</w:compat>" +
      '<w:mailMerge><w:connectString w:val=""/></w:mailMerge>' +
      "</w:settings>";

    const options = settingsDesc.parse(parseXml(xml).elements![0]!, readCtx)!;
    expect(options.compatibility).toMatchObject({
      useFELayout: false,
      doNotExpandShiftReturn: false,
      doNotFlipMirrorIndents: false,
    });
    expect(options.mailMerge).toEqual({ connectString: "" });

    const output = settingsDesc.stringify(options, writeCtx)!;
    expect(output).toContain('<w:useFELayout w:val="0"/>');
    expect(output).toContain('<w:doNotExpandShiftReturn w:val="0"/>');
    expect(output).toContain('w:name="doNotFlipMirrorIndents" w:uri=');
    expect(output).toContain('w:val="0"/>');
    expect(output).toContain('<w:connectString w:val=""/>');
  });

  it("folds a settings-root compatSetting into the compatibility domain", () => {
    // Word 2010 transitional documents can strand the entry outside w:compat.
    const xml =
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:compat><w:useFELayout/></w:compat>` +
      `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>` +
      `</w:settings>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const parsed = settingsDesc.parse(el, readCtx);
    expect(parsed.compatibility).toMatchObject({ useFELayout: true, version: 14 });

    const out = settingsDesc.stringify(parsed, writeCtx)!;
    expect(out).toContain(
      '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>',
    );
  });

  it("merges duplicate compatibility containers on parse", () => {
    const xml =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:compat><w:doNotExpandShiftReturn/></w:compat>" +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/></w:compat>' +
      "</w:settings>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const parsed = settingsDesc.parse(el, readCtx);
    expect(parsed.compatibility).toMatchObject({
      doNotExpandShiftReturn: true,
      version: 12,
    });

    const output = settingsDesc.stringify(parsed, writeCtx)!;
    expect(output.match(/<w:compat>/g)).toHaveLength(1);
    expect(output).toContain("<w:doNotExpandShiftReturn/>");
    expect(output).toContain('w:name="compatibilityMode"');
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

  it("retains duplicate known compatSettings", () => {
    const xml =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:compat>" +
      '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>' +
      '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>' +
      "</w:compat></w:settings>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const parsed = settingsDesc.parse(el, readCtx);
    expect(parsed.compatibility).toMatchObject({
      overrideTableStyleFontSizeAndJustification: false,
    });
    expect((parsed.compatibility as CompatibilityOptions).compatSettings).toHaveLength(2);
    const output = settingsDesc.stringify(parsed, writeCtx)!;
    expect(output.match(/w:name="overrideTableStyleFontSizeAndJustification"/g)).toHaveLength(2);
    expect(output).toContain('w:val="0"');
    expect(output).toContain('w:val="1"');
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

  it("preserves noncanonical lexical values for on/off settings", () => {
    const source =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:displayBackgroundShape w:val="on"/>' +
      '<w:autoHyphenation w:val="t"/>' +
      '<w:bordersDoNotSurroundFooter w:val="false"/>' +
      "</w:settings>";
    const el = parseXml(source).elements?.[0];
    if (!el) throw new Error("no root");
    const parsed = settingsDesc.parse(el, readCtx);
    expect(parsed.displayBackgroundShape).toBe("on");
    expect(parsed.autoHyphenation).toBe("t");
    expect(parsed.bordersDoNotSurroundFooter).toBe("false");

    const xml = settingsDesc.stringify(parsed, writeCtx)!;
    expect(xml).toContain('<w:displayBackgroundShape w:val="on"/>');
    expect(xml).toContain('<w:autoHyphenation w:val="t"/>');
    expect(xml).toContain('<w:bordersDoNotSurroundFooter w:val="false"/>');
  });

  it("round-trips documentProtection enforcement as false", () => {
    const source =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:documentProtection w:enforcement="0" w:edit="readOnly"/>' +
      "</w:settings>";
    const element = parseXml(source).elements?.[0];
    if (!element) throw new Error("parsed settings has no root element");
    const parsed = settingsDesc.parse(element, readCtx);
    expect(parsed.documentProtection).toEqual({
      enforcement: false,
      enforcementRaw: "0",
      edit: "readOnly",
    });

    const xml = settingsDesc.stringify(parsed, writeCtx)!;
    expect(xml).toContain('<w:documentProtection w:enforcement="0" w:edit="readOnly"/>');
  });

  it("preserves source-absent documentProtection enforcement", () => {
    const source =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:documentProtection w:edit="forms"/>' +
      "</w:settings>";
    const element = parseXml(source).elements?.[0];
    if (!element) throw new Error("parsed settings has no root element");
    const result = settingsDesc.parse(element, readCtx);
    expect(result.documentProtection).toMatchObject({ edit: "forms", enforcement: null });
    expect(settingsDesc.stringify(result, writeCtx)).toContain(
      '<w:documentProtection w:edit="forms"/>',
    );
    expect(settingsDesc.stringify({ documentProtection: { edit: "forms" } }, writeCtx)).toContain(
      '<w:documentProtection w:enforcement="1" w:edit="forms"/>',
    );
  });

  it("retains empty legacy documentProtection credentials", () => {
    const source =
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:documentProtection w:enforcement="1" w:edit="readOnly" w:cryptAlgorithmSid="" w:hash="" w:salt=""/>' +
      "</w:settings>";
    const element = parseXml(source).elements?.[0];
    if (!element) throw new Error("parsed settings has no root element");
    const parsed = settingsDesc.parse(element, readCtx);
    expect(parsed.documentProtection).toMatchObject({
      cryptoAlgorithmSidRaw: "",
      hashRaw: "",
      saltRaw: "",
    });
    const xml = settingsDesc.stringify(parsed, writeCtx)!;
    expect(xml).toContain('w:cryptAlgorithmSid=""');
    expect(xml).toContain('w:hash=""');
    expect(xml).toContain('w:salt=""');
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

  it("drops unmapped non-standard elements (no verbatim fallback)", () => {
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

    // Re-emit uses structured generation (fixed SETTINGS_NS).
    const xml = settingsDesc.stringify(parsed, writeCtx)!;
    expect(xml).toContain('<w:defaultTabStop w:val="720"/>');
    expect(xml).not.toContain("wpsCustomData");
  });

  it("generates structured settings XML", () => {
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
      childOrder: ["shapedefaults", "shapelayout"],
      shapedefaults: { ext: "edit", spidmax: 3074, colormru: { ext: "edit", colors: "212121" } },
      shapelayout: { ext: "edit", idmap: { ext: "edit", data: "1" } },
    });
  });

  it("preserves shapeDefaults o: child order", () => {
    const opts: SettingsOptions = {
      shapeDefaults: {
        childOrder: ["shapelayout", "shapedefaults"],
        shapelayout: { ext: "edit", idmap: { ext: "edit", data: "1" } },
        shapedefaults: { ext: "edit", spidmax: 3074 },
      },
    };
    const xml = settingsDesc.stringify(opts, writeCtx)!;
    const layoutStart = xml.indexOf("<o:shapelayout");
    const defaultsStart = xml.indexOf("<o:shapedefaults");
    expect(layoutStart).toBeGreaterThan(0);
    expect(defaultsStart).toBeGreaterThan(layoutStart);

    const roundTripped = roundTripXml(opts);
    expect(roundTripped.indexOf("<o:shapelayout")).toBeGreaterThan(0);
    expect(roundTripped.indexOf("<o:shapedefaults")).toBeGreaterThan(
      roundTripped.indexOf("<o:shapelayout"),
    );
  });
});

describe("settings compat/style-pane fidelity", () => {
  it("emits only the given compat fields — no injected compatibilityMode", () => {
    const xml = settingsDesc.stringify({ compatibility: { useFELayout: true } }, writeCtx)!;
    expect(xml).toContain("<w:useFELayout/>");
    expect(xml).not.toContain("compatibilityMode");
  });

  it("round-trips the legacy stylePaneFormatFilter hex bitmask", () => {
    const result = roundTrip({
      stylePaneFormatFilter: { allStyles: true, val: "0001" },
    });
    expect(result.stylePaneFormatFilter).toEqual({ allStyles: true, val: "0001" });
  });

  it("round-trips w14:enableOpenTypeKerning inside w:compat", () => {
    const result = roundTrip({
      compatibility: { useFELayout: true, enableOpenTypeKerning: true },
    });
    expect(result.compatibility).toMatchObject({
      useFELayout: true,
      enableOpenTypeKerning: true,
    });
    const xml = settingsDesc.stringify(
      { compatibility: { enableOpenTypeKerning: true } },
      writeCtx,
    )!;
    expect(xml).toContain("<w14:enableOpenTypeKerning/>");
  });
});

describe("mailMergeRecipientsDesc round-trip", () => {
  it("round-trips active flags, columns, and unique tags", () => {
    const opts = {
      recipients: [
        { active: false, column: 0, uniqueTag: "AAAA" },
        { column: 2, uniqueTag: "BB==" },
      ],
    };
    const xml = mailMergeRecipientsDesc.stringify(opts, unsupportedWriteCtx)!;
    expect(xml).toContain('<w:active w:val="0"/>');
    expect(xml).toContain('<w:column w:val="2"/>');
    expect(xml).toContain('<w:uniqueTag w:val="BB=="/>');
    expect(mailMergeRecipientsDesc.parse(parseXml(xml).elements![0]!, readCtx)).toEqual(opts);
  });

  it("round-trips the legacy wne recipient vocabulary", () => {
    const opts: MailMergeRecipientsOptions = {
      namespace: "wne",
      recipients: [{ active: true, hash: "65" }, { hash: "66" }],
    };
    const xml = mailMergeRecipientsDesc.stringify(opts, unsupportedWriteCtx)!;
    expect(xml).toContain("<wne:recipients");
    expect(xml).toContain('<wne:recipientData><wne:active wne:val="1"/><wne:hash wne:val="65"/>');
    expect(mailMergeRecipientsDesc.parse(parseXml(xml).elements![0]!, readCtx)).toEqual(opts);
  });
});
