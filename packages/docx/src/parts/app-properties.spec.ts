import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { AppPropertiesInput } from "./app-properties";
import { appPropertiesDesc } from "./app-properties";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: AppPropertiesInput) {
  const xml = appPropertiesDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return appPropertiesDesc.parse(el, readCtx) as Record<string, unknown>;
}

describe("appPropertiesDesc round-trip", () => {
  it("round-trips string properties", () => {
    const result = roundTrip({
      template: "Normal.dotm",
      manager: "Alice",
      company: "Acme Corp",
      hyperlinkBase: "https://example.com",
      application: "Microsoft Office Word",
      appVersion: "16.0000",
    });
    expect(result.template).toBe("Normal.dotm");
    expect(result.manager).toBe("Alice");
    expect(result.company).toBe("Acme Corp");
    expect(result.hyperlinkBase).toBe("https://example.com");
    expect(result.application).toBe("Microsoft Office Word");
    expect(result.appVersion).toBe("16.0000");
  });

  it("round-trips number properties", () => {
    const result = roundTrip({
      pages: 3,
      words: 420,
      characters: 2500,
      lines: 40,
      paragraphs: 12,
      slides: 0,
      notes: 1,
      totalTime: 15,
      hiddenSlides: 0,
      mmClips: 2,
      charactersWithSpaces: 2900,
      docSecurity: 1,
    });
    expect(result.pages).toBe(3);
    expect(result.words).toBe(420);
    expect(result.characters).toBe(2500);
    expect(result.lines).toBe(40);
    expect(result.paragraphs).toBe(12);
    expect(result.slides).toBe(0);
    expect(result.notes).toBe(1);
    expect(result.totalTime).toBe(15);
    expect(result.hiddenSlides).toBe(0);
    expect(result.mmClips).toBe(2);
    expect(result.charactersWithSpaces).toBe(2900);
    expect(result.docSecurity).toBe(1);
  });

  it("round-trips boolean properties (emitted as 0/1)", () => {
    const result = roundTrip({
      scaleCrop: false,
      linksUpToDate: true,
      sharedDoc: true,
      hyperlinksChanged: false,
    });
    expect(result.scaleCrop).toBe(false);
    expect(result.linksUpToDate).toBe(true);
    expect(result.sharedDoc).toBe(true);
    expect(result.hyperlinksChanged).toBe(false);
  });

  it("round-trips a full document app-properties set", () => {
    const result = roundTrip({
      template: "Normal.dotm",
      company: "Acme",
      pages: 1,
      words: 10,
      characters: 60,
      paragraphs: 1,
      lines: 1,
      notes: 1,
      totalTime: 0,
      application: "Microsoft Office Word",
      appVersion: "16.0000",
      docSecurity: 0,
      scaleCrop: false,
      linksUpToDate: false,
      sharedDoc: false,
      hyperlinksChanged: false,
    });
    expect(result.template).toBe("Normal.dotm");
    expect(result.company).toBe("Acme");
    expect(result.pages).toBe(1);
    expect(result.words).toBe(10);
    expect(result.application).toBe("Microsoft Office Word");
    expect(result.appVersion).toBe("16.0000");
    expect(result.docSecurity).toBe(0);
  });

  it("emits the extended-properties root element with vt namespace", () => {
    const xml = appPropertiesDesc.stringify({ application: "Word" }, writeCtx)!;
    expect(xml).toContain(
      'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"',
    );
    expect(xml).toContain(
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"',
    );
  });

  it("round-trips empty properties", () => {
    const result = roundTrip({});
    expect(Object.keys(result).length).toBe(0);
  });

  it("round-trips special XML characters in string fields", () => {
    const result = roundTrip({ company: 'A <B> & "C"' });
    expect(result.company).toBe('A <B> & "C"');
  });
});
