import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { buildCorePropertiesXmlString, parseCorePropsElement } from "./core";

const NS =
  'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
  'xmlns:dcterms="http://purl.org/dc/terms/" ' +
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"';

describe("core properties", () => {
  it("round-trips category, contentStatus, identifier, language, version", () => {
    const opts = {
      category: "Word Perf. Complex",
      contentStatus: "Draft",
      identifier: "urn:uuid:1",
      language: "en-US",
      version: "1.2",
    };
    const xml = buildCorePropertiesXmlString(opts);
    expect(xml).toContain("<cp:category>Word Perf. Complex</cp:category>");
    expect(xml).toContain("<cp:contentStatus>Draft</cp:contentStatus>");
    expect(xml).toContain("<dc:identifier>urn:uuid:1</dc:identifier>");
    expect(xml).toContain("<dc:language>en-US</dc:language>");
    expect(xml).toContain("<cp:version>1.2</cp:version>");

    const el = parseXml(xml).elements?.[0];
    const back = parseCorePropsElement(el);
    expect(back).toMatchObject(opts);
  });

  it("keeps an empty category as an empty string (presence-based)", () => {
    // Word writes whitespace-only category elements; the parser reduces them to
    // an empty element and parse must capture "" so the field survives.
    const xml = `<cp:coreProperties ${NS}><cp:category></cp:category></cp:coreProperties>`;
    const el = parseXml(xml).elements?.[0];
    const back = parseCorePropsElement(el);
    expect(back.category).toBe("");

    const rebuilt = buildCorePropertiesXmlString(back);
    expect(rebuilt).toContain("<cp:category></cp:category>");
  });

  it("omits the new fields when not supplied", () => {
    const xml = buildCorePropertiesXmlString({ title: "T" });
    expect(xml).not.toContain("cp:category");
    expect(xml).not.toContain("cp:contentStatus");
    expect(xml).not.toContain("dc:identifier");
    expect(xml).not.toContain("dc:language");
    expect(xml).not.toContain("cp:version");
  });
});
