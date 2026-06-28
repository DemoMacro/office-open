import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { AltChunkOptions } from "@parts/alt-chunk/alt-chunk";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../context";
import {
  altChunkDesc,
  customXmlBlockDesc,
  sdtBlockDesc,
  setBodyParseChild,
  subDocDesc,
} from "./bodychildren";
import type { SdtBlockOptions } from "./bodychildren";
import type { CustomXmlBlockDescriptorOptions } from "./bodychildren";

// Register a minimal child parser for SDT/customXml
setBodyParseChild((el, _ctx) => {
  if (el.name === "w:p") return { paragraph: {} } as any;
  return { paragraph: {} } as any;
});

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
  stringifyChild: (child: unknown) =>
    typeof child === "string" ? `<w:p><w:r><w:t>${child}</w:t></w:r></w:p>` : "<w:p/>",
  fileData: {
    document: {
      relationships: { addRelationship: () => {} },
    },
    altChunks: { addAltChunk: () => {} },
    subDocs: { addSubDoc: () => {} },
  },
} as unknown as BodyContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

// ── sdtBlockDesc ──

function roundTripSdt(opts: SdtBlockOptions) {
  const xml = sdtBlockDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return sdtBlockDesc.parse(el, readCtx);
}

describe("sdtBlockDesc round-trip", () => {
  it("round-trips SDT with alias", () => {
    const result = roundTripSdt({
      properties: { alias: "MyAlias" },
      children: [],
    });
    expect(result.properties.alias).toBe("MyAlias");
  });

  it("round-trips SDT with tag", () => {
    const result = roundTripSdt({
      properties: { tag: "my-tag" },
      children: [],
    });
    expect(result.properties.tag).toBe("my-tag");
  });

  it("round-trips SDT with id", () => {
    const result = roundTripSdt({
      properties: { id: 42 },
      children: [],
    });
    expect(result.properties.id).toBe(42);
  });

  it("round-trips SDT with lock", () => {
    const result = roundTripSdt({
      properties: { lock: "sdtLocked" },
      children: [],
    });
    expect(result.properties.lock).toBe("sdtLocked");
  });

  it("round-trips SDT with temporary", () => {
    const result = roundTripSdt({
      properties: { temporary: true },
      children: [],
    });
    expect(result.properties.temporary).toBe(true);
  });

  it("round-trips SDT with showingPlaceholder", () => {
    const result = roundTripSdt({
      properties: { showingPlaceholder: true },
      children: [],
    });
    expect(result.properties.showingPlaceholder).toBe(true);
  });

  it("round-trips SDT with equation type", () => {
    const result = roundTripSdt({
      properties: { equation: true },
      children: [],
    });
    expect(result.properties.equation).toBe(true);
  });

  it("round-trips SDT with comboBox", () => {
    const result = roundTripSdt({
      properties: {
        comboBox: {
          items: [
            { displayText: "Option A", value: "a" },
            { displayText: "Option B", value: "b" },
          ],
          lastValue: "a",
        },
      },
      children: [],
    });
    expect(result.properties.comboBox).toBeDefined();
    expect(result.properties.comboBox!.items).toHaveLength(2);
    expect(result.properties.comboBox!.items![0].displayText).toBe("Option A");
    expect(result.properties.comboBox!.lastValue).toBe("a");
  });

  it("round-trips SDT with dropDownList", () => {
    const result = roundTripSdt({
      properties: {
        dropDownList: {
          items: [
            { displayText: "Yes", value: "1" },
            { displayText: "No", value: "0" },
          ],
        },
      },
      children: [],
    });
    expect(result.properties.dropDownList).toBeDefined();
    expect(result.properties.dropDownList!.items).toHaveLength(2);
  });

  it("round-trips SDT with text type", () => {
    const result = roundTripSdt({
      properties: { text: { multiLine: true } },
      children: [],
    });
    expect(result.properties.text).toBeDefined();
    expect(result.properties.text!.multiLine).toBe(true);
  });

  it("round-trips SDT with dataBinding", () => {
    const result = roundTripSdt({
      properties: {
        dataBinding: {
          xpath: "/root/data",
          storeItemID: "{12345}",
          prefixMappings: "xmlns:ns='http://example.com'",
        },
      },
      children: [],
    });
    expect(result.properties.dataBinding).toBeDefined();
    expect(result.properties.dataBinding!.xpath).toBe("/root/data");
    expect(result.properties.dataBinding!.storeItemID).toBe("{12345}");
  });

  it("round-trips SDT with label and tabIndex", () => {
    const result = roundTripSdt({
      properties: { label: 100, tabIndex: 5 },
      children: [],
    });
    expect(result.properties.label).toBe(100);
    expect(result.properties.tabIndex).toBe(5);
  });

  it("round-trips SDT with picture type", () => {
    const result = roundTripSdt({
      properties: { picture: true },
      children: [],
    });
    expect(result.properties.picture).toBe(true);
  });

  it("round-trips SDT with richText type", () => {
    const result = roundTripSdt({
      properties: { richText: true },
      children: [],
    });
    expect(result.properties.richText).toBe(true);
  });

  it("round-trips SDT with citation type", () => {
    const result = roundTripSdt({
      properties: { citation: true },
      children: [],
    });
    expect(result.properties.citation).toBe(true);
  });

  it("round-trips SDT with combined properties", () => {
    const result = roundTripSdt({
      properties: {
        alias: "Form",
        tag: "form-field",
        id: 99,
        lock: "contentLocked",
        temporary: true,
      },
      children: [],
    });
    expect(result.properties.alias).toBe("Form");
    expect(result.properties.tag).toBe("form-field");
    expect(result.properties.id).toBe(99);
    expect(result.properties.lock).toBe("contentLocked");
    expect(result.properties.temporary).toBe(true);
  });

  it("round-trips SDT with checkbox (default symbols)", () => {
    const result = roundTripSdt({
      properties: { alias: "Agree", checkbox: { checked: true } },
    });
    expect(result.properties.checkbox?.checked).toBe(true);
    expect(result.properties.checkbox?.checkedState?.val).toBe("2612");
    expect(result.properties.checkbox?.uncheckedState?.val).toBe("2610");
  });

  it("round-trips SDT with checkbox unchecked + custom symbols", () => {
    const result = roundTripSdt({
      properties: {
        checkbox: {
          checked: false,
          checkedState: { val: "2714", font: "MS Gothic" },
          uncheckedState: { val: "2715", font: "MS Gothic" },
        },
      },
    });
    expect(result.properties.checkbox?.checked).toBe(false);
    expect(result.properties.checkbox?.checkedState?.val).toBe("2714");
    expect(result.properties.checkbox?.uncheckedState?.val).toBe("2715");
  });

  it("round-trips SDT sdtEndPr run properties", () => {
    const result = roundTripSdt({
      properties: { alias: "EndPr" },
      endProperties: { bold: true, size: 24 },
    });
    expect(result.endProperties).toBeDefined();
    expect(result.endProperties!.bold).toBe(true);
    expect(result.endProperties!.size).toBe(24);
  });
});

// ── customXmlBlockDesc ──

function roundTripCustomXml(opts: CustomXmlBlockDescriptorOptions) {
  const xml = customXmlBlockDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return customXmlBlockDesc.parse(el, readCtx);
}

describe("customXmlBlockDesc round-trip", () => {
  it("round-trips customXml with element name", () => {
    const result = roundTripCustomXml({
      element: "myElement",
    });
    expect(result.element).toBe("myElement");
  });

  it("round-trips customXml with uri", () => {
    const result = roundTripCustomXml({
      element: "item",
      uri: "http://example.com/ns",
    });
    expect(result.element).toBe("item");
    expect(result.uri).toBe("http://example.com/ns");
  });

  it("round-trips customXml with customXmlPr", () => {
    const result = roundTripCustomXml({
      element: "item",
      customXmlPr: {
        placeholder: "Enter text",
        attributes: [
          { name: "attr1", val: "val1" },
          { name: "attr2", val: "val2", uri: "http://example.com" },
        ],
      },
    });
    expect(result.customXmlPr).toBeDefined();
    expect(result.customXmlPr!.placeholder).toBe("Enter text");
    expect(result.customXmlPr!.attributes).toHaveLength(2);
    expect(result.customXmlPr!.attributes![0].name).toBe("attr1");
    expect(result.customXmlPr!.attributes![1].uri).toBe("http://example.com");
  });
});

// ── altChunkDesc / subDocDesc (stringify-only tests) ──

describe("altChunkDesc stringify", () => {
  it("produces w:altChunk element", () => {
    const xml = altChunkDesc.stringify(
      {
        data: "<p>Hello</p>",
        extension: "html",
        contentType: "text/html",
      },
      writeCtx,
    )!;
    expect(xml).toContain("<w:altChunk");
    expect(xml).toContain('r:id="rId');
  });

  it("includes matchSrc when set", () => {
    const xml = altChunkDesc.stringify(
      {
        data: "<p>Hello</p>",
        extension: "html",
        contentType: "text/html",
        matchSource: true,
      },
      writeCtx,
    )!;
    expect(xml).toContain("<w:matchSrc/>");
  });

  it("parses matchSrc back from w:altChunkPr", () => {
    const xml = `<w:altChunk xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:altChunkPr><w:matchSrc/></w:altChunkPr></w:altChunk>`;
    const doc = parseXml(xml);
    const result = altChunkDesc.parse(doc.elements![0], readCtx) as AltChunkOptions;
    expect(result.matchSource).toBe(true);
  });
});

describe("subDocDesc stringify", () => {
  it("produces w:subDoc element", () => {
    const xml = subDocDesc.stringify(
      {
        data: new Uint8Array([1, 2, 3]),
      },
      writeCtx,
    )!;
    expect(xml).toContain("<w:subDoc");
    expect(xml).toContain('r:id="rId');
  });
});
