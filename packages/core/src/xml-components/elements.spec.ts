import { describe, expect, it } from "vite-plus/test";

import type { Context } from "./base";
import {
  BuilderElement,
  EmptyElement,
  attrObj,
  chartAttr,
  hpsMeasureObj,
  numberValObj,
  onOffObj,
  stringContainerObj,
  stringEnumValObj,
  stringValObj,
  wrapEl,
} from "./elements";

const emptyContext: Context = { stack: [] };

describe("onOffObj (CT_OnOff)", () => {
  it("should emit no val attribute when true (default)", () => {
    expect(onOffObj("w:b")).toEqual({ "w:b": {} });
  });

  it("should emit no val attribute when explicitly true", () => {
    expect(onOffObj("w:b", true)).toEqual({ "w:b": {} });
  });

  it("should return cached singleton for same name with true", () => {
    expect(onOffObj("w:b", true)).toBe(onOffObj("w:b", true));
  });

  it("should emit w:val=false for w: element", () => {
    expect(onOffObj("w:b", false)).toEqual({ "w:b": { _attr: { "w:val": false } } });
  });

  it("should emit m:val=false for m: element (dynamic namespace)", () => {
    expect(onOffObj("m:hideBot", false)).toEqual({ "m:hideBot": { _attr: { "m:val": false } } });
  });

  it("should emit a:val=false for a: element", () => {
    expect(onOffObj("a:noAutofit", false)).toEqual({
      "a:noAutofit": { _attr: { "a:val": false } },
    });
  });

  it("should handle element without namespace prefix", () => {
    expect(onOffObj("bold", false)).toEqual({ bold: { _attr: { "bold:val": false } } });
  });
});

describe("hpsMeasureObj (CT_HpsMeasure)", () => {
  it("should accept a number", () => {
    expect(hpsMeasureObj("w:sz", 24)).toEqual({ "w:sz": { _attr: { "w:val": 24 } } });
  });

  it("should accept a universal measure string", () => {
    expect(hpsMeasureObj("w:sz", "12pt")).toEqual({ "w:sz": { _attr: { "w:val": "12pt" } } });
  });
});

describe("stringValObj", () => {
  it("should emit val attribute with w: namespace", () => {
    expect(stringValObj("w:pStyle", "Heading1")).toEqual({
      "w:pStyle": { _attr: { "w:val": "Heading1" } },
    });
  });

  it("should emit val attribute with m: namespace", () => {
    expect(stringValObj("m:mathFont", "Cambria Math")).toEqual({
      "m:mathFont": { _attr: { "m:val": "Cambria Math" } },
    });
  });
});

describe("numberValObj", () => {
  it("should emit val attribute with number", () => {
    expect(numberValObj("w:ilvl", 0)).toEqual({
      "w:ilvl": { _attr: { "w:val": 0 } },
    });
  });
});

describe("stringEnumValObj", () => {
  it("should emit val attribute with enum string", () => {
    expect(stringEnumValObj("w:jc", "center")).toEqual({
      "w:jc": { _attr: { "w:val": "center" } },
    });
  });
});

describe("stringContainerObj", () => {
  it("should contain text content", () => {
    expect(stringContainerObj("w:t", "Hello")).toEqual({ "w:t": ["Hello"] });
  });
});

describe("attrObj", () => {
  it("should filter out undefined values", () => {
    expect(attrObj("w:spacing", { "w:before": 100, "w:after": undefined })).toEqual({
      "w:spacing": { _attr: { "w:before": 100 } },
    });
  });

  it("should include all defined values", () => {
    expect(attrObj("w:spacing", { "w:before": 100, "w:after": 200 })).toEqual({
      "w:spacing": { _attr: { "w:before": 100, "w:after": 200 } },
    });
  });
});

describe("EmptyElement (CT_Empty)", () => {
  it("should produce an empty element", () => {
    const el = new EmptyElement("w:bookmarkStart");
    expect(el.prepForXml(emptyContext)).toEqual({ "w:bookmarkStart": {} });
  });
});

describe("BuilderElement", () => {
  it("should create an element with no attributes or children", () => {
    const el = new BuilderElement({ name: "w:p" });
    expect(el.prepForXml(emptyContext)).toEqual({ "w:p": {} });
  });

  it("should create an element with attributes (single attr unwrapped)", () => {
    const el = new BuilderElement({
      name: "w:pPr",
      attributes: {
        style: { key: "w:val", value: "Test" },
      },
    });
    expect(el.prepForXml(emptyContext)).toEqual({
      "w:pPr": { _attr: { "w:val": "Test" } },
    });
  });

  it("should create an element with children", () => {
    const child = new EmptyElement("w:r");
    const el = new BuilderElement({ name: "w:p", children: [child] });
    expect(el.prepForXml(emptyContext)).toEqual({
      "w:p": [{ "w:r": {} }],
    });
  });

  it("should create an element with both attributes and children", () => {
    const child = stringContainerObj("w:t", "text");
    const el = new BuilderElement({
      name: "w:r",
      attributes: { lang: { key: "xml:lang", value: "en-US" } },
      children: [child],
    });
    expect(el.prepForXml(emptyContext)).toEqual({
      "w:r": [{ _attr: { "xml:lang": "en-US" } }, { "w:t": ["text"] }],
    });
  });

  it("should accept IXmlableObject children directly", () => {
    const el = new BuilderElement({
      name: "w:pPr",
      children: [onOffObj("w:b", true), stringValObj("w:pStyle", "Heading1")],
    });
    expect(el.prepForXml(emptyContext)).toEqual({
      "w:pPr": [{ "w:b": {} }, { "w:pStyle": { _attr: { "w:val": "Heading1" } } }],
    });
  });
});

describe("chartAttr", () => {
  it("should create attribute object with explicit keys", () => {
    const attr = chartAttr({ "r:id": "rId1", "c:val": 42 });
    expect(attr).toEqual({
      _attr: { "r:id": "rId1", "c:val": 42 },
    });
  });
});

describe("wrapEl", () => {
  it("should wrap a component in a named element", () => {
    const inner = new EmptyElement("w:r");
    const wrapped = wrapEl("w:r", inner);
    expect(wrapped.prepForXml(emptyContext)).toEqual({
      "w:r": [{ "w:r": {} }],
    });
  });
});
