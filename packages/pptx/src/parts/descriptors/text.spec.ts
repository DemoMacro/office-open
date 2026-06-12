import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { RunOptions } from "@shared/shape/paragraph/run";
import type { RunPropertiesOptions } from "@shared/shape/paragraph/run-properties";
import { describe, expect, it } from "vite-plus/test";

import { runPropertiesDesc, textRunDesc, paragraphDesc } from "./text";
import type { ParagraphDescriptorOptions } from "./text";

// ── Mock PPTX write context ──

class MockWriteContext {
  private _nextRelId = 1;
  addRelationship() {
    return `rId${this._nextRelId++}`;
  }
  addMedia() {
    return "";
  }
  addHyperlink() {}
  addImage() {}
}
const writeCtx = new MockWriteContext() as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

// ── runPropertiesDesc ──

describe("runPropertiesDesc round-trip", () => {
  function roundTrip(opts: RunPropertiesOptions) {
    const xml = runPropertiesDesc.stringify(opts, writeCtx)!;
    const doc = parseXml(xml);
    const el = doc.elements![0];
    return runPropertiesDesc.parse(el, readCtx);
  }

  it("round-trips size", () => {
    const result = roundTrip({ size: 24 });
    expect(result.size).toBe(24);
  });

  it("round-trips bold", () => {
    const result = roundTrip({ bold: true });
    expect(result.bold).toBe(true);
  });

  it("round-trips italic", () => {
    const result = roundTrip({ italic: true });
    expect(result.italic).toBe(true);
  });

  it("round-trips underline", () => {
    const result = roundTrip({ underline: "single" });
    expect(result.underline).toBe("single");
  });

  it("round-trips font", () => {
    const result = roundTrip({ font: "Arial" });
    expect(result.font).toBe("Arial");
  });

  it("round-trips lang", () => {
    const result = roundTrip({ lang: "zh-CN" });
    expect(result.lang).toBe("zh-CN");
  });

  it("round-trips strike", () => {
    const result = roundTrip({ strike: "sngStrike" });
    expect(result.strike).toBe("singleStrike");
  });

  it("round-trips baseline", () => {
    const result = roundTrip({ baseline: 30000 });
    expect(result.baseline).toBe(30000);
  });

  it("round-trips spacing", () => {
    const result = roundTrip({ spacing: 200 });
    expect(result.spacing).toBe(200);
  });

  it("round-trips capitalization", () => {
    const result = roundTrip({ capitalization: "all" });
    expect(result.capitalization).toBe("all");
  });

  it("round-trips shadow", () => {
    const result = roundTrip({ shadow: true });
    expect(result.shadow).toBe(true);
  });

  it("round-trips outline", () => {
    const result = roundTrip({ outline: true });
    expect(result.outline).toBe(true);
  });

  it("round-trips rightToLeft", () => {
    const result = roundTrip({ rightToLeft: true });
    expect(result.rightToLeft).toBe(true);
  });

  it("round-trips noProof", () => {
    const result = roundTrip({ noProof: true });
    expect(result.noProof).toBe(true);
  });

  it("round-trips dirty", () => {
    const result = roundTrip({ dirty: true });
    expect(result.dirty).toBe(true);
  });

  it("round-trips kumimoji", () => {
    const result = roundTrip({ kumimoji: true });
    expect(result.kumimoji).toBe(true);
  });

  it("round-trips alternateLanguage", () => {
    const result = roundTrip({ alternateLanguage: "en-US" });
    expect(result.alternateLanguage).toBe("en-US");
  });

  it("round-trips normalizeHeight", () => {
    const result = roundTrip({ normalizeHeight: true });
    expect(result.normalizeHeight).toBe(true);
  });

  it("round-trips solid fill", () => {
    const result = roundTrip({ fill: { type: "solid", color: "FF0000" } });
    const fill = result.fill! as { type: string; color: { value: string } };
    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("FF0000");
  });

  it("round-trips combined properties", () => {
    const result = roundTrip({
      size: 18,
      bold: true,
      italic: true,
      underline: "double",
      font: "Calibri",
      lang: "en-US",
    });
    expect(result.size).toBe(18);
    expect(result.bold).toBe(true);
    expect(result.italic).toBe(true);
    expect(result.underline).toBe("double");
    expect(result.font).toBe("Calibri");
    expect(result.lang).toBe("en-US");
  });
});

// ── textRunDesc ──

describe("textRunDesc round-trip", () => {
  function roundTrip(opts: RunOptions) {
    const xml = textRunDesc.stringify(opts, writeCtx)!;
    const doc = parseXml(xml);
    const el = doc.elements![0];
    return textRunDesc.parse(el, readCtx);
  }

  it("round-trips plain text", () => {
    const result = roundTrip({ text: "Hello" });
    expect(result.text).toBe("Hello");
  });

  it("round-trips text with formatting", () => {
    const result = roundTrip({ text: "Bold", bold: true, size: 24 });
    expect(result.text).toBe("Bold");
    expect(result.bold).toBe(true);
    expect(result.size).toBe(24);
  });

  it("round-trips text with fill", () => {
    const result = roundTrip({ text: "Red", fill: { type: "solid", color: "FF0000" } });
    expect(result.text).toBe("Red");
    const fill = result.fill! as { type: string; color: { value: string } };
    expect(fill.color.value).toBe("FF0000");
  });
});

// ── paragraphDesc ──

describe("paragraphDesc round-trip", () => {
  function roundTrip(opts: ParagraphDescriptorOptions) {
    const xml = paragraphDesc.stringify(opts, writeCtx)!;
    const doc = parseXml(xml);
    const el = doc.elements![0];
    return paragraphDesc.parse(el, readCtx);
  }

  it("round-trips simple text", () => {
    const result = roundTrip({ text: "Hello paragraph" });
    expect(result.text).toBe("Hello paragraph");
  });

  it("round-trips children runs", () => {
    const result = roundTrip({
      children: [
        { text: "Hello ", bold: true },
        { text: "World", italic: true },
      ],
    });
    const children = result.children!;
    expect(children).toHaveLength(2);
    expect(children[0].text).toBe("Hello ");
    expect(children[0].bold).toBe(true);
    expect(children[1].text).toBe("World");
    expect(children[1].italic).toBe(true);
  });

  it("round-trips alignment", () => {
    const result = roundTrip({
      text: "Centered",
      properties: { alignment: "center" },
    });
    const props = result.properties!;
    expect(props.alignment).toBe("center");
  });

  it("round-trips indentLevel", () => {
    const result = roundTrip({
      text: "Indented",
      properties: { indentLevel: 2 },
    });
    const props = result.properties!;
    expect(props.indentLevel).toBe(2);
  });

  it("round-trips lineSpacing", () => {
    const result = roundTrip({
      text: "Spaced",
      properties: { lineSpacing: 150 },
    });
    const props = result.properties!;
    expect(props.lineSpacing).toBe(150);
  });

  it("round-trips marginIndent", () => {
    const result = roundTrip({
      text: "Indented",
      properties: { marginIndent: 457200 },
    });
    const props = result.properties!;
    expect(props.marginIndent).toBe(457200);
  });

  it("round-trips bullet none", () => {
    const result = roundTrip({
      text: "No bullet",
      properties: { bullet: { type: "none" } },
    });
    const props = result.properties!;
    const bullet = props.bullet!;
    expect(bullet.type).toBe("none");
  });

  it("round-trips bullet char", () => {
    const result = roundTrip({
      text: "Bulleted",
      properties: { bullet: { type: "char", char: "•", color: "FF0000" } },
    });
    const props = result.properties!;
    const bullet = props.bullet! as { type: string; char?: string; color?: string };
    expect(bullet.type).toBe("char");
    expect(bullet.char).toBe("•");
    expect(bullet.color).toBe("FF0000");
  });

  it("round-trips bullet autoNum", () => {
    const result = roundTrip({
      text: "Numbered",
      properties: { bullet: { type: "autoNum", format: "arabicPeriod", startAt: 1 } },
    });
    const props = result.properties!;
    const bullet = props.bullet! as { type: string; format?: string; startAt?: number };
    expect(bullet.type).toBe("autoNum");
    expect(bullet.format).toBe("arabicPeriod");
    expect(bullet.startAt).toBe(1);
  });

  it("round-trips marginBottom and marginTop", () => {
    const result = roundTrip({
      text: "With margins",
      properties: { marginTop: 400, marginBottom: 200 },
    });
    const props = result.properties!;
    expect(props.marginTop).toBe(400);
    expect(props.marginBottom).toBe(200);
  });

  it("round-trips properties with children", () => {
    const result = roundTrip({
      properties: { alignment: "right" },
      children: [{ text: "Right-aligned run" }],
    });
    const props = result.properties!;
    expect(props.alignment).toBe("right");
    const children = result.children!;
    expect(children[0].text).toBe("Right-aligned run");
  });
});
