import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ReadContext, WriteContext } from "../../descriptor";
import { bodyPropertiesDesc, createBodyProperties, VerticalAnchor } from "./body-properties";
import { textListStyleDesc, DEFAULT_TEXT_LIST_STYLE } from "./list-style";
import { paragraphDesc } from "./paragraph";
import type { ParagraphDescriptorOptions } from "./paragraph";
import { textRunDesc } from "./run";
import { runPropertiesDesc } from "./run-properties";
import { textBodyDesc } from "./text-body";
import type { RunOptions, RunPropertiesOptions } from "./types";

// ── Mock write context ──

class MockWriteContext {
  addRelationship() {
    return "rId1";
  }
  addMedia() {
    return "";
  }
  addHyperlink() {}
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
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    return runPropertiesDesc.parse(el, readCtx);
  }

  it("round-trips size/bold/italic/underline", () => {
    const r = roundTrip({ size: 24, bold: true, italic: true, underline: "single" });
    expect(r.size).toBe(24);
    expect(r.bold).toBe(true);
    expect(r.italic).toBe(true);
    expect(r.underline).toBe("single");
  });

  it("round-trips font/lang/spacing/baseline", () => {
    const r = roundTrip({ font: "Arial", lang: "zh-CN", spacing: 200, baseline: 30000 });
    expect(r.font).toBe("Arial");
    expect(r.lang).toBe("zh-CN");
    expect(r.spacing).toBe(200);
    expect(r.baseline).toBe(30000);
  });

  it("round-trips strike/capitalization", () => {
    const r = roundTrip({ strike: "sngStrike", capitalization: "all" });
    expect(r.strike).toBe("singleStrike");
    expect(r.capitalization).toBe("all");
  });

  it("round-trips shadow/outline/rightToLeft", () => {
    const r = roundTrip({ shadow: true, outline: true, rightToLeft: true });
    expect(r.shadow).toBe(true);
    expect(r.outline).toBe(true);
    expect(r.rightToLeft).toBe(true);
  });

  it("round-trips solid fill", () => {
    const r = roundTrip({ fill: { type: "solid", color: "FF0000" } });
    const fill = r.fill! as { type: string; color: { value: string } };
    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("FF0000");
  });

  it("round-trips hover hyperlink (a:hlinkMouseOver)", () => {
    const r = roundTrip({
      mouseoverHyperlink: { url: "https://example.org", tooltip: "Hover tip" },
    });
    expect(r.mouseoverHyperlink).toBeDefined();
    expect(r.mouseoverHyperlink?.tooltip).toBe("Hover tip");
  });
});

// ── textRunDesc ──

describe("textRunDesc round-trip", () => {
  function roundTrip(opts: RunOptions) {
    const xml = textRunDesc.stringify(opts, writeCtx)!;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    return textRunDesc.parse(el, readCtx);
  }

  it("round-trips plain text", () => {
    expect(roundTrip({ text: "Hello" }).text).toBe("Hello");
  });

  it("round-trips text with formatting", () => {
    const r = roundTrip({ text: "Bold", bold: true, size: 24 });
    expect(r.text).toBe("Bold");
    expect(r.bold).toBe(true);
    expect(r.size).toBe(24);
  });
});

// ── paragraphDesc ──

describe("paragraphDesc round-trip", () => {
  function roundTrip(opts: ParagraphDescriptorOptions) {
    const xml = paragraphDesc.stringify(opts, writeCtx)!;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    return paragraphDesc.parse(el, readCtx);
  }

  it("round-trips simple text", () => {
    expect(roundTrip({ text: "Hello paragraph" }).text).toBe("Hello paragraph");
  });

  it("round-trips children runs", () => {
    const r = roundTrip({
      children: [
        { text: "Hello ", bold: true },
        { text: "World", italic: true },
      ],
    });
    const children = r.children! as RunOptions[];
    expect(children).toHaveLength(2);
    expect(children[0]?.text).toBe("Hello ");
    expect(children[0]?.bold).toBe(true);
    expect(children[1]?.text).toBe("World");
    expect(children[1]?.italic).toBe(true);
  });

  it("round-trips alignment/indentLevel/lineSpacing", () => {
    const r = roundTrip({
      text: "x",
      properties: { alignment: "center", indentLevel: 2, lineSpacing: 150 },
    });
    expect(r.properties?.alignment).toBe("center");
    expect(r.properties?.indentLevel).toBe(2);
    expect(r.properties?.lineSpacing).toBe(150);
  });

  it("round-trips bullet none/char/autoNum", () => {
    const none = roundTrip({ text: "a", properties: { bullet: { type: "none" } } });
    expect(none.properties?.bullet?.type).toBe("none");
    const ch = roundTrip({
      text: "b",
      properties: { bullet: { type: "char", char: "•", color: "FF0000" } },
    });
    const chBullet = ch.properties?.bullet as { type: string; char?: string; color?: string };
    expect(chBullet.type).toBe("char");
    expect(chBullet.char).toBe("•");
    expect(chBullet.color).toBe("FF0000");
    const an = roundTrip({
      text: "c",
      properties: { bullet: { type: "autoNum", format: "arabicPeriod", startAt: 1 } },
    });
    const anBullet = an.properties?.bullet as {
      type: string;
      format?: string;
      startAt?: number;
    };
    expect(anBullet.type).toBe("autoNum");
    expect(anBullet.format).toBe("arabicPeriod");
    expect(anBullet.startAt).toBe(1);
  });

  it("round-trips bullet color/size/font follows-text and points variants", () => {
    const r = roundTrip({
      text: "x",
      properties: {
        bullet: {
          type: "char",
          char: "•",
          colorFollowsText: true,
          sizePoints: 1200,
          fontFollowsText: true,
        },
      },
    });
    const b = r.properties?.bullet as {
      type: string;
      colorFollowsText?: boolean;
      sizePoints?: number;
      fontFollowsText?: boolean;
    };
    expect(b.type).toBe("char");
    expect(b.colorFollowsText).toBe(true);
    expect(b.sizePoints).toBe(1200);
    expect(b.fontFollowsText).toBe(true);
  });

  it("round-trips a picture bullet (a:buBlip)", () => {
    const r = roundTrip({
      text: "y",
      properties: { bullet: { type: "picture", embed: "rId2" } },
    });
    const b = r.properties?.bullet as { type: string; embed?: string };
    expect(b.type).toBe("picture");
    expect(b.embed).toBe("rId2");
  });

  it("round-trips tab stops (a:tabLst)", () => {
    const r = roundTrip({
      text: "x",
      properties: {
        tabStops: [
          { position: 914400, alignment: "l" },
          { position: 4572000, alignment: "dec" },
        ],
      },
    });
    expect(r.properties?.tabStops).toEqual([
      { position: 914400, alignment: "l" },
      { position: 4572000, alignment: "dec" },
    ]);
  });

  it("preserves single-run formatting (no text-shorthand collapse)", () => {
    const r = roundTrip({ children: [{ text: "Bold", bold: true }] });
    const children = r.children! as RunOptions[];
    expect(children).toHaveLength(1);
    expect(children[0]?.text).toBe("Bold");
    expect(children[0]?.bold).toBe(true);
  });
});

// ── bodyPropertiesDesc ──

describe("bodyPropertiesDesc round-trip", () => {
  function roundTrip(opts: Parameters<typeof bodyPropertiesDesc.stringify>[0]) {
    const xml = bodyPropertiesDesc.stringify(opts, writeCtx)!;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    return bodyPropertiesDesc.parse(el, readCtx);
  }

  it("createBodyProperties defaults to a:bodyPr; DOCX passes wps:bodyPr", () => {
    expect(createBodyProperties({ wrap: "square" })).toContain("<a:bodyPr");
    expect(createBodyProperties({ wrap: "square" }, "wps:bodyPr")).toContain("<wps:bodyPr");
  });

  it("round-trips anchor/wrap/numCol", () => {
    const r = roundTrip({ anchor: VerticalAnchor.CENTER, wrap: "square", numCol: 2 });
    expect(r.anchor).toBe(VerticalAnchor.CENTER);
    expect(r.wrap).toBe("square");
    expect(r.numCol).toBe(2);
  });

  it("round-trips margins shorthand", () => {
    const r = roundTrip({ margins: { top: 100, bottom: 200, left: 300, right: 400 } });
    expect(r.tIns).toBe(100);
    expect(r.bIns).toBe(200);
    expect(r.lIns).toBe(300);
    expect(r.rIns).toBe(400);
  });

  it("round-trips autofit variants", () => {
    const norm = roundTrip({ normAutofit: { fontScale: 80000 } });
    expect(norm.normAutofit?.fontScale).toBe(80000);
    const sp = roundTrip({ spAutoFit: true });
    expect(sp.spAutoFit).toBe(true);
    const none = roundTrip({ noAutoFit: true });
    expect(none.noAutoFit).toBe(true);
  });
});

// ── textBodyDesc ──

describe("textBodyDesc round-trip", () => {
  it("round-trips bodyProperties + paragraphs", () => {
    const inner = textBodyDesc.stringify(
      { bodyProperties: { wrap: "square" }, paragraphs: [{ text: "Hi" }] },
      writeCtx,
    )!;
    // Wrap in a container tag (p:txBody / xdr:txBody at the call site).
    const el = parseXml(`<a:txBody>${inner}</a:txBody>`).elements?.[0];
    if (!el) throw new Error("no root");
    const r = textBodyDesc.parse(el, readCtx);
    expect(r.bodyProperties?.wrap).toBe("square");
    expect(r.paragraphs?.[0]?.text).toBe("Hi");
  });

  it("emits empty bodyPr/lstStyle and one a:p when bare", () => {
    const inner = textBodyDesc.stringify({}, writeCtx)!;
    expect(inner).toContain("<a:bodyPr");
    expect(inner).toContain("<a:lstStyle");
    expect(inner).toContain("<a:p");
  });
});

// ── textListStyleDesc ──

describe("textListStyleDesc round-trip", () => {
  it("round-trips DEFAULT_TEXT_LIST_STRUCTURE structure", () => {
    const xml = textListStyleDesc.stringify(DEFAULT_TEXT_LIST_STYLE, writeCtx)!;
    const el = parseXml(`<root>${xml}</root>`).elements?.[0];
    if (!el) throw new Error("no root");
    const r = textListStyleDesc.parse(el, readCtx);
    expect(r.title?.levels?.[0]?.defaultRun?.size).toBe(4400);
    expect(r.body?.levels?.[1]?.defaultRun?.size).toBe(2400);
    expect(r.other?.emptyDefaultParagraph).toBe(true);
  });
});
