import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { ReadContext, WriteContext } from "../../descriptor";
import { bodyPropertiesDesc, createBodyProperties } from "./body-properties";
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

  it("emits default shadow/outline from boolean sugar, parses back as full objects", () => {
    const r = roundTrip({ shadow: true, outline: true, rightToLeft: true });
    // shadow: true → default outer shadow; parsed back as EffectListOptions.
    expect((r.shadow as { outerShadow?: unknown }).outerShadow).toBeDefined();
    // outline: true → default solid outline; parsed back as OutlineOptions.
    expect((r.outline as { type?: string }).type).toBe("solidFill");
    expect(r.rightToLeft).toBe(true);
  });

  it("round-trips full OutlineOptions (a:ln)", () => {
    const r = roundTrip({
      outline: { width: 9525, type: "solidFill", color: { value: "FF0000" } },
    });
    const o = r.outline as { width?: number; type?: string; color?: { value: string } };
    expect(o.width).toBe(9525);
    expect(o.type).toBe("solidFill");
    expect(o.color?.value).toBe("FF0000");
  });

  it("round-trips full EffectListOptions (a:effectLst)", () => {
    const r = roundTrip({
      shadow: { outerShadow: { blurRadius: 50000, color: { value: "FF0000" } } },
    });
    const s = r.shadow as { outerShadow?: { blurRadius?: number; color?: { value: string } } };
    expect(s.outerShadow?.blurRadius).toBe(50000);
    expect(s.outerShadow?.color?.value).toBe("FF0000");
  });

  it("round-trips RunFont object (latin/eastAsia/complexScript distinct)", () => {
    const r = roundTrip({
      font: { latin: "Arial", eastAsia: "宋体", complexScript: "Times New Roman" },
    });
    const f = r.font as { latin?: string; eastAsia?: string; complexScript?: string };
    expect(f.latin).toBe("Arial");
    expect(f.eastAsia).toBe("宋体");
    expect(f.complexScript).toBe("Times New Roman");
  });

  it("round-trips TextFont with panose/pitchFamily/charset", () => {
    const r = roundTrip({
      font: {
        latin: { typeface: "Arial", panose: "020B0604020202020204", pitchFamily: 34, charset: 0 },
      },
    });
    const f = r.font as {
      latin?: { typeface: string; panose?: string; pitchFamily?: number; charset?: number };
    };
    expect(f.latin?.typeface).toBe("Arial");
    expect(f.latin?.panose).toBe("020B0604020202020204");
    expect(f.latin?.pitchFamily).toBe(34);
    expect(f.latin?.charset).toBe(0);
  });

  it("round-trips highlight (CT_Color)", () => {
    const r = roundTrip({ highlight: { value: "FFFF00" } });
    expect((r.highlight as { value?: string }).value).toBe("FFFF00");
  });

  it("round-trips underline line/fill follow-text markers", () => {
    const r = roundTrip({ underline: "single", underlineLine: true, underlineFill: true });
    expect(r.underlineLine).toBe(true);
    expect(r.underlineFill).toBe(true);
  });

  it("round-trips explicit uLn/uFill", () => {
    const r = roundTrip({
      underline: "single",
      underlineLine: { width: 12700, type: "solidFill", color: { value: "FF0000" } },
      underlineFill: { type: "solid", color: "00FF00" },
    });
    const line = r.underlineLine as { width?: number; color?: { value?: string } };
    expect(line.width).toBe(12700);
    expect((line.color as { value?: string }).value).toBe("FF0000");
    const fill = r.underlineFill as { type: string; color: { value: string } };
    expect(fill.type).toBe("solid");
    expect(fill.color.value).toBe("00FF00");
  });

  it("round-trips kern/err/smtClean attributes", () => {
    const r = roundTrip({ kern: 12, err: true, smtClean: false });
    expect(r.kern).toBe(12);
    expect(r.err).toBe(true);
    expect(r.smtClean).toBe(false);
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

  it("emits internal slide jump with ppaction token", () => {
    const xml = runPropertiesDesc.stringify({ hyperlink: { slide: 3, tooltip: "Go" } }, writeCtx)!;
    expect(xml).toContain('action="ppaction://hlinksldjump"');
    expect(xml).toContain('tooltip="Go"');
    expect(xml).toContain('r:id="{hlink:');
  });

  it("parses internal slide jump from r:id → slideN.xml + ppaction token", () => {
    const xml = `<a:rPr><a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/></a:rPr>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const slideReadCtx = {
      resolveRelationship: (rId: string) => (rId === "rId7" ? "../slides/slide3.xml" : undefined),
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const r = runPropertiesDesc.parse(el, slideReadCtx);
    expect(r.hyperlink?.slide).toBe(3);
    expect(r.hyperlink?.action).toBeUndefined();
    expect(r.hyperlink?.url).toBeUndefined();
  });

  it("parses external url hyperlink without slide mis-detection", () => {
    const xml = `<a:rPr><a:hlinkClick r:id="rId2"/></a:rPr>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const urlReadCtx = {
      resolveRelationship: (rId: string) => (rId === "rId2" ? "https://example.com" : undefined),
      getPart: () => undefined,
      getRaw: () => undefined,
    } as unknown as ReadContext;
    const r = runPropertiesDesc.parse(el, urlReadCtx);
    expect(r.hyperlink?.url).toBe("https://example.com");
    expect(r.hyperlink?.slide).toBeUndefined();
  });

  it("round-trips action-only hyperlink token without r:id (CT_Hyperlink r:id optional)", () => {
    const xml = runPropertiesDesc.stringify(
      { hyperlink: { action: "ppaction://hlinkshowjump?jump=nextslide" } },
      writeCtx,
    )!;
    // CT_Hyperlink r:id is optional — action-only tokens carry no r:id.
    expect(xml).not.toContain("r:id");
    expect(xml).toContain('action="ppaction://hlinkshowjump?jump=nextslide"');
    // Round-trip: action read back verbatim; no url/slide synthesized.
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("no root");
    const r = runPropertiesDesc.parse(el, readCtx);
    expect(r.hyperlink?.action).toBe("ppaction://hlinkshowjump?jump=nextslide");
    expect(r.hyperlink?.url).toBeUndefined();
    expect(r.hyperlink?.slide).toBeUndefined();
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

  it("stringifies a string child as a text-only run", () => {
    const xml = paragraphDesc.stringify({ children: ["hi", "there"] }, writeCtx)!;
    expect(xml).toContain("<a:t>hi</a:t>");
    expect(xml).toContain("<a:t>there</a:t>");
  });

  it("round-trips mixed string and run children", () => {
    const r = roundTrip({
      children: ["lead ", { text: "run", bold: true }],
    });
    const children = r.children! as RunOptions[];
    expect(children).toHaveLength(2);
    expect(children[0]?.text).toBe("lead ");
    expect(children[1]?.text).toBe("run");
    expect(children[1]?.bold).toBe(true);
  });

  it("round-trips alignment/indentLevel/lineSpacingPercent", () => {
    const r = roundTrip({
      text: "x",
      properties: { alignment: "center", indentLevel: 2, lineSpacingPercent: 150 },
    });
    expect(r.properties?.alignment).toBe("center");
    expect(r.properties?.indentLevel).toBe(2);
    expect(r.properties?.lineSpacingPercent).toBe(150);
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
          sizePoints: 12,
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
    expect(b.sizePoints).toBe(12);
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

  it("round-trips defRPr (default run properties)", () => {
    const r = roundTrip({
      text: "x",
      properties: {
        alignment: "center",
        defaultRunProperties: { size: 24, bold: true, font: "Arial" },
      },
    });
    expect(r.properties?.defaultRunProperties?.size).toBe(24);
    expect(r.properties?.defaultRunProperties?.bold).toBe(true);
    expect(r.properties?.defaultRunProperties?.font).toBe("Arial");
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
    const r = roundTrip({ anchor: "ctr", wrap: "square", numCol: 2 });
    expect(r.anchor).toBe("ctr");
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
    expect((r.paragraphs?.[0] as { text?: string })?.text).toBe("Hi");
  });

  it("emits empty bodyPr/lstStyle and one a:p when bare", () => {
    const inner = textBodyDesc.stringify({}, writeCtx)!;
    expect(inner).toContain("<a:bodyPr");
    expect(inner).toContain("<a:lstStyle");
    expect(inner).toContain("<a:p");
  });

  it("expands text shorthand to a single paragraph", () => {
    const inner = textBodyDesc.stringify({ text: "Hi" }, writeCtx)!;
    expect(inner).toContain("<a:t>Hi</a:t>");
  });

  it("expands string paragraphs entries to one-run paragraphs", () => {
    const inner = textBodyDesc.stringify({ paragraphs: ["A", "B"] }, writeCtx)!;
    expect(inner).toContain("<a:t>A</a:t>");
    expect(inner).toContain("<a:t>B</a:t>");
  });
});

// ── textListStyleDesc ──

describe("textListStyleDesc round-trip", () => {
  it("round-trips DEFAULT_TEXT_LIST_STRUCTURE structure", () => {
    const xml = textListStyleDesc.stringify(DEFAULT_TEXT_LIST_STYLE, writeCtx)!;
    const el = parseXml(`<root>${xml}</root>`).elements?.[0];
    if (!el) throw new Error("no root");
    const r = textListStyleDesc.parse(el, readCtx);
    expect(r.title?.levels?.[0]?.defaultRun?.size).toBe(44);
    expect(r.body?.levels?.[1]?.defaultRun?.size).toBe(24);
    expect(r.other?.emptyDefaultParagraph).toBe(true);
  });
});

// ── textBodyDesc top-level sugar ──

describe("textBodyDesc top-level sugar", () => {
  it("merges anchor/autoFit/columns sugar into bodyProperties on stringify", () => {
    const xml = textBodyDesc.stringify(
      { text: "Hi", anchor: "t", autoFit: "normal", columns: 2 },
      writeCtx,
    )!;
    expect(xml).toContain('anchor="t"');
    expect(xml).toContain("normAutofit");
    expect(xml).toContain('numCol="2"');
  });

  it("maps autoFit shape to spAutoFit", () => {
    const xml = textBodyDesc.stringify({ text: "Hi", autoFit: "shape" }, writeCtx)!;
    expect(xml).toContain("spAutoFit");
  });

  it("emits normalized bodyProperties on round-trip (sugar is input-only)", () => {
    const xml = textBodyDesc.stringify(
      { text: "Hi", anchor: "t", autoFit: "normal", columns: 2 },
      writeCtx,
    )!;
    const el = parseXml(`<root>${xml}</root>`).elements?.[0];
    if (!el) throw new Error("no root");
    const r = textBodyDesc.parse(el, readCtx);
    expect(r.bodyProperties?.anchor).toBe("t");
    expect(r.bodyProperties?.normAutofit).toEqual({});
    expect(r.bodyProperties?.numCol).toBe(2);
    expect(r.anchor).toBeUndefined();
    expect(r.autoFit).toBeUndefined();
  });

  it("explicit bodyProperties wins over sugar", () => {
    const xml = textBodyDesc.stringify({ anchor: "t", bodyProperties: { anchor: "b" } }, writeCtx)!;
    expect(xml).toContain('anchor="b"');
    expect(xml).not.toContain('anchor="t"');
  });
});
