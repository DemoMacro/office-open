import type { ParagraphDescriptorOptions } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { fromDrawingParagraph, toDrawingParagraph } from "./text";

describe("fromDrawingParagraph (a:p → w:p)", () => {
  it("maps text shorthand and alignment", () => {
    const drawing: ParagraphDescriptorOptions = {
      text: "Hello",
      properties: { alignment: "center" },
    };
    const docx = fromDrawingParagraph(drawing);
    expect(docx.text).toBe("Hello");
    expect(docx.alignment).toBe("center");
  });

  it("passes font size (points) and direct run flags straight through", () => {
    const drawing: ParagraphDescriptorOptions = {
      children: [{ text: "hi", bold: true, italic: true, size: 24, font: "Arial" }],
    };
    const run = fromDrawingParagraph(drawing).children![0] as Record<string, unknown>;
    expect(run.text).toBe("hi");
    expect(run.bold).toBe(true);
    expect(run.italic).toBe(true);
    expect(run.size).toBe(24);
    expect(run.font).toBe("Arial");
  });

  it("maps solid fill → color hex and drops non-solid fills", () => {
    const withSolid: ParagraphDescriptorOptions = {
      children: [{ text: "a", fill: { type: "solid", color: "FF0000" } }],
    };
    expect((fromDrawingParagraph(withSolid).children![0] as { color?: string }).color).toBe(
      "FF0000",
    );

    const withString: ParagraphDescriptorOptions = { children: [{ text: "a", fill: "00FF00" }] };
    expect((fromDrawingParagraph(withString).children![0] as { color?: string }).color).toBe(
      "00FF00",
    );

    const withGradient: ParagraphDescriptorOptions = {
      children: [
        { text: "a", fill: { type: "gradient", stops: [{ color: "000000", position: 0 }] } },
      ],
    };
    expect(
      (fromDrawingParagraph(withGradient).children![0] as { color?: string }).color,
    ).toBeUndefined();
  });

  it("converts underline and strike enums", () => {
    const drawing: ParagraphDescriptorOptions = {
      children: [
        { text: "u", underline: "double", strike: "doubleStrike" },
        { text: "v", underline: "single", strike: "singleStrike" },
      ],
    };
    const [a, b] = fromDrawingParagraph(drawing).children as [
      { underline?: { type?: string }; doubleStrike?: boolean; strike?: boolean },
      { underline?: { type?: string }; strike?: boolean },
    ];
    expect(a.underline?.type).toBe("double");
    expect(a.doubleStrike).toBe(true);
    expect(b.underline?.type).toBe("single");
    expect(b.strike).toBe(true);
  });

  it("maps baseline sign to subscript/superscript flags", () => {
    const drawing: ParagraphDescriptorOptions = {
      children: [
        { text: "up", baseline: 30000 },
        { text: "down", baseline: -25000 },
      ],
    };
    const [up, down] = fromDrawingParagraph(drawing).children as [
      { superScript?: boolean; subScript?: boolean },
      { superScript?: boolean; subScript?: boolean },
    ];
    expect(up.superScript).toBe(true);
    expect(down.subScript).toBe(true);
  });

  it("converts run spacing (1/100 pt → twips) and capitalization", () => {
    const drawing: ParagraphDescriptorOptions = {
      children: [{ text: "a", spacing: 100, capitalization: "all" }],
    };
    const run = fromDrawingParagraph(drawing).children![0] as {
      characterSpacing?: number;
      allCaps?: boolean;
    };
    expect(run.characterSpacing).toBe(20); // 1 pt = 20 twips
    expect(run.allCaps).toBe(true);
  });

  it("converts paragraph before/after (1/100 pt → twips)", () => {
    const drawing: ParagraphDescriptorOptions = {
      properties: { spaceBefore: 500, spaceAfter: 250 },
    };
    const spacing = fromDrawingParagraph(drawing).spacing!;
    expect(spacing.before).toBe(100); // 5 pt
    expect(spacing.after).toBe(50); // 2.5 pt
  });

  it("converts percent line spacing to auto, points line spacing to exact", () => {
    const pct = fromDrawingParagraph({ properties: { lineSpacingPercent: 150 } }).spacing!;
    expect(pct.line).toBe(360); // 1.5 × 240
    expect(pct.lineRule).toBe("auto");

    const pts = fromDrawingParagraph({ properties: { lineSpacingPoints: 24 } }).spacing!;
    expect(pts.line).toBe(480); // 24 pt × 20
    expect(pts.lineRule).toBe("exact");
  });

  it("converts indents (EMU → twips)", () => {
    const drawing: ParagraphDescriptorOptions = {
      properties: { marginIndent: 914400, marginRight: 457200 },
    };
    const indent = fromDrawingParagraph(drawing).indent!;
    expect(indent.start).toBe(1440); // 1 inch
    expect(indent.end).toBe(720); // 0.5 inch
  });

  it("emits a docx bullet only when the source is actually bulleted", () => {
    const bulleted = fromDrawingParagraph({
      properties: { bullet: { type: "char", char: "•" }, indentLevel: 1 },
    });
    expect(bulleted.bullet).toEqual({ level: 1 });

    const bareLevel = fromDrawingParagraph({ properties: { indentLevel: 2 } });
    expect(bareLevel.bullet).toBeUndefined();
  });

  it("wraps a hyperlinked run in a w:hyperlink child", () => {
    const drawing: ParagraphDescriptorOptions = {
      children: [
        { text: "link", bold: true, hyperlink: { url: "https://example.com", tooltip: "go" } },
      ],
    };
    const child = fromDrawingParagraph(drawing).children![0] as {
      hyperlink?: {
        url?: string;
        tooltip?: string;
        children?: { text?: string; bold?: boolean }[];
      };
    };
    expect(child.hyperlink?.url).toBe("https://example.com");
    expect(child.hyperlink?.tooltip).toBe("go");
    const inner = child.hyperlink?.children?.[0];
    expect(inner?.text).toBe("link");
    expect(inner?.bold).toBe(true);
  });

  it("converts a soft break to a w:br run", () => {
    const drawing: ParagraphDescriptorOptions = {
      children: [{ text: "a" }, { break: true }],
    };
    const children = fromDrawingParagraph(drawing).children as [
      { text: string },
      { break: number },
    ];
    expect(children[1].break).toBe(1);
  });
});

describe("toDrawingParagraph (w:p → a:p)", () => {
  it("maps text shorthand and alignment back, normalizing start/end", () => {
    const drawing = toDrawingParagraph({ text: "Hi", alignment: "start" });
    expect(drawing.text).toBe("Hi");
    expect(drawing.properties?.alignment).toBe("left");
  });

  it("collapses docx underline styles to single/double and maps strike back", () => {
    const drawing = toDrawingParagraph({
      children: [
        { text: "a", underline: { type: "dotted" }, doubleStrike: true },
        { text: "b", underline: { type: "double" }, strike: true },
      ],
    });
    const [a, b] = drawing.children as [
      { underline?: string; strike?: string },
      { underline?: string; strike?: string },
    ];
    expect(a.underline).toBe("single"); // dotted collapsed
    expect(a.strike).toBe("doubleStrike");
    expect(b.underline).toBe("double");
    expect(b.strike).toBe("singleStrike");
  });

  it("recovers run spacing (twips → 1/100 pt) and color", () => {
    const drawing = toDrawingParagraph({
      children: [{ text: "a", characterSpacing: 20, color: "FF0000" }],
    });
    const run = drawing.children![0] as { spacing?: number; fill?: string };
    expect(run.spacing).toBe(100); // 20 twips → 100 hundredths
    expect(run.fill).toBe("FF0000");
  });

  it("recovers super/subscript as signed baseline", () => {
    const up = toDrawingParagraph({ children: [{ text: "a", superScript: true }] });
    expect((up.children![0] as { baseline?: number }).baseline).toBe(30000);
    const down = toDrawingParagraph({ children: [{ text: "a", subScript: true }] });
    expect((down.children![0] as { baseline?: number }).baseline).toBe(-25000);
  });

  it("converts auto/exact line spacing back to percent/points", () => {
    const auto = toDrawingParagraph({ spacing: { line: 360, lineRule: "auto" } });
    expect(auto.properties?.lineSpacingPercent).toBe(150);
    const exact = toDrawingParagraph({ spacing: { line: 480, lineRule: "exact" } });
    expect(exact.properties?.lineSpacingPoints).toBe(24);
  });

  it("flattens a w:hyperlink child into a run carrying the hyperlink", () => {
    const drawing = toDrawingParagraph({
      children: [
        {
          hyperlink: { url: "https://example.com", tooltip: "go", children: [{ text: "link" }] },
        },
      ],
    });
    const run = drawing.children![0] as {
      text?: string;
      hyperlink?: { url?: string; tooltip?: string };
    };
    expect(run.text).toBe("link");
    expect(run.hyperlink?.url).toBe("https://example.com");
    expect(run.hyperlink?.tooltip).toBe("go");
  });

  it("maps a docx bullet to a:p indentLevel + a default char bullet", () => {
    const drawing = toDrawingParagraph({ bullet: { level: 2 } });
    expect(drawing.properties?.indentLevel).toBe(2);
    expect(drawing.properties?.bullet).toEqual({ type: "char", char: "•" });
  });

  it("drops docx-only paragraph children lossily", () => {
    const drawing = toDrawingParagraph({
      children: [{ text: "keep" }, { commentReference: 1 }],
    });
    expect(drawing.children?.length).toBe(1);
    expect((drawing.children![0] as { text?: string }).text).toBe("keep");
  });
});

describe("round-trip a:p → w:p → a:p", () => {
  it("preserves core high-fidelity run fields", () => {
    const original: ParagraphDescriptorOptions = {
      children: [
        { text: "x", bold: true, size: 18, underline: "single", fill: "ABCDEF", lang: "en-US" },
      ],
      properties: { alignment: "right", spaceBefore: 200 },
    };
    const back = toDrawingParagraph(fromDrawingParagraph(original));
    const run = back.children![0] as {
      text?: string;
      bold?: boolean;
      size?: number;
      underline?: string;
      fill?: string;
      lang?: string;
    };
    expect(run).toMatchObject({
      text: "x",
      bold: true,
      size: 18,
      underline: "single",
      fill: "ABCDEF",
      lang: "en-US",
    });
    expect(back.properties?.alignment).toBe("right");
    expect(back.properties?.spaceBefore).toBe(200); // 200 hundredths → 40 twips → 200 hundredths
  });
});
