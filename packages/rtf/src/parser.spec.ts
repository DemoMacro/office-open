import type {
  ParagraphOptions,
  RunOptions,
  SectionChild,
  TableCellOptions,
  TableOptions,
} from "@office-open/docx";
import { describe, expect, it } from "vite-plus/test";

import { parseRtf } from ".";

function firstParagraph(source: string): ParagraphOptions {
  const children = parseRtf(source).sections[0]?.children ?? [];
  const child = children[0];
  if (!child || !("paragraph" in child) || typeof child.paragraph === "string") {
    throw new Error("expected the first child to be a paragraph");
  }
  return child.paragraph;
}

function firstTable(source: string): TableOptions {
  const children = parseRtf(source).sections[0]?.children ?? [];
  const child = children[0];
  if (!child || !("table" in child)) throw new Error("expected the first child to be a table");
  return child.table;
}

function runs(paragraph: ParagraphOptions): RunOptions[] {
  return paragraph.children as RunOptions[];
}

describe("parseRtf text and formatting", () => {
  it("parses standard text into a document paragraph", () => {
    expect(parseRtf("{\\rtf1 Hello, RTF!}")).toEqual({
      sections: [{ children: [{ paragraph: { children: [{ text: "Hello, RTF!" }] } }] }],
    });
  });

  it("scopes bold formatting to its group and maps paragraph alignment", () => {
    const paragraph = firstParagraph(String.raw`{\rtf1{\b Bold} plain\par}{\qc Centered\par}`);
    expect(runs(paragraph)).toEqual([{ bold: true, text: "Bold" }, { text: " plain" }]);
  });

  it("parses tabs, hexadecimal, Unicode, and common character controls", () => {
    const children = parseRtf(String.raw`{\rtf1\tab A\'42\u9786?}`).sections[0]?.children ?? [];
    const paragraph = children[0] && "paragraph" in children[0] ? children[0].paragraph : undefined;
    if (typeof paragraph === "string" || !paragraph?.children)
      throw new Error("expected run children");
    expect(paragraph.children).toEqual([{ children: [{ tab: true }, { text: "AB☺" }] }]);
  });

  it("maps fonts, sizes, colors, and underline", () => {
    const paragraph = firstParagraph(
      String.raw`{\rtf1{\fonttbl{\f0 Arial;}}{\colortbl;\red255\green0\blue0;}\f0\fs20\cf1\ul Red}`,
    );
    expect(runs(paragraph)[0]).toEqual({
      font: "Arial",
      size: 10,
      color: "FF0000",
      underline: { type: "single" },
      text: "Red",
    });
  });

  it("ignores starred unknown destinations and known metadata destinations", () => {
    const source = String.raw`{\rtf1Visible{\*\generator Hidden}{\info{\title Also Hidden}}After}`;
    expect(parseRtf(source).sections[0]?.children[0]).toEqual({
      paragraph: { children: [{ text: "VisibleAfter" }] },
    });
  });

  it("skips an unlisted destination-like control group", () => {
    const children =
      parseRtf(String.raw`{\rtf1Keep{\nonshppict Hidden}}`).sections[0]?.children ?? [];
    const paragraph = children[0] && "paragraph" in children[0] ? children[0].paragraph : undefined;
    expect(typeof paragraph === "string" ? paragraph : paragraph?.children?.[0]).toEqual({
      text: "Keep",
    });
  });
});

describe("parseRtf tables", () => {
  it("projects cell boundaries and twip column widths", () => {
    const table = firstTable(String.raw`{\rtf1\trowd\cellx1000\cellx2000 A\cell B\cell\row}`);
    expect(table.columnWidths).toEqual([1000, 2000]);
    expect(table.rows).toHaveLength(1);
    const row = table.rows[0];
    if (!row || !("cells" in row)) throw new Error("expected a table row");
    expect(row.cells).toHaveLength(2);
  });

  it("supports a nested table inside an outer cell", () => {
    const table = firstTable(
      String.raw`{\rtf1\trowd\cellx2000 Outer{\trowd\cellx800 Inner\cell}\nestrow End\cell\row}`,
    );
    const row = table.rows[0];
    if (!row || !("cells" in row)) throw new Error("expected a table row");
    const cell = row.cells.find(
      (candidate): candidate is TableCellOptions => "children" in candidate,
    );
    const blocks: SectionChild[] = cell?.children ?? [];
    const nested = blocks.find(
      (block): block is Extract<SectionChild, { table: TableOptions }> => "table" in block,
    );
    expect(nested?.table.columnWidths).toEqual([800]);
    expect(JSON.stringify(blocks)).toContain("Outer");
    expect(JSON.stringify(blocks)).toContain("Inner");
    expect(JSON.stringify(blocks)).toContain("End");
  });

  it("keeps consecutive RTF rows in one docx table", () => {
    const source = String.raw`{\rtf1\trowd\cellx100 A\cell\row\trowd\cellx100 B\cell\row}`;
    const table = firstTable(source);
    const rows = table.rows.filter(
      (candidate): candidate is Extract<typeof candidate, { cells: unknown }> =>
        "cells" in candidate,
    );
    expect(rows).toHaveLength(2);
  });

  it("recovers when a row definition is repeated after cells", () => {
    const table = firstTable(
      String.raw`{\rtf1\trowd\cellx100 A\cell B\cell\trowd\cellx200 C\cell\row}`,
    );
    const rows = table.rows.filter((candidate) => "cells" in candidate);
    expect(rows).toHaveLength(2);
  });

  it("recovers missing and extra document group braces", () => {
    const source = String.raw`{\rtf1{\b Text}\par}}`;
    expect(parseRtf(source).sections[0]?.children[0]).toEqual({
      paragraph: { children: [{ bold: true, text: "Text" }] },
    });
  });
});

describe("parseRtf recovery", () => {
  it("rejects input without an RTF root", () => {
    expect(() => parseRtf("plain text")).toThrow('RTF must begin with "{\\rtf"');
  });

  it("closes an unterminated group and table at end of input", () => {
    const document = parseRtf(String.raw`{\rtf1 Unclosed\trowd\cellx100 Text\cell`);
    expect(document.sections[0]?.children[0]).toEqual({
      paragraph: { children: [{ text: "Unclosed" }] },
    });
  });

  it("keeps parsing content after a premature root closing brace", () => {
    const source = String.raw`{\rtf1 First}\par Second`;
    const children = parseRtf(source).sections[0]?.children ?? [];
    expect(children).toEqual([
      { paragraph: { children: [{ text: "First" }] } },
      { paragraph: { children: [{ text: "Second" }] } },
    ]);
  });
});
