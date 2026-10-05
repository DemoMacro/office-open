import type {
  ParagraphOptions,
  RunOptions,
  SectionChild,
  TableCellOptions,
  TableOptions,
} from "@office-open/docx";
import { describe, expect, it } from "vite-plus/test";

import { parseDocument } from ".";
import { generateDocument } from "./generate";

function roundTrip(source: string): ReturnType<typeof parseDocument> {
  const options = parseDocument(source);
  return parseDocument(new TextDecoder().decode(generateDocument(options)));
}

function firstParagraph(source: string): ParagraphOptions {
  const children = parseDocument(source).sections[0]?.children ?? [];
  const child = children[0];
  if (!child || !("paragraph" in child) || typeof child.paragraph === "string") {
    throw new Error("expected the first child to be a paragraph");
  }
  return child.paragraph;
}

function firstTable(source: string): TableOptions {
  const children = parseDocument(source).sections[0]?.children ?? [];
  const child = children[0];
  if (!child || !("table" in child)) throw new Error("expected the first child to be a table");
  return child.table;
}

function runs(paragraph: ParagraphOptions): RunOptions[] {
  return paragraph.children as RunOptions[];
}

function firstInlineChild(paragraph: ParagraphOptions): unknown {
  const [run] = runs(paragraph);
  const children = run?.children;
  return children ? children[0] : undefined;
}

describe("parseDocument RTF text and formatting", () => {
  it("parses standard text into a document paragraph", () => {
    expect(parseDocument("{\\rtf1 Hello, RTF!}")).toEqual({
      sections: [{ children: [{ paragraph: { children: [{ text: "Hello, RTF!" }] } }] }],
    });
  });

  it("scopes bold formatting to its group and maps paragraph alignment", () => {
    const paragraph = firstParagraph(String.raw`{\rtf1{\b Bold} plain\par}{\qc Centered\par}`);
    expect(runs(paragraph)).toEqual([{ bold: true, text: "Bold" }, { text: " plain" }]);
  });

  it("parses tabs, hexadecimal, Unicode, and common character controls", () => {
    const children =
      parseDocument(String.raw`{\rtf1\tab A\'42\u9786?}`).sections[0]?.children ?? [];
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
    expect(parseDocument(source).sections[0]?.children[0]).toEqual({
      paragraph: { children: [{ text: "VisibleAfter" }] },
    });
  });

  it("skips an unlisted destination-like control group", () => {
    const children =
      parseDocument(String.raw`{\rtf1Keep{\nonshppict Hidden}}`).sections[0]?.children ?? [];
    const paragraph = children[0] && "paragraph" in children[0] ? children[0].paragraph : undefined;
    expect(typeof paragraph === "string" ? paragraph : paragraph?.children?.[0]).toEqual({
      text: "Keep",
    });
  });
});

describe("parseDocument RTF canonical projections", () => {
  it("preserves metadata and run details through RTF round trip", () => {
    const source = String.raw`{\rtf1{\info{\edmins12}{\nofpages2}{\nofwords30}{\nofchars150}{\nofcharsws180}}{\super Up}{\nosupersub normal}{\sub Down}{\strike struck}}`;
    const parsed = parseDocument(source);
    expect(parsed.appProperties).toEqual({
      totalTime: 12,
      pages: 2,
      words: 30,
      characters: 150,
      charactersWithSpaces: 180,
    });
    const paragraph = firstParagraph(source);
    expect(runs(paragraph)).toEqual([
      { verticalAlign: "superscript", text: "Up" },
      { verticalAlign: "baseline", text: "normal" },
      { verticalAlign: "subscript", text: "Down" },
      { strike: true, text: "struck" },
    ]);
    expect(roundTrip(source)).toEqual(parsed);
  });

  it("preserves paragraph style references through RTF round trip", () => {
    const source = String.raw`{\rtf1{\stylesheet{\s7 Base;}}\s7 Body\par}`;
    const parsed = parseDocument(source);
    expect(firstParagraph(source).style).toBe("rtf-style-7");
    expect(roundTrip(source)).toEqual(parsed);
  });
});

describe("parseDocument RTF tables", () => {
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
    expect(parseDocument(source).sections[0]?.children[0]).toEqual({
      paragraph: { children: [{ bold: true, text: "Text" }] },
    });
  });
});

describe("parseDocument recovery", () => {
  it("rejects input without an RTF root", () => {
    expect(() => parseDocument("plain text")).toThrow('RTF must begin with "{\\rtf"');
  });

  it("closes an unterminated group and table at end of input", () => {
    const document = parseDocument(String.raw`{\rtf1 Unclosed\trowd\cellx100 Text\cell`);
    expect(document.sections[0]?.children[0]).toEqual({
      paragraph: { children: [{ text: "Unclosed" }] },
    });
  });

  it("keeps parsing content after a premature root closing brace", () => {
    const source = String.raw`{\rtf1 First}\par Second`;
    const children = parseDocument(source).sections[0]?.children ?? [];
    expect(children).toEqual([
      { paragraph: { children: [{ text: "First" }] } },
      { paragraph: { children: [{ text: "Second" }] } },
    ]);
  });
});

describe("parseDocument RTF rich destinations", () => {
  it("projects headers, footers, and footnotes", () => {
    const document = parseDocument(
      String.raw`{\rtf1{\header Header}{\footer Footer}{\footnote Note}Body}`,
    );
    expect(document.sections[0]?.headers?.default?.[0]).toEqual({
      paragraph: { children: [{ text: "Header" }] },
    });
    expect(document.sections[0]?.footers?.default?.[0]).toEqual({
      paragraph: { children: [{ text: "Footer" }] },
    });
    expect(document.footnotes?.[0]?.children[0]).toEqual({
      paragraph: { children: [{ text: "Note" }] },
    });
  });

  it("projects direct hyperlinks, comments, and canonical shapes", () => {
    const document = parseDocument(
      String.raw`{\rtf1{\hlink https://example.com}Link{\doccomm Summary}{\shpinst\shpleft100\shptop100\shpwidth1000\shpheight500}}`,
    );
    expect(document.sections[0]?.children[0]).toEqual({
      paragraph: {
        children: [
          { children: [{ hyperlink: { url: "https://example.com", children: ["Link"] } }] },
        ],
      },
    });
    expect(document.description).toBe("Summary");
    expect(document.sections.at(-1)?.children.at(-1)).toEqual({
      paragraph: {
        children: [
          {
            wpsShape: {
              children: [],
              transformation: { width: 635000, height: 317500 },
              floating: {
                horizontalPosition: { relative: "column", offset: 63500 },
                verticalPosition: { relative: "paragraph", offset: 63500 },
              },
            },
          },
        ],
      },
    });
  });

  it("rejects shape instructions without a canonical extent", () => {
    expect(() =>
      parseDocument(String.raw`{\rtf1{\shpinst\shptop1000}}`),
    ).toThrowErrorMatchingInlineSnapshot(
      `[RtfParseError: Invalid RTF at 1:1: shape width and height are required]`,
    );
    expect(() => parseDocument(String.raw`{\rtf1{\shpinst\shptop1000}}`)).toThrow(
      expect.objectContaining({
        context: {
          part: "RTF shape destination",
          path: "\\shpinst",
          name: "shape",
          reason: "shape width and height are required",
        },
      }),
    );
  });

  it("projects HYPERLINK, PAGE, and DATE fields", () => {
    const paragraph = firstParagraph(
      String.raw`{\rtf1{\field{\*\fldinst HYPERLINK "https://example.com"}{\fldrslt Link}}\par{\field{\*\fldinst PAGE}{\fldrslt 1}}\par{\field{\*\fldinst DATE}{\fldrslt Today}}}`,
    );
    expect(firstInlineChild(paragraph)).toEqual({
      hyperlink: { url: "https://example.com", children: ["Link"] },
    });
    const page = firstParagraph(String.raw`{\rtf1{\field{\*\fldinst PAGE}{\fldrslt 1}}}`);
    expect(firstInlineChild(page)).toEqual({
      complexField: { instruction: "PAGE", result: "1" },
    });
    const date = firstParagraph(String.raw`{\rtf1{\field{\*\fldinst DATE}{\fldrslt Today}}}`);
    expect(firstInlineChild(date)).toEqual({
      complexField: { instruction: "DATE", result: "Today" },
    });
  });

  it("extracts PNG picture bytes and twip dimensions", () => {
    const png = "89504e470d0a1a0a";
    const paragraph = firstParagraph(String.raw`{\rtf1{\pict\pngblip\picw100\pich50 ${png}}}`);
    expect(firstInlineChild(paragraph)).toEqual({
      picture: {
        type: "png",
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        transformation: { width: 63500, height: 31750 },
      },
    });
  });

  it("rejects a malformed picture with structured context", () => {
    expect(() => parseDocument(String.raw`{\rtf1{\pict\pngblip\picw0\pich0 zz}Text}`)).toThrow(
      expect.objectContaining({
        context: {
          part: "RTF picture destination",
          path: "\\pict",
          name: "picture",
          reason: "missing data or dimensions",
        },
      }),
    );
  });

  it("pairs bookmark starts and ends", () => {
    const children =
      parseDocument(String.raw`{\rtf1{\bkmkstart Mark}A{\bkmkend Mark}}`).sections[0]?.children ??
      [];
    expect(children).toHaveLength(1);
    const first = children[0];
    const inline =
      first && "paragraph" in first && typeof first.paragraph !== "string"
        ? firstInlineChild(first.paragraph)
        : undefined;
    expect(inline).toEqual({
      bookmarkStart: { id: 1, name: "Mark" },
    });
    const endParagraph = firstParagraph(String.raw`{\rtf1{\bkmkend Mark}}`);
    expect(firstInlineChild(endParagraph)).toEqual({
      bookmarkEnd: { id: 1 },
    });
  });

  it("projects list references and levels", () => {
    const document = parseDocument(String.raw`{\rtf1\ls3\ilvl1\par Item}`);
    expect(firstParagraph(String.raw`{\rtf1\ls3\ilvl1\par Item}`)).toMatchObject({
      numbering: { reference: "rtf-list-3", level: 1 },
    });
    expect(document.numbering?.abstractNumberings[0]).toMatchObject({
      reference: "rtf-list-3",
      levels: [{ level: 0, format: "decimal" }],
    });
  });

  it("projects a complete list table and override", () => {
    const source = String.raw`{\rtf1{\listtable{\list\listtemplateid9\listhybrid{\listlevel\levelnfc2\leveljc1\levelstartat3\levelindent720{\leveltext\'02\'00.;}{\levelnumbers;}}{\listname rtf-list-template-9;}\listid77}}{\listoverridetable{\listoverride\listid77\listoverridecount0\ls4}}\ls4\ilvl0\par Item}`;
    expect(parseDocument(source).numbering?.abstractNumberings).toEqual([
      {
        reference: "rtf-list-template-9",
        levels: [
          {
            level: 0,
            format: "lowerRoman",
            alignment: "center",
            start: 3,
            text: "%1.",
            paragraph: { indent: { left: 720 } },
          },
        ],
      },
    ]);
    expect(firstParagraph(source).numbering).toEqual({
      reference: "rtf-list-template-9",
      level: 0,
    });
  });

  it("projects stylesheet metadata and body formatting", () => {
    const source = String.raw`{\rtf1{\colortbl;\red255\green0\blue0;}{\stylesheet{\s7 Base;\sbasedon7\snext7\shidden\spriority9\sqformat\ql\li120\sb240\brdrb\brdrw15\cbpat1}{\cs8 Link;\b\i}}\s7 Body}`;
    const styles = parseDocument(source).styles;
    expect(styles?.paragraphStyles?.[0]).toMatchObject({
      id: "rtf-style-7",
      basedOn: "rtf-style-7",
      next: "rtf-style-7",
      hidden: true,
      uiPriority: 9,
      quickFormat: true,
      paragraph: {
        alignment: "left",
        indent: { left: 120 },
        spacing: { before: 240 },
        border: { bottom: { style: "single", size: 15 } },
        shading: { type: "clear", fill: "FF0000" },
      },
    });
    expect(styles?.characterStyles?.[0]).toMatchObject({
      id: "rtf-character-style-8",
      run: { bold: true, italic: true },
    });
  });

  it("rejects an unlisted non-starred destination", () => {
    expect(() => parseDocument(String.raw`{\rtf1Keep{\unknown Hidden}}`)).toThrow(
      expect.objectContaining({
        context: {
          part: "RTF destination",
          path: "destination/unknown",
          name: "unknown",
          reason: "no canonical DocumentOptions equivalent and writer does not regenerate it",
        },
      }),
    );
  });

  it("projects a complete list table and override", () => {
    const source = String.raw`{\rtf1{\listtable{\list\listtemplateid9\listhybrid{\listlevel\levelnfc2\leveljc1\levelstartat3\levelindent720{\leveltext\'02\'00.;}{\levelnumbers;}}{\listname rtf-list-template-9;}\listid77}}{\listoverridetable{\listoverride\listid77\listoverridecount0\ls4}}\ls4\ilvl0\par Item}`;
    expect(parseDocument(source).numbering?.abstractNumberings).toEqual([
      {
        reference: "rtf-list-template-9",
        levels: [
          {
            level: 0,
            format: "lowerRoman",
            alignment: "center",
            start: 3,
            text: "%1.",
            paragraph: { indent: { left: 720 } },
          },
        ],
      },
    ]);
    expect(firstParagraph(source).numbering).toEqual({
      reference: "rtf-list-template-9",
      level: 0,
    });
  });

  it("projects stylesheet metadata and body formatting", () => {
    const source = String.raw`{\rtf1{\colortbl;\red255\green0\blue0;}{\stylesheet{\s7 Base;\sbasedon7\snext7\shidden\spriority9\sqformat\ql\li120\sb240\brdrb\brdrw15\cbpat1}{\cs8 Link;\b\i}}\s7 Body}`;
    const styles = parseDocument(source).styles;
    expect(styles?.paragraphStyles?.[0]).toMatchObject({
      id: "rtf-style-7",
      basedOn: "rtf-style-7",
      next: "rtf-style-7",
      hidden: true,
      uiPriority: 9,
      quickFormat: true,
      paragraph: {
        alignment: "left",
        indent: { left: 120 },
        spacing: { before: 240 },
        border: { bottom: { style: "single", size: 15 } },
        shading: { type: "clear", fill: "FF0000" },
      },
    });
    expect(styles?.characterStyles?.[0]).toMatchObject({
      id: "rtf-character-style-8",
      run: { bold: true, italic: true },
    });
  });

  it("rejects an unlisted non-starred destination", () => {
    expect(() => parseDocument(String.raw`{\rtf1Keep{\unknown Hidden}}`)).toThrow(
      expect.objectContaining({
        context: {
          part: "RTF destination",
          path: "destination/unknown",
          name: "unknown",
          reason: "no canonical DocumentOptions equivalent and writer does not regenerate it",
        },
      }),
    );
  });

  it("projects section page size", () => {
    const document = parseDocument(String.raw`{\rtf1\sectd\pgwsxn10000\pghsxn12000\par Page}`);
    expect(document.sections[0]?.properties).toEqual({
      pageSize: { width: 10000, height: 12000 },
    });
  });

  it("binds section properties to the section after its break", () => {
    const document = parseDocument(
      String.raw`{\rtf1\par First\sect\sectd\pgwsxn10000\pghsxn12000\par Second}`,
    );
    expect(document.sections[0]?.properties).toBeUndefined();
    expect(document.sections[1]?.properties).toEqual({
      pageSize: { width: 10000, height: 12000 },
    });
  });

  it("resets inherited section properties on sectd", () => {
    const document = parseDocument(String.raw`{\rtf1\pghsxn12000\sectd\par Page}`);
    expect(document.sections[0]?.properties).toBeUndefined();
  });

  it("projects paragraph shading and borders", () => {
    const paragraph = firstParagraph(String.raw`{\rtf1\shading1000\brdrb\brdrs80\par Box}`);
    expect(paragraph.shading).toEqual({ type: "clear", fill: "auto" });
    expect(paragraph.border?.bottom).toEqual({ style: "single" });
  });

  it("recognizes an embedded object class", () => {
    expect(
      firstParagraph(String.raw`{\rtf1{\object{\objclass Excel.Sheet.8}{\objdata\bin4 ABCD}}}`),
    ).toEqual({
      children: [
        {
          children: [
            {
              object: {
                embed: { data: new Uint8Array([65, 66, 67, 68]), progId: "Excel.Sheet.8" },
              },
            },
          ],
        },
      ],
    });
  });
});
