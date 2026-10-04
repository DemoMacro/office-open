import type { DocumentOptions } from "@office-open/docx";
import { strFromU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import {
  generateChartDocument,
  generateOdp,
  generateOds,
  generateOdt,
  parseChartDocument,
  parseOdp,
  parseOds,
  parseOdt,
} from "./index";
import type { ChartChartOptions } from "./index";
import { type OdpOptions } from "./odp";
import { type OdsOptions } from "./ods";
import type { OdtTextSectionOptions } from "./odt";

describe("ODF package contract", () => {
  it("stores an uncompressed leading mimetype and a complete manifest", () => {
    const data = generateOdt({ sections: [{ children: [{ paragraph: "ODF" }] }] });
    const entries = unzipSync(data);
    expect(strFromU8(entries.mimetype!)).toBe("application/vnd.oasis.opendocument.text");
    expect(data.subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    const manifest = strFromU8(entries["META-INF/manifest.xml"]!);
    expect(manifest).toContain('manifest:version="1.3"');
    for (const path of ["/", "mimetype", "content.xml", "styles.xml", "meta.xml"]) {
      expect(manifest).toContain(`manifest:full-path="${path}"`);
    }
  });

  it("rejects a package whose manifest omits a payload entry", () => {
    const data = generateOdt({ sections: [] });
    const entries = unzipSync(data);
    entries["META-INF/manifest.xml"] = new TextEncoder().encode(
      '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="mimetype" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>',
    );
    const tampered = zipSync(entries);
    expect(() => parseOdt(tampered)).toThrow(/Manifest does not declare content\.xml/);
  });

  it("accepts legacy manifests without a version or mimetype entry", () => {
    const data = generateOdt({ sections: [{ children: [{ paragraph: "ODF" }] }] });
    const entries = unzipSync(data);
    entries["META-INF/manifest.xml"] = new TextEncoder().encode(
      strFromU8(entries["META-INF/manifest.xml"]!)
        .replace(' manifest:version="1.3"', "")
        .replace(/<manifest:file-entry manifest:full-path="mimetype"[^>]*>/u, ""),
    );
    expect(() => parseOdt(zipSync(entries))).not.toThrow();
  });

  it("accepts default-namespace manifests from compatibility writers", () => {
    const data = generateOdt({ sections: [{ children: [{ paragraph: "ODF" }] }] });
    const entries = unzipSync(data);
    entries["META-INF/manifest.xml"] = new TextEncoder().encode(
      '<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE manifest><manifest xmlns="urn:oasis:names:tc:opendocument:xmlns:manifest:1.3" version="1.3"><file-entry full-path="/" version="1.3" media-type="application/vnd.oasis.opendocument.text"/><file-entry full-path="mimetype" media-type="application/vnd.oasis.opendocument.text"/><file-entry full-path="content.xml" media-type="text/xml"/><file-entry full-path="styles.xml" media-type="text/xml"/><file-entry full-path="meta.xml" media-type="text/xml"/></manifest>',
    );
    expect(() => parseOdt(zipSync(entries))).not.toThrow();
  });

  it("treats document metadata as optional", () => {
    const data = generateOdt({ sections: [] });
    const entries = unzipSync(data);
    delete entries["meta.xml"];
    expect(parseOdt(zipSync(entries)).sections).toHaveLength(1);
  });
});

describe("ODT mapping", () => {
  it("round-trips typed text sections", () => {
    const sections: OdtTextSectionOptions[] = [
      {
        name: "Notes",
        styleName: "S1",
        protected: true,
        children: [{ paragraph: { text: "Inside" } }],
      },
    ];
    const parsed = parseOdt(
      generateOdt({
        sections: [{ children: [{ paragraph: "Before" }] }],
        textSections: sections,
      }),
    );
    expect(parsed.sections![0]!.children).toEqual([
      { paragraph: { text: "Before" } },
      { paragraph: { text: "Inside" } },
    ]);
    expect(parsed.textSections).toEqual(sections);
  });

  it("round-trips inline pictures through draw:image", () => {
    const data = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    picture: {
                      type: "png",
                      data,
                      transformation: { width: 9525, height: 9525 },
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const paragraph = parsed.sections[0]?.children[0] as {
      paragraph: { children?: Array<{ picture?: { type: string } }> };
    };
    expect(paragraph.paragraph.children?.[0]?.picture?.type).toBe("png");
  });

  it("round-trips page size and margins through page-layout", () => {
    const source: DocumentOptions = {
      sections: [
        {
          properties: {
            pageSize: { width: 11906, height: 16838 },
            pageMargin: { top: 1134, right: 850, bottom: 1134, left: 850 },
          },
          children: [{ paragraph: "Body" }],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const properties = parsed.sections[0]?.properties;
    expect(properties?.pageSize).toEqual({ width: 11906, height: 16838 });
    expect(properties?.pageMargin).toEqual({
      top: 1134,
      right: 850,
      bottom: 1134,
      left: 850,
    });
  });

  it("round-trips extended core properties through meta.xml", () => {
    const source: DocumentOptions = {
      title: "Meta",
      category: "Report",
      language: "zh-CN",
      identifier: "DOC-42",
      lastPrinted: "2026-02-03T04:05:06Z",
      sections: [{ children: [{ paragraph: "Body" }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    expect(parsed.category).toBe("Report");
    expect(parsed.language).toBe("zh-CN");
    expect(parsed.identifier).toBe("DOC-42");
    expect(parsed.lastPrinted).toBe("2026-02-03T04:05:06Z");
  });

  it("round-trips paragraphs, run formatting, and tables", () => {
    const source: DocumentOptions = {
      title: "ODT round trip",
      creator: "office-open",
      created: "2026-01-02T03:04:05Z",
      sections: [
        {
          children: [
            {
              paragraph: {
                alignment: "center",
                heading: "Heading1",
                children: [{ text: "Formatted", bold: true, italic: true, size: 14 }],
              },
            },
            {
              table: {
                columnWidths: [5670, 5670, 5670],
                rows: [
                  {
                    cells: [
                      { children: [{ paragraph: "Cell" }], columnSpan: 2 },
                      { children: [{ paragraph: "Covered" }] },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    expect(parsed.title).toBe(source.title);
    expect(parsed.creator).toBe(source.creator);
    expect(parsed.created).toBe(source.created);
    expect(parsed.sections).toHaveLength(1);
    const firstChild = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in firstChild)) throw new Error("Expected an ODT paragraph");
    const paragraph =
      typeof firstChild.paragraph === "string"
        ? { text: firstChild.paragraph }
        : firstChild.paragraph;
    expect(paragraph.alignment).toBe("center");
    expect(paragraph.heading).toBe("Heading1");
    expect(paragraph.children![0]).toMatchObject({
      text: "Formatted",
      bold: true,
      italic: true,
      size: 14,
    });
    const secondChild = parsed.sections[0]!.children[1]!;
    if (!("table" in secondChild)) throw new Error("Expected an ODT table");
    const table = secondChild.table;
    expect(table.columnWidths![0]).toBeCloseTo(5670, 0);
    const row = table.rows[0]!;
    if (!("cells" in row)) throw new Error("Expected table cells");
    const cell = row.cells[0]!;
    if (!("children" in cell)) throw new Error("Expected a table cell");
    expect(cell.columnSpan).toBe(2);
    const cellChild = cell.children[0]!;
    if (!("paragraph" in cellChild)) throw new Error("Expected a cell paragraph");
    const cellParagraph =
      typeof cellChild.paragraph === "string" ? { text: cellChild.paragraph } : cellChild.paragraph;
    expect(cellParagraph.text).toBe("Cell");
  });

  it("round-trips bullet lists, line breaks, and page breaks", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            { paragraph: { text: "First", bullet: { level: 0 } } },
            { paragraph: { text: "Nested", bullet: { level: 1 } } },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const children = parsed.sections[0]!.children;
    const first = children[0]!;
    if (!("paragraph" in first)) throw new Error("Expected a list paragraph");
    const firstParagraph =
      typeof first.paragraph === "string" ? { text: first.paragraph } : first.paragraph;
    expect(firstParagraph.text).toBe("First");
    expect(firstParagraph.bullet).toEqual({ level: 0 });
    const second = children[1]!;
    if (!("paragraph" in second)) throw new Error("Expected a nested list paragraph");
    const secondParagraph =
      typeof second.paragraph === "string" ? { text: second.paragraph } : second.paragraph;
    expect(secondParagraph.bullet).toEqual({ level: 1 });
  });

  it("round-trips numbered lists through a number list style", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            { paragraph: { text: "First", numbering: { reference: "num", level: 0 } } },
            { paragraph: { text: "Second", numbering: { reference: "num", level: 0 } } },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const children = parsed.sections[0]!.children;
    const first = children[0]!;
    if (!("paragraph" in first)) throw new Error("Expected a list paragraph");
    const firstParagraph =
      typeof first.paragraph === "string" ? { text: first.paragraph } : first.paragraph;
    expect(firstParagraph.numbering).toMatchObject({ level: 0 });
    const second = children[1]!;
    if (!("paragraph" in second)) throw new Error("Expected a list paragraph");
    const secondParagraph =
      typeof second.paragraph === "string" ? { text: second.paragraph } : second.paragraph;
    expect(secondParagraph.numbering).toMatchObject({ level: 0 });
  });

  it("round-trips hyperlinks as text:a elements", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    hyperlink: {
                      url: "https://example.com",
                      tooltip: "Example",
                      children: ["Site"],
                    },
                  },
                  { hyperlink: { anchor: "top", children: ["Jump"] } },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? { children: [] } : child.paragraph;
    const runs = paragraph.children as Array<{
      hyperlink?: { url?: string; anchor?: string; tooltip?: string; children?: string[] };
    }>;
    expect(runs[0]?.hyperlink).toMatchObject({
      url: "https://example.com",
      tooltip: "Example",
    });
    expect(runs[0]?.hyperlink?.children?.join("")).toBe("Site");
    expect(runs[1]?.hyperlink).toMatchObject({ anchor: "top" });
    expect(runs[1]?.hyperlink?.children?.join("")).toBe("Jump");
  });

  it("round-trips bookmarks as text:bookmark elements", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            { paragraph: { children: [{ bookmark: { name: "intro" } }, { text: "Body" }] } },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? { children: [] } : child.paragraph;
    const runs = paragraph.children as Array<{ bookmark?: { name: string } }>;
    expect(runs[0]?.bookmark).toEqual({ name: "intro" });
  });

  it("round-trips footnotes through text:note bodies", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ footnoteReference: 1 }, { text: "Body" }],
              },
            },
          ],
        },
      ],
      footnotes: [{ id: 1, children: [{ paragraph: "Note text" }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? { children: [] } : child.paragraph;
    const runs = paragraph.children as Array<{ footnoteReference?: number }>;
    expect(runs[0]?.footnoteReference).toBe(1);
    const note = parsed.footnotes?.[0];
    expect(note?.id).toBe(1);
    const noteChild = note?.children[0];
    if (!noteChild || typeof noteChild === "string" || !("paragraph" in noteChild))
      throw new Error("Expected a note paragraph");
    expect(noteChild.paragraph).toEqual({ text: "Note text" });
  });

  it("round-trips endnotes through text:note bodies", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ endnoteReference: 1 }, { text: "Body" }],
              },
            },
          ],
        },
      ],
      endnotes: [{ id: 1, children: [{ paragraph: "Endnote text" }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? { children: [] } : child.paragraph;
    const runs = paragraph.children as Array<{ endnoteReference?: number }>;
    expect(runs[0]?.endnoteReference).toBe(1);
    const note = parsed.endnotes?.[0];
    expect(note?.id).toBe(1);
    const noteChild = note?.children[0];
    if (!noteChild || typeof noteChild === "string" || !("paragraph" in noteChild))
      throw new Error("Expected a note paragraph");
    expect(noteChild.paragraph).toEqual({ text: "Endnote text" });
  });

  it("preserves runs of multiple spaces through text:s", () => {
    const source: DocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "a  b   c" } }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? { text: "" } : child.paragraph;
    expect(paragraph.text).toBe("a  b   c");
  });

  it("round-trips tab stops through paragraph tab-stop styles", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                text: "Tabbed",
                tabStops: [
                  { type: "right", position: 9026, leader: "dot" },
                  { type: "decimal", position: 4513 },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? {} : child.paragraph;
    expect(paragraph.tabStops).toEqual([
      { type: "right", position: 9026, leader: "dot" },
      { type: "decimal", position: 4513 },
    ]);
  });

  it("round-trips font declarations through font-face-decls", () => {
    const source: DocumentOptions = {
      fonts: [{ name: "Arial", family: "swiss", pitch: "variable" }],
      sections: [{ children: [{ paragraph: "Body" }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    expect(parsed.fonts).toEqual([{ name: "Arial", family: "swiss", pitch: "variable" }]);
  });

  it("round-trips footnote numbering configuration", () => {
    const source: DocumentOptions = {
      sections: [{ children: [{ paragraph: "Body" }] }],
      settings: {
        footnoteProperties: { pos: "pageBottom", numFmt: "decimal", numStart: 2 },
        endnoteProperties: { pos: "docEnd", numFmt: "lowerRoman" },
      },
    };
    const parsed = parseOdt(generateOdt(source));
    expect(parsed.settings?.footnoteProperties).toMatchObject({
      pos: "pageBottom",
      numFmt: "decimal",
      numStart: 2,
    });
    expect(parsed.settings?.endnoteProperties).toMatchObject({
      pos: "docEnd",
      numFmt: "lowerRoman",
    });
  });

  it("round-trips headers and footers through the master page", () => {
    const source: DocumentOptions = {
      sections: [
        {
          headers: {
            default: [
              {
                paragraph: {
                  children: [
                    "Header ",
                    { simpleField: { instruction: " PAGE ", cachedValue: "1" } },
                  ],
                },
              },
            ],
          },
          footers: {
            default: [{ paragraph: { text: "Footer text" } }],
          },
          children: [{ paragraph: "Body" }],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const headerParagraph = parsed.sections[0]!.headers!.default![0]!;
    if (!("paragraph" in headerParagraph)) throw new Error("Expected a header paragraph");
    const header = typeof headerParagraph.paragraph === "string" ? {} : headerParagraph.paragraph;
    expect(header.children).toEqual([
      "Header ",
      { simpleField: { instruction: " PAGE ", cachedValue: "1" } },
    ]);
    const footerParagraph = parsed.sections[0]!.footers!.default![0]!;
    if (!("paragraph" in footerParagraph)) throw new Error("Expected a footer paragraph");
    const footer = typeof footerParagraph.paragraph === "string" ? {} : footerParagraph.paragraph;
    expect(footer.text).toBe("Footer text");
  });

  it("round-trips heading outline numbering through text:outline-style", () => {
    const source: DocumentOptions = {
      numbering: {
        abstractNumberings: [
          {
            reference: "Outline",
            levels: [
              { level: 0, format: "decimal", text: "%1.", start: 1, paragraphStyle: "Heading1" },
              {
                level: 1,
                format: "lowerLetter",
                text: "%2)",
                start: 1,
                paragraphStyle: "Heading2",
              },
            ],
          },
        ],
      },
      sections: [
        {
          children: [
            {
              paragraph: {
                heading: "Heading1",
                text: "Chapter",
                numbering: { reference: "Outline", level: 0 },
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const numbering = parsed.numbering?.abstractNumberings?.[0];
    expect(numbering?.reference).toBe("Outline");
    expect(numbering?.levels[0]).toMatchObject({
      level: 0,
      format: "decimal",
      text: "%1.",
      paragraphStyle: "Heading1",
    });
    expect(numbering?.levels[1]).toMatchObject({
      level: 1,
      format: "lowerLetter",
      text: "%2)",
      paragraphStyle: "Heading2",
    });
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? {} : child.paragraph;
    expect(paragraph.heading).toBe("Heading1");
    expect(paragraph.numbering).toEqual({ reference: "Outline", level: 0 });
  });

  it("round-trips document defaults through style:default-style", () => {
    const source: DocumentOptions = {
      styles: {
        default: {
          document: {
            paragraph: { alignment: "center", indent: { left: 567 } },
            run: { font: "Arial", size: 12, bold: true },
          },
        },
      },
      sections: [{ children: [{ paragraph: "Body" }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    const defaults = parsed.styles?.default?.document;
    expect(defaults?.paragraph).toMatchObject({ alignment: "center", indent: { left: 567 } });
    expect(defaults?.run).toMatchObject({ font: "Arial", size: 12, bold: true });
  });

  it("round-trips a line break inside a run", () => {
    const source: DocumentOptions = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ text: "before", break: 2 }, { text: "after" }],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph = typeof child.paragraph === "string" ? { children: [] } : child.paragraph;
    expect(paragraph.children).toBeDefined();
    const run = paragraph.children![0]! as { break?: number; text?: string };
    expect(run.break).toBe(1);
  });

  it("round-trips pageBreakBefore via fo:break-before", () => {
    const source: DocumentOptions = {
      sections: [{ children: [{ paragraph: { text: "New page", pageBreakBefore: true } }] }],
    };
    const parsed = parseOdt(generateOdt(source));
    const child = parsed.sections[0]!.children[0]!;
    if (!("paragraph" in child)) throw new Error("Expected a paragraph");
    const paragraph =
      typeof child.paragraph === "string" ? { pageBreakBefore: undefined } : child.paragraph;
    expect(paragraph.pageBreakBefore).toBe(true);
  });
});

describe("ODS mapping", () => {
  it("round-trips typed chart subdocument structure", () => {
    const chart: ChartChartOptions = {
      class: "chart:bar",
      width: 5760000,
      height: 3240000,
      href: "..",
      styleName: "ch1",
      title: { text: "Chart title", x: 0, y: 0, width: 5760000, height: 448056 },
      legend: {
        position: "end",
        x: 5402208,
        y: 1512360,
        expansion: "high",
        styleName: "ch2",
      },
      plotArea: {
        x: 115200,
        y: 64800,
        width: 5172120,
        height: 3110400,
        styleName: "ch3",
        cellRange: "Sheet1.A1:Sheet1.A2",
        dataSourceHasLabels: "row",
        axes: [
          { dimension: "x", name: "primary-x", styleName: "ch4" },
          {
            dimension: "y",
            name: "primary-y",
            styleName: "ch5",
            grids: [{ class: "major", styleName: "ch6" }],
          },
        ],
        series: [
          {
            values: "Sheet1.A2:Sheet1.A2",
            label: "Sheet1.A1:Sheet1.A1",
            class: "chart:bar",
            styleName: "ch7",
            domains: ["Sheet1.A1:Sheet1.A1"],
            dataPoints: [{ repeated: 1 }],
          },
        ],
        wall: { styleName: "ch8" },
        floor: { styleName: "ch9" },
      },
    };
    const parsed = parseChartDocument(generateChartDocument({ title: "Chart", chart })).chart!;
    expect(parsed).toMatchObject(chart);
  });

  it("round-trips sheets, dimensions, cached formulas, and primitive values", () => {
    const source = {
      title: "ODS round trip",
      worksheets: [
        {
          name: "Data",
          columns: [{ min: 1, max: 2, width: 120 }],
          rows: [
            { height: 18, cells: [{ value: "Text" }, { value: true }] },
            { cells: [{ value: 4.5 }, { value: 9, formula: "SUM(A1:B1)" }] },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(source));
    expect(parsed.title).toBe(source.title);
    const sheet = parsed.worksheets![0]!;
    expect(sheet.name).toBe("Data");
    expect(sheet.columns).toEqual([
      { min: 1, max: 1, width: 120 },
      { min: 2, max: 2, width: 120 },
    ]);
    expect(sheet.rows![0]!.height).toBe(18);
    expect(sheet.rows![0]!.cells).toEqual([
      { reference: "A1", value: "Text" },
      { reference: "B1", value: true },
    ]);
    expect(sheet.rows![1]!.cells![0]).toMatchObject({ reference: "A2", value: 4.5 });
    expect(sheet.rows![1]!.cells![1]).toMatchObject({
      reference: "B2",
      value: 9,
      formula: "SUM(A1:B1)",
    });
  });

  it("round-trips date cell values through office:date-value", () => {
    const date = new Date("2026-02-03T04:05:06Z");
    const source = {
      worksheets: [{ rows: [{ cells: [{ value: date }] }] }],
    };
    const parsed = parseOds(generateOds(source));
    expect(parsed.worksheets![0]!.rows![0]!.cells![0]!.value).toEqual(date);
  });

  it("round-trips number formats through data styles", () => {
    const source = {
      worksheets: [
        {
          rows: [
            {
              cells: [
                { value: 1.5, style: { numFmt: "0.00" } },
                { value: 0.25, style: { numFmt: "0.00%" } },
                { value: 1200, style: { numFmt: "#,##0" } },
              ],
            },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(source));
    const cells = parsed.worksheets![0]!.rows![0]!.cells!;
    expect(cells[0]!.style).toEqual({ numFmt: "0.00" });
    expect(cells[1]!.style).toEqual({ numFmt: "0.00%" });
    expect(cells[2]!.style).toEqual({ numFmt: "#,##0" });
  });

  it("round-trips cell fonts and fills through table-cell styles", () => {
    const source: OdsOptions = {
      worksheets: [
        {
          rows: [
            {
              cells: [
                {
                  value: "Styled",
                  style: {
                    font: { bold: true, italic: true, color: "4472C4", size: 12 },
                    fill: { type: "solid", color: "C6EFCE" },
                  },
                },
              ],
            },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(source));
    expect(parsed.worksheets![0]!.rows![0]!.cells![0]!.style).toEqual({
      font: { bold: true, italic: true, color: "4472C4", size: 12 },
      fill: { type: "solid", color: "C6EFCE" },
    });
  });

  it("round-trips cell alignment through style properties", () => {
    const source: OdsOptions = {
      worksheets: [
        {
          rows: [
            {
              cells: [
                {
                  value: "Centered",
                  style: {
                    alignment: { horizontal: "center", vertical: "center", wrapText: true },
                  },
                },
              ],
            },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(source));
    expect(parsed.worksheets![0]!.rows![0]!.cells![0]!.style).toEqual({
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
    });
  });

  it("round-trips cell borders through border shorthand styles", () => {
    const source: OdsOptions = {
      worksheets: [
        {
          rows: [
            {
              cells: [
                {
                  value: "Bordered",
                  style: {
                    border: {
                      top: { style: "thin", color: "000000" },
                      bottom: { style: "double", color: "4472C4" },
                      left: { style: "dashed", color: "C00000" },
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(source));
    expect(parsed.worksheets![0]!.rows![0]!.cells![0]!.style).toEqual({
      border: {
        top: { style: "thin", color: "000000" },
        bottom: { style: "double", color: "4472C4" },
        left: { style: "dashed", color: "C00000" },
      },
    });
  });

  it("round-trips diagonal cell borders through ODF diagonal styles", () => {
    const source: OdsOptions = {
      worksheets: [
        {
          rows: [
            {
              cells: [
                {
                  value: "Diagonal",
                  style: {
                    border: {
                      diagonal: { style: "thin", color: "7030A0" },
                      diagonalDown: true,
                      diagonalUp: true,
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    };
    const parsed = parseOds(generateOds(source));
    expect(parsed.worksheets![0]!.rows![0]!.cells![0]!.style).toEqual({
      border: {
        diagonal: { style: "thin", color: "7030A0" },
        diagonalDown: true,
        diagonalUp: true,
      },
    });
  });
});

describe("ODP mapping", () => {
  it("round-trips page size and positioned text shapes", () => {
    const source = {
      size: { width: 9144000, height: 6858000 },
      slides: [
        {
          children: [
            {
              shape: {
                name: "Title",
                x: 914400,
                y: 914400,
                width: 3657600,
                height: 914400,
                textBody: {
                  paragraphs: [
                    { text: "Slide text" },
                    { children: [{ text: "Bold", bold: true, size: 20 }] },
                  ],
                },
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdp(generateOdp(source));
    expect(parsed.size).toEqual(source.size);
    expect(parsed.slides).toHaveLength(1);
    const slideChild = parsed.slides![0]!.children![0]!;
    if (!("shape" in slideChild)) throw new Error("Expected an ODP shape");
    const shape = slideChild.shape;
    expect(shape.name).toBe("Title");
    expect(shape.x).toBe(914400);
    expect(shape.y).toBe(914400);
    expect(shape.width).toBe(3657600);
    expect(shape.height).toBe(914400);
    expect(shape.textBody!.paragraphs![0]).toEqual({ text: "Slide text" });
    expect(shape.textBody!.paragraphs![1]).toMatchObject({
      children: [{ text: "Bold", bold: true, size: 20 }],
    });
  });

  it("round-trips speaker notes through presentation:notes", () => {
    const source = { slides: [{ notes: "Speaker note" }] };
    const parsed = parseOdp(generateOdp(source));
    expect(parsed.slides![0]!.notes).toBe("Speaker note");
  });

  it("round-trips slide tables through table:table", () => {
    const source: OdpOptions = {
      slides: [
        {
          children: [
            {
              table: {
                columnWidths: [1828800, 1828800],
                rows: [
                  { cells: [{ text: "Cell A" }, { text: "Cell B" }] },
                  { cells: [{ text: "C" }, { text: "D", columnSpan: 1 }] },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdp(generateOdp(source));
    const child = parsed.slides![0]!.children![0]!;
    if (!("table" in child)) throw new Error("Expected an ODP table");
    expect(child.table.columnWidths).toEqual([1828800, 1828800]);
    expect(child.table.rows[0]!.cells[0]!.children).toEqual([{ text: "Cell A" }]);
    expect(child.table.rows[1]!.cells[1]!.columnSpan).toBe(1);
  });

  it("round-trips connectors through draw:connector", () => {
    const source = {
      slides: [
        {
          children: [
            {
              connector: {
                x1: 914400,
                y1: 914400,
                x2: 3657600,
                y2: 3657600,
                startConnection: { id: 2, index: 3 },
                endConnection: { id: 5, index: 1 },
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdp(generateOdp(source));
    const child = parsed.slides![0]!.children![0]!;
    if (!("connector" in child)) throw new Error("Expected an ODP connector");
    expect(child.connector.x1).toBe(914400);
    expect(child.connector.y2).toBe(3657600);
    expect(child.connector.startConnection).toEqual({ id: 2, index: 3 });
    expect(child.connector.endConnection).toEqual({ id: 5, index: 1 });
  });

  it("round-trips slide pictures through draw:image", () => {
    const data = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const source: OdpOptions = {
      slides: [
        {
          children: [
            {
              picture: { type: "png", data, x: 914400, y: 914400, width: 1828800, height: 1828800 },
            },
          ],
        },
      ],
    };
    const parsed = parseOdp(generateOdp(source));
    const child = parsed.slides![0]!.children![0]!;
    if (!("picture" in child)) throw new Error("Expected an ODP picture");
    expect(child.picture.type).toBe("png");
    expect(child.picture.x).toBe(914400);
    expect(child.picture.width).toBe(1828800);
  });

  it("round-trips lines through draw:line", () => {
    const source: OdpOptions = {
      slides: [{ children: [{ line: { x1: 914400, y1: 914400, x2: 3657600, y2: 3657600 } }] }],
    };
    const parsed = parseOdp(generateOdp(source));
    const child = parsed.slides![0]!.children![0]!;
    if (!("line" in child)) throw new Error("Expected an ODP line");
    expect(child.line.x1).toBe(914400);
    expect(child.line.y2).toBe(3657600);
  });

  it("round-trips groups through draw:g", () => {
    const source: OdpOptions = {
      slides: [
        {
          children: [
            {
              group: {
                children: [
                  { line: { x1: 914400, y1: 914400, x2: 1828800, y2: 1828800 } },
                  { shape: { x: 1828800, y: 1828800, width: 914400, height: 914400 } },
                ],
              },
            },
          ],
        },
      ],
    };
    const parsed = parseOdp(generateOdp(source));
    const child = parsed.slides![0]!.children![0]!;
    if (!("group" in child)) throw new Error("Expected an ODP group");
    expect(child.group.children).toHaveLength(2);
    const nested = child.group.children[0]!;
    if (!("line" in nested)) throw new Error("Expected a nested line");
    expect(nested.line.x2).toBe(1828800);
  });

  it("preserves chart frames verbatim through rawXml", () => {
    const source: OdpOptions = {
      slides: [
        {
          children: [
            {
              rawXml:
                '<draw:frame svg:x="1cm" svg:y="1cm"><draw:object xlink:href="./Object 1"/></draw:frame>',
            },
          ],
        },
      ],
    };
    const parsed = parseOdp(generateOdp(source));
    const child = parsed.slides![0]!.children![0]!;
    if (!("rawXml" in child)) throw new Error("Expected rawXml preservation");
    expect(child.rawXml).toContain("draw:object");
  });

  it("flattens text sections in document order", () => {
    const data = generateOdt({
      sections: [{ children: [{ paragraph: "Before" }, { paragraph: "After" }] }],
    });
    const entries = unzipSync(data);
    entries["content.xml"] = new TextEncoder().encode(
      strFromU8(entries["content.xml"]!).replace(
        /<text:p>After<\/text:p>/u,
        '<text:section text:name="S1"><text:p>Inside</text:p></text:section><text:p>After</text:p>',
      ),
    );
    const parsed = parseOdt(zipSync(entries));
    expect(parsed.odfExtensions).toHaveLength(0);
    const texts = parsed.sections[0]!.children.map((child) =>
      "paragraph" in child && typeof child.paragraph === "object"
        ? child.paragraph.text
        : "paragraph" in child
          ? typeof child.paragraph === "string"
            ? child.paragraph
            : undefined
          : undefined,
    );
    expect(texts).toEqual(["Before", "Inside", "After"]);
  });
});
