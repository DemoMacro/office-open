import type { DocumentOptions } from "@office-open/docx";
import { strFromU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateOdp, generateOds, generateOdt, parseOdp, parseOds, parseOdt } from "./index";

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

  it("treats document metadata as optional", () => {
    const data = generateOdt({ sections: [] });
    const entries = unzipSync(data);
    delete entries["meta.xml"];
    expect(parseOdt(zipSync(entries)).sections).toHaveLength(1);
  });
});

describe("ODT mapping", () => {
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
});
