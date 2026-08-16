// Combines every safe authoring domain in one document — text runs, styles,
// numbering, revisions, drawings, math, fields, tables, multi-section layout —
// to surface combination defects the single-feature demos cannot hit.
import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";
import type { DocumentOptions } from "@office-open/docx";

const png1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const doc: DocumentOptions = {
  styles: {
    paragraphStyles: [
      {
        id: "Lead",
        name: "Lead",
        basedOn: "Normal",
        run: { size: 28, color: "2E74B5", bold: true },
        paragraph: { spacing: { after: 120 } },
      },
    ],
    characterStyles: [{ id: "Emphasis", name: "Emphasis", run: { italic: true, color: "C00000" } }],
  },
  numbering: {
    config: [
      {
        reference: "max-numbering",
        levels: [
          { level: 0, format: "decimal", text: "%1.", alignment: "start" },
          { level: 1, format: "lowerLetter", text: "%2)", alignment: "start" },
        ],
      },
    ],
  },
  footnotes: {
    1: { children: ["Footnote body with detail."] },
  },
  endnotes: {
    1: { children: ["Endnote body."] },
  },
  comments: {
    children: [
      {
        id: 0,
        author: "Max",
        initials: "M",
        date: new Date(),
        children: [{ children: [{ text: "Document-level comment body." }] }],
      },
    ],
  },
  settings: {
    defaultTabStop: 720,
    trackRevisions: true,
    evenAndOddHeaders: true,
    displayBackgroundShape: true,
  },
  background: { color: "FDF6E3" },
  customProperties: [{ name: "Reviewed", value: true }],
  appProperties: { company: "Example", application: "office-open", appVersion: "1.0000" },
  sections: [
    // ── Section 1: every run/block domain in one flow ──
    {
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
        titlePage: true,
      },
      headers: {
        default: [
          {
            paragraph: {
              children: [
                "Header text",
                {
                  picture: { data: png1x1, type: "png", transformation: { width: 16, height: 16 } },
                },
              ],
            },
          },
        ],
        first: [{ paragraph: "First page header" }],
        even: [{ paragraph: "Even header" }],
      },
      footers: {
        default: [
          {
            paragraph: {
              children: [{ children: ["CURRENT"] }, " of ", { children: ["TOTAL_PAGES"] }],
            },
          },
        ],
      },
      children: [
        { paragraph: { style: "Lead", children: ["Maximum coverage document"] } },
        {
          paragraph: {
            children: [
              { text: "Bold ", bold: true },
              { text: "italic ", italic: true },
              { text: "highlight", highlight: "yellow" },
              { text: " shaded", shading: { fill: "D9E2F3" } },
              { text: " strike", strike: true },
              { text: " underline", underline: { type: "single" } },
              { text: " superscript", superScript: true },
              { style: "Emphasis", text: " charStyle" },
              { children: [{ tab: true }] },
              "after tab",
            ],
          },
        },
        {
          paragraph: {
            tabStops: [{ type: "right", position: 9026, leader: "dot" }],
            children: ["Left", { children: [{ tab: true }] }, "Right"],
          },
        },
        { paragraph: { children: [{ footnoteReference: 1 }, " main text"] } },
        { paragraph: { children: [{ endnoteReference: 1 }, " endnoted"] } },
        {
          paragraph: {
            children: [{ bookmark: { name: "targetAnchor", wrap: ["Anchored text."] } }],
          },
        },
        {
          paragraph: {
            children: [
              { hyperlink: { url: "https://example.com", children: ["External link"] } },
              " | ",
              { hyperlink: { anchor: "targetAnchor", children: ["Internal anchor"] } },
            ],
          },
        },
        {
          paragraph: {
            numbering: { reference: "max-numbering", level: 0 },
            children: ["Numbered item one"],
          },
        },
        {
          paragraph: {
            numbering: { reference: "max-numbering", level: 1 },
            children: ["Nested letter item"],
          },
        },
        { paragraph: { bullet: { level: 0 }, children: ["Bulleted item"] } },
        {
          paragraph: {
            children: [
              {
                insertion: {
                  id: 1,
                  author: "Max",
                  date: "2026-08-16T00:00:00Z",
                  children: ["Inserted revision text."],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                deletion: {
                  id: 2,
                  author: "Max",
                  date: "2026-08-16T00:00:00Z",
                  children: ["Deleted text"],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                moveTo: {
                  name: "moved",
                  author: "Max",
                  date: "2026-08-16T00:00:00Z",
                  wrap: ["Moved text."],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              { permStart: { id: 10, editor: "everyone" } },
              "Editable range.",
              { permEnd: 10 },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                sdt: {
                  properties: { alias: "demo-sdt", tag: "sdt" },
                  children: ["Structured content inside an SDT."],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "☐ Legacy checkbox: ",
              { formField: { name: "MaxCheck", checkBox: { checked: true, sizeAuto: true } } },
            ],
          },
        },
        { paragraph: { children: [{ pageBreak: true }] } },
        { paragraph: { children: ["After a page break."] } },
        {
          paragraph: {
            children: [
              {
                ruby: { base: "漢字", text: "かんじ", alignment: "center", fontSize: 9 },
              },
              " annotated",
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                math: {
                  children: [
                    "E=",
                    {
                      fraction: {
                        numerator: [{ superScript: { children: ["mc"], superScript: ["2"] } }],
                        denominator: ["1"],
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                picture: {
                  data: png1x1,
                  type: "png",
                  transformation: { width: 64, height: 64 },
                  altText: { name: "inline-png", description: "An inline picture" },
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                picture: {
                  data: png1x1,
                  type: "png",
                  transformation: { width: 96, height: 96 },
                  floating: {
                    horizontalPosition: { relative: "margin", offset: 200000 },
                    verticalPosition: { relative: "paragraph", offset: 100000 },
                    wrap: { type: 2, side: "bothSides" },
                    margins: { top: 10000, bottom: 10000, left: 10000, right: 10000 },
                  },
                },
              },
              "Text wrapping a floating picture.",
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  presetGeometry: { preset: "roundRect" },
                  fill: { type: "solid", color: "4472C4" },
                  transformation: { width: 200, height: 80 },
                  children: ["Shape text"],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                wpgGroup: {
                  transformation: { height: "3cm", width: "8cm" },
                  children: [
                    {
                      type: "wps",
                      data: {
                        children: ["Grouped ellipse caption"],
                        presetGeometry: { preset: "ellipse" },
                        fill: { type: "solid", color: "ED7D31" },
                      },
                      transformation: {
                        pixels: { x: 0, y: 0 },
                        emus: { x: 1_828_800, y: 1_082_480 },
                      },
                    },
                    {
                      type: "png",
                      data: png1x1,
                      fileName: "grouped.png",
                      transformation: {
                        pixels: { x: 0, y: 0 },
                        emus: { x: 914_400, y: 914_400 },
                        offset: { emus: { x: 1_905_000, y: 0 }, pixels: { x: 200, y: 0 } },
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "column",
                  categories: ["Q1", "Q2"],
                  series: [{ name: "Sales", values: [10, 20] }],
                  showLegend: true,
                  transformation: { width: 300, height: 200 },
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                smartArt: {
                  layout: "simple1",
                  data: { nodes: [{ text: "One" }, { text: "Two" }] },
                  transformation: { width: 300, height: 150 },
                },
              },
            ],
          },
        },
        {
          table: {
            columnWidths: [4500, 4500],
            rows: [
              {
                tableHeader: true,
                cells: [{ children: [{ paragraph: "H1" }] }, { children: [{ paragraph: "H2" }] }],
              },
              {
                cells: [
                  {
                    columnSpan: 2,
                    shading: { fill: "E2EFDA" },
                    children: [{ paragraph: "Merged across both columns" }],
                  },
                ],
              },
            ],
            width: { size: 100, type: "pct" },
            borders: {
              top: { style: "single", size: 4, color: "000000" },
              bottom: { style: "single", size: 4, color: "000000" },
              left: { style: "single", size: 4, color: "000000" },
              right: { style: "single", size: 4, color: "000000" },
            },
          },
        },
        {
          paragraph: {
            children: [
              {
                simpleField: { instruction: ' DATE \\@ "yyyy-MM-dd" ', cachedValue: "2026-08-16" },
              },
            ],
          },
        },
        { paragraph: { children: ["Seq: ", { seqIdentifier: "figure" }] } },
        { toc: { alias: "TOC", hyperlink: true } },
        { paragraph: { children: [{ columnBreak: true }, "Second column text"] } },
      ],
    },
    // ── Section 2: page-layout domains ──
    {
      properties: {
        type: "nextPage",
        page: {
          size: { orientation: "landscape", width: 15840, height: 12240 },
          pageNumbers: { start: 1 },
          borders: { top: { style: "double", size: 4, color: "2E74B5" } },
        },
        column: { count: 2, space: 360 },
      },
      children: [
        { paragraph: { children: ["Landscape two-column section with a page border."] } },
        { paragraph: { children: ["Line numbering is active in this section."] } },
      ],
    },
  ],
};

const buffer = await generateDocument(doc);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/135-max-coverage.docx", buffer);
console.log("written 135-max-coverage.docx");
