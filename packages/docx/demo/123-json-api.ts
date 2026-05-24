// JSON-friendly API demo — pure-object syntax, no `new` for content elements

import * as fs from "fs";

import { Document, Packer } from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      // Headers/footers: plain SectionChild arrays
      headers: {
        default: [
          { paragraph: { children: [{ text: "Header", bold: true }] } },
          { paragraph: { children: ["Second header line"] } },
        ],
      },
      footers: {
        default: [
          { paragraph: { children: ["Footer — page ", { text: "plain object", italics: true }] } },
        ],
      },
      children: [
        // ─── Paragraph shorthand: string ───────────────────
        { paragraph: "I am a string paragraph" },

        // ─── Paragraph children: string + IRunOptions mix ──
        {
          paragraph: {
            children: [
              "String child",
              { text: " + bold child", bold: true },
              { text: " + italic", italics: true },
            ],
          },
        },

        // ─── Paragraph with alignment ──────────────────────
        {
          paragraph: {
            alignment: "center",
            children: [{ text: "Centered text", bold: true, size: 32 }],
          },
        },

        // ─── Table: rows and cells as plain objects ────────
        {
          table: {
            rows: [
              {
                children: [
                  { children: [{ paragraph: "Row1 Cell1" }] },
                  { children: [{ paragraph: "Row1 Cell2" }] },
                  { children: [{ paragraph: "Row1 Cell3" }] },
                ],
              },
              {
                children: [
                  {
                    children: [
                      { paragraph: { children: [{ text: "Underlined cell", underline: {} }] } },
                    ],
                  },
                  {
                    children: [
                      { paragraph: "Cell with" },
                      {
                        paragraph: { children: ["multiple", { text: " paragraphs", bold: true }] },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        },

        // ─── Paragraph with mixed children types ───────────
        {
          paragraph: {
            children: [
              { text: "IRunOptions child", color: "FF0000" },
              " + string child",
              { text: " + another IRunOptions", bold: true },
            ],
          },
        },

        // ─── Table of Contents (plain object) ──────────────
        { toc: { hyperlink: true, headingStyleRange: "1-3" } },

        // ─── Headings ──────────────────────────────────────
        { paragraph: { heading: "Heading1", children: ["First Heading"] } },
        { paragraph: { children: ["Content under first heading."] } },
        { paragraph: { heading: "Heading2", children: ["Second Heading"] } },
        { paragraph: { children: ["Content under second heading."] } },
      ],
    },

    // ─── Section 2: different header/footer slots ─────────────────
    {
      headers: {
        first: [{ paragraph: { children: [{ text: "First page only header", bold: true }] } }],
        default: [{ paragraph: { children: ["Default header for section 2"] } }],
      },
      footers: {
        default: [{ paragraph: { children: ["Section 2 footer"] } }],
      },
      properties: {
        titlePage: true,
      },
      children: [{ paragraph: { children: ["Section 2 — first page has its own header"] } }],
    },

    // ─── Section 3: nested table (table inside table cell) ────────
    {
      children: [
        { paragraph: { children: ["Section 3 — nested table demo"] } },
        {
          table: {
            rows: [
              {
                children: [
                  {
                    children: [
                      { paragraph: "Outer cell" },
                      {
                        table: {
                          rows: [
                            {
                              children: [
                                { children: [{ paragraph: "Inner A" }] },
                                { children: [{ paragraph: "Inner B" }] },
                              ],
                            },
                          ],
                        },
                      },
                    ],
                  },
                  { children: [{ paragraph: "Sibling cell" }] },
                ],
              },
            ],
          },
        },
      ],
    },

    // ─── Section 4: textbox with SectionChild children ────────────
    {
      children: [
        { paragraph: { children: ["Section 4 — textbox with JSON children"] } },
        {
          textbox: {
            style: { width: "4in", height: "1in" },
            children: [
              { paragraph: { children: [{ text: "Textbox paragraph", bold: true }] } },
              { paragraph: { children: ["Second line in textbox"] } },
            ],
          },
        },
      ],
    },

    // ─── Section 5: SDT block with SectionChild children ──────────
    {
      children: [
        { paragraph: { children: ["Section 5 — SDT block with JSON children"] } },
        {
          sdt: {
            properties: { richText: true, alias: "BlockContent", tag: "block-content" },
            children: [
              { paragraph: { children: ["This is a block-level content control."] } },
              { paragraph: { children: ["It can contain multiple paragraphs and tables."] } },
            ],
          },
        },
        {
          sdt: {
            properties: {
              alias: "Color",
              tag: "combo-color",
              comboBox: {
                items: [
                  { displayText: "Red", value: "red" },
                  { displayText: "Blue", value: "blue" },
                ],
                lastValue: "Red",
              },
            },
            children: [{ paragraph: { children: ["Red"] } }],
          },
        },
      ],
    },

    // ─── Section 6: altChunk (embedded HTML/plain text) ────────────
    {
      children: [
        { paragraph: { children: ["Section 6 — altChunk with JSON API"] } },
        {
          altChunk: {
            data: "<html><body><p>This is <b>embedded HTML</b> via JSON API.</p></body></html>",
            contentType: "text/html",
            extension: "html",
          },
        },
        {
          altChunk: {
            data: "Plain text chunk inserted via JSON API.",
            contentType: "text/plain",
            extension: "txt",
          },
        },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
