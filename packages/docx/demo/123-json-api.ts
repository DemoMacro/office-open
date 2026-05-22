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
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
