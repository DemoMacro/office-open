import * as fs from "fs";

import { generatePresentation, type PresentationOptions } from "@office-open/pptx";

// This demo uses the new JSON-friendly API — no `new Shape()`, `new Paragraph()`, etc.
// All slide children are plain objects keyed by type ({ shape: {...} }, { table: {...} }, etc.).
// Paragraphs and text runs are also plain objects.

const pres: PresentationOptions = {
  title: "JSON API Demo",
  creator: "Demo",
  slides: [
    // ── Slide 1: Title slide with shape, background, and rich text ──
    {
      background: { fill: "1B2A4A" },
      children: [
        {
          shape: {
            x: "2.1cm",
            y: "3.2cm",
            width: "19.1cm",
            height: "2.1cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "JSON-Friendly API",
                      size: 44,
                      bold: true,
                      fill: "FFFFFF",
                      font: "Calibri",
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          shape: {
            x: "4.8cm",
            y: "5.8cm",
            width: "13.8cm",
            height: "1.1cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Pure objects, no class instantiation needed",
                      size: 18,
                      fill: "B0C4DE",
                    },
                  ],
                },
              ],
            },
          },
        },
        {
          connector: {
            x1: "7.9cm",
            y1: "7.9cm",
            x2: "15.3cm",
            y2: "7.9cm",
            endArrowhead: "triangle",
            outline: { color: "FFC000", width: "2pt" },
          },
        },
      ],
    },

    // ── Slide 2: Table with rich formatting ──
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.3cm",
            textBody: { text: "Table (JSON API)" },
            fill: "4472C4",
          },
        },
        {
          table: {
            x: "1.3cm",
            y: "2.6cm",
            width: "18.5cm",
            height: "7.4cm",
            rows: [
              {
                cells: [
                  { text: "Feature", fill: "4472C4" },
                  { text: "Old API", fill: "4472C4" },
                  { text: "New API", fill: "4472C4" },
                ],
              },
              {
                cells: [
                  { text: "Slide children" },
                  { text: "new Shape({...})" },
                  { text: "{ shape: {...} }" },
                ],
              },
              {
                cells: [
                  { text: "Paragraphs" },
                  { text: "new Paragraph({...})" },
                  { text: "{ children: [...] }" },
                ],
              },
              {
                cells: [
                  { text: "Text runs" },
                  { text: "new TextRun({...})" },
                  { text: "{ text, bold, ... }" },
                ],
              },
              {
                cells: [
                  { text: "Groups" },
                  { text: "new GroupShape({...})" },
                  { text: "{ group: {...} }" },
                ],
              },
              {
                cells: [
                  { text: "Table cells" },
                  { text: "new Paragraph({...})" },
                  { text: "{ children: [...] } in cell" },
                ],
              },
              {
                cells: [
                  { text: "Cell with children" },
                  {
                    children: [
                      {
                        children: [
                          { text: "Line 1, ", bold: true },
                          { text: "Line 1 cont.", fill: "FF0000" },
                        ],
                      },
                    ],
                  },
                  {
                    children: [
                      "String paragraph",
                      { children: [{ text: " + object paragraph", italic: true }] },
                    ],
                  },
                ],
              },
            ],
            columnWidths: [2000000, 2500000, 2500000],
            firstRow: true,
            bandRow: true,
          },
        },
      ],
      transition: { type: "push", speed: "med", direction: "left" },
    },

    // ── Slide 3: Lines, connectors, and mixed children ──
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "7.9cm",
            height: "1.1cm",
            textBody: { text: "Lines & Connectors" },
            fill: "70AD47",
          },
        },
        {
          line: {
            x1: "1.3cm",
            y1: "2.1cm",
            x2: "21.2cm",
            y2: "2.1cm",
            outline: { color: "4472C4", width: "2pt" },
          },
        },
        {
          line: {
            x1: "1.3cm",
            y1: "2.6cm",
            x2: "10.6cm",
            y2: "7.9cm",
            outline: { color: "ED7D31", width: "3pt" },
          },
        },
        {
          connector: {
            x1: "11.9cm",
            y1: "2.6cm",
            x2: "21.2cm",
            y2: "7.9cm",
            endArrowhead: "stealth",
            outline: { color: "70AD47", width: "2pt" },
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "9.0cm",
            width: "9.3cm",
            height: "1.3cm",
            textBody: { text: "\u2190 Line        Connector \u2192" },
            fill: "F2F2F2",
          },
        },
      ],
    },

    // ── Slide 4: Group shape (recursive JSON API) ──
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "10.6cm",
            height: "1.3cm",
            textBody: { text: "Group Shapes (JSON)" },
            fill: "7030A0",
          },
        },
        {
          group: {
            x: "1.3cm",
            y: "2.4cm",
            width: "9.3cm",
            height: "5.3cm",
            children: [
              {
                shape: {
                  x: "0.0cm",
                  y: "0.0cm",
                  width: "4.2cm",
                  height: "2.4cm",
                  textBody: { text: "A" },
                  fill: "4472C4",
                },
              },
              {
                shape: {
                  x: "5.0cm",
                  y: "0.0cm",
                  width: "4.2cm",
                  height: "2.4cm",
                  textBody: { text: "B" },
                  fill: "ED7D31",
                },
              },
              {
                shape: {
                  x: "0.0cm",
                  y: "2.9cm",
                  width: "9.3cm",
                  height: "2.4cm",
                  textBody: { text: "C (wide)" },
                  fill: "70AD47",
                },
              },
            ],
          },
        },
        {
          group: {
            x: "11.9cm",
            y: "2.4cm",
            width: "7.9cm",
            height: "5.3cm",
            rotation: 15,
            children: [
              {
                shape: { x: "0.0cm", y: "0.0cm", width: "7.9cm", height: "5.3cm", fill: "FFC000" },
              },
              {
                shape: {
                  x: "0.8cm",
                  y: "1.1cm",
                  width: "6.3cm",
                  height: "3.2cm",
                  textBody: { text: "Rotated Group" },
                  fill: "FFFFFF",
                },
              },
            ],
          },
        },
      ],
    },

    // ── Slide 5: Rich text with plain-object paragraphs & runs ──
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "15.9cm",
            height: "1.1cm",
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Rich Text (Plain Objects)", size: 28, bold: true }],
                },
              ],
            },
          },
        },
        // Bold / Italic / Underline / Color
        {
          shape: {
            x: "1.3cm",
            y: "2.1cm",
            width: "18.5cm",
            height: "0.9cm",
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "Bold", bold: true, size: 18 },
                    { text: " | " },
                    { text: "Italic", italic: true, size: 18 },
                    { text: " | " },
                    { text: "Underline", underline: "single", size: 18 },
                    { text: " | " },
                    { text: "Red", fill: "FF0000", size: 18 },
                    { text: " " },
                    { text: "Blue", fill: "4472C4", size: 18 },
                  ],
                },
              ],
            },
          },
        },
        // Superscript / Subscript
        {
          shape: {
            x: "1.3cm",
            y: "3.4cm",
            width: "10.6cm",
            height: "0.9cm",
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "E = mc", size: 20 },
                    { text: "2", size: 14, baseline: 30000 },
                    { text: "    H", size: 20 },
                    { text: "2", size: 14, baseline: -25000 },
                    { text: "O", size: 20 },
                  ],
                },
              ],
            },
          },
        },
        // Bullet list (plain-object paragraphs)
        {
          shape: {
            x: "1.3cm",
            y: "4.8cm",
            width: "7.9cm",
            height: "4.0cm",
            outline: { color: "4472C4", width: "1pt" },
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "char", char: "\u25CF" } },
                  children: [{ text: "First point" }],
                },
                {
                  properties: { bullet: { type: "char", char: "\u25CF", color: "4472C4" } },
                  children: [{ text: "Second point" }],
                },
                {
                  properties: { bullet: { type: "char", char: "\u25A0", size: 75 } },
                  children: [{ text: "Third point" }],
                },
              ],
            },
          },
        },
        // Numbered list
        {
          shape: {
            x: "10.1cm",
            y: "4.8cm",
            width: "7.9cm",
            height: "4.0cm",
            outline: { color: "ED7D31", width: "1pt" },
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                  children: [{ text: "Step one" }],
                },
                {
                  properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                  children: [{ text: "Step two" }],
                },
                {
                  properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                  children: [{ text: "Step three" }],
                },
              ],
            },
          },
        },
        // Gradient fill shape
        {
          shape: {
            x: "1.3cm",
            y: "9.5cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Gradient Fill" },
            fill: {
              type: "gradient",
              stops: [
                { position: 0, color: "4472C4" },
                { position: 100, color: "ED7D31" },
              ],
            },
          },
        },
        // Radial gradient
        {
          shape: {
            x: "10.1cm",
            y: "9.5cm",
            width: "7.9cm",
            height: "2.6cm",
            textBody: { text: "Radial Gradient" },
            fill: {
              type: "gradient",
              path: "circle",
              stops: [
                { position: 0, color: "FFFFFF" },
                { position: 100, color: "70AD47" },
              ],
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
