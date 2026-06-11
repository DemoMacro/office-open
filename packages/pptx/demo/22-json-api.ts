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
            x: 80,
            y: 120,
            width: 720,
            height: 80,
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "JSON-Friendly API",
                      fontSize: 44,
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
            x: 180,
            y: 220,
            width: 520,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Pure objects, no class instantiation needed",
                      fontSize: 18,
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
            x1: 300,
            y1: 300,
            x2: 580,
            y2: 300,
            endArrowhead: "triangle",
            outline: { color: "FFC000", width: 2 },
          },
        },
      ],
    },

    // ── Slide 2: Table with rich formatting ──
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 50,
            textBody: { text: "Table (JSON API)" },
            fill: "4472C4",
          },
        },
        {
          table: {
            x: 50,
            y: 100,
            width: 700,
            height: 280,
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
            x: 50,
            y: 20,
            width: 300,
            height: 40,
            textBody: { text: "Lines & Connectors" },
            fill: "70AD47",
          },
        },
        { line: { x1: 50, y1: 80, x2: 800, y2: 80, outline: { color: "4472C4", width: 2 } } },
        { line: { x1: 50, y1: 100, x2: 400, y2: 300, outline: { color: "ED7D31", width: 3 } } },
        {
          connector: {
            x1: 450,
            y1: 100,
            x2: 800,
            y2: 300,
            endArrowhead: "stealth",
            outline: { color: "70AD47", width: 2 },
          },
        },
        {
          shape: {
            x: 50,
            y: 340,
            width: 350,
            height: 50,
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
            x: 50,
            y: 20,
            width: 400,
            height: 50,
            textBody: { text: "Group Shapes (JSON)" },
            fill: "7030A0",
          },
        },
        {
          group: {
            x: 50,
            y: 90,
            width: 350,
            height: 200,
            children: [
              {
                shape: {
                  x: 0,
                  y: 0,
                  width: 160,
                  height: 90,
                  textBody: { text: "A" },
                  fill: "4472C4",
                },
              },
              {
                shape: {
                  x: 190,
                  y: 0,
                  width: 160,
                  height: 90,
                  textBody: { text: "B" },
                  fill: "ED7D31",
                },
              },
              {
                shape: {
                  x: 0,
                  y: 110,
                  width: 350,
                  height: 90,
                  textBody: { text: "C (wide)" },
                  fill: "70AD47",
                },
              },
            ],
          },
        },
        {
          group: {
            x: 450,
            y: 90,
            width: 300,
            height: 200,
            rotation: 15,
            children: [
              { shape: { x: 0, y: 0, width: 300, height: 200, fill: "FFC000" } },
              {
                shape: {
                  x: 30,
                  y: 40,
                  width: 240,
                  height: 120,
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
            x: 50,
            y: 20,
            width: 600,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Rich Text (Plain Objects)", fontSize: 28, bold: true }],
                },
              ],
            },
          },
        },
        // Bold / Italic / Underline / Color
        {
          shape: {
            x: 50,
            y: 80,
            width: 700,
            height: 35,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "Bold", bold: true, fontSize: 18 },
                    { text: " | " },
                    { text: "Italic", italic: true, fontSize: 18 },
                    { text: " | " },
                    { text: "Underline", underline: "single", fontSize: 18 },
                    { text: " | " },
                    { text: "Red", fill: "FF0000", fontSize: 18 },
                    { text: " " },
                    { text: "Blue", fill: "4472C4", fontSize: 18 },
                  ],
                },
              ],
            },
          },
        },
        // Superscript / Subscript
        {
          shape: {
            x: 50,
            y: 130,
            width: 400,
            height: 35,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "E = mc", fontSize: 20 },
                    { text: "2", fontSize: 14, baseline: 30000 },
                    { text: "    H", fontSize: 20 },
                    { text: "2", fontSize: 14, baseline: -25000 },
                    { text: "O", fontSize: 20 },
                  ],
                },
              ],
            },
          },
        },
        // Bullet list (plain-object paragraphs)
        {
          shape: {
            x: 50,
            y: 180,
            width: 300,
            height: 150,
            outline: { color: "4472C4", width: 1 },
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
            x: 380,
            y: 180,
            width: 300,
            height: 150,
            outline: { color: "ED7D31", width: 1 },
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
            x: 50,
            y: 360,
            width: 300,
            height: 100,
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
            x: 380,
            y: 360,
            width: 300,
            height: 100,
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
