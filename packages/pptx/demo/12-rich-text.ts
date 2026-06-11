import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Rich Text Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        // Title
        {
          shape: {
            x: 50,
            y: 30,
            width: 600,
            height: 50,
            textBody: {
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Rich Text Formatting",
                      fontSize: 36,
                      bold: true,
                      font: "Calibri",
                    },
                  ],
                },
              ],
            },
          },
        },
        // Bold + Italic + Underline
        {
          shape: {
            x: 50,
            y: 100,
            width: 600,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "Bold", bold: true, fontSize: 18 },
                    { text: " | " },
                    { text: "Italic", italic: true, fontSize: 18 },
                    { text: " | " },
                    {
                      text: "Underline",
                      underline: "single",
                      fontSize: 18,
                    },
                    { text: " | " },
                    {
                      text: "Bold+Italic+Underline",
                      bold: true,
                      italic: true,
                      underline: "double",
                      fontSize: 18,
                    },
                  ],
                },
              ],
            },
          },
        },
        // Strikethrough
        {
          shape: {
            x: 50,
            y: 150,
            width: 600,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Single Strike",
                      strike: "sngStrike",
                      fontSize: 18,
                    },
                    { text: " | " },
                    {
                      text: "Double Strike",
                      strike: "dblStrike",
                      fontSize: 18,
                    },
                  ],
                },
              ],
            },
          },
        },
        // Superscript + Subscript
        {
          shape: {
            x: 50,
            y: 200,
            width: 600,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "E = mc", fontSize: 20 },
                    { text: "2", fontSize: 14, baseline: 30000 },
                    { text: "    H", fontSize: 20 },
                    { text: "2", fontSize: 14, baseline: -25000 },
                    { text: "O    x", fontSize: 20 },
                    { text: "n+1", fontSize: 14, baseline: 30000 },
                  ],
                },
              ],
            },
          },
        },
        // Character spacing
        {
          shape: {
            x: 50,
            y: 250,
            width: 600,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "Normal spacing", fontSize: 18 },
                    {
                      text: "    Wide spacing",
                      fontSize: 18,
                      spacing: 400,
                    },
                    {
                      text: "    Tight spacing",
                      fontSize: 18,
                      spacing: -100,
                    },
                  ],
                },
              ],
            },
          },
        },
        // Capitalization
        {
          shape: {
            x: 50,
            y: 300,
            width: 600,
            height: 40,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    { text: "Normal text", fontSize: 18 },
                    {
                      text: "    ALL CAPS",
                      fontSize: 18,
                      capitalization: "all",
                    },
                    {
                      text: "    Small Caps",
                      fontSize: 18,
                      capitalization: "small",
                    },
                  ],
                },
              ],
            },
          },
        },
        // Text color
        {
          shape: {
            x: 50,
            y: 350,
            width: 600,
            height: 40,
            fill: "333333",
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Red",
                      fontSize: 20,
                      fill: "FF0000",
                    },
                    { text: " | " },
                    {
                      text: "Green",
                      fontSize: 20,
                      fill: "00FF00",
                    },
                    { text: " | " },
                    {
                      text: "Blue",
                      fontSize: 20,
                      fill: "4472C4",
                    },
                    { text: " | " },
                    {
                      text: "Yellow",
                      fontSize: 20,
                      fill: "FFC000",
                    },
                  ],
                },
              ],
            },
          },
        },
        // Alignment demo
        {
          shape: {
            x: 50,
            y: 420,
            width: 600,
            height: 120,
            outline: { color: "999999", width: 1 },
            textBody: {
              children: [
                {
                  properties: { alignment: "left", bullet: { type: "none" } },
                  children: [{ text: "Left aligned text", fontSize: 16 }],
                },
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Center aligned text", fontSize: 16 }],
                },
                {
                  properties: { alignment: "right", bullet: { type: "none" } },
                  children: [{ text: "Right aligned text", fontSize: 16 }],
                },
                {
                  properties: { alignment: "justify", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Justified text that is long enough to wrap to multiple lines to demonstrate the justification effect clearly",
                      fontSize: 16,
                    },
                  ],
                },
              ],
            },
          },
        },
        // Bullet & numbering
        {
          shape: {
            x: 50,
            y: 710,
            width: 280,
            height: 200,
            outline: { color: "4472C4", width: 1 },
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "char", char: "\u25CF" } },
                  children: [{ text: "Bullet Point 1" }],
                },
                {
                  properties: { bullet: { type: "char", char: "\u25CF", color: "4472C4" } },
                  children: [{ text: "Blue Bullet 2" }],
                },
                {
                  properties: { bullet: { type: "char", char: "\u25A0", size: 75 } },
                  children: [{ text: "Small Square 3" }],
                },
                {
                  properties: { bullet: { type: "char", char: "\u27A2" } },
                  children: [{ text: "Arrow Bullet 4" }],
                },
              ],
            },
          },
        },
        {
          shape: {
            x: 360,
            y: 710,
            width: 280,
            height: 200,
            outline: { color: "ED7D31", width: 1 },
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                  children: [{ text: "First item" }],
                },
                {
                  properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                  children: [{ text: "Second item" }],
                },
                {
                  properties: { bullet: { type: "autoNum", format: "arabicPeriod" } },
                  children: [{ text: "Third item" }],
                },
                {
                  properties: { bullet: { type: "autoNum", format: "alphaLcParenBoth" } },
                  children: [{ text: "Sub-item a" }],
                },
                {
                  properties: { bullet: { type: "autoNum", format: "alphaLcParenBoth" } },
                  children: [{ text: "Sub-item b" }],
                },
              ],
            },
          },
        },
        // RTL, NoProof, Shadow, Outline
        {
          shape: {
            x: 660,
            y: 710,
            width: 280,
            height: 200,
            outline: { color: "70AD47", width: 1 },
            textBody: {
              children: [
                {
                  children: [{ text: "Normal text", fontSize: 16 }],
                },
                {
                  children: [
                    {
                      text: "Right-to-Left",
                      rightToLeft: true,
                      fontSize: 16,
                    },
                  ],
                },
                {
                  children: [{ text: "No proof text", noProof: true, fontSize: 16 }],
                },
                {
                  children: [{ text: "Shadow text", shadow: true, fontSize: 16 }],
                },
                {
                  children: [{ text: "Outline text", outline: true, fontSize: 16 }],
                },
              ],
            },
          },
        },
        // Line spacing demo
        {
          shape: {
            x: 50,
            y: 560,
            width: 280,
            height: 120,
            outline: { color: "4472C4", width: 1 },
            textBody: {
              children: [
                {
                  properties: { lineSpacing: 100, bullet: { type: "none" } },
                  children: [{ text: "Single spacing (1.0)", fontSize: 14 }],
                },
                {
                  properties: { lineSpacing: 150, bullet: { type: "none" } },
                  children: [{ text: "1.5x spacing", fontSize: 14 }],
                },
                {
                  properties: { lineSpacing: 200, bullet: { type: "none" } },
                  children: [{ text: "Double spacing (2.0)", fontSize: 14 }],
                },
              ],
            },
          },
        },
        // Fixed line spacing
        {
          shape: {
            x: 360,
            y: 560,
            width: 280,
            height: 120,
            outline: { color: "ED7D31", width: 1 },
            textBody: {
              children: [
                {
                  properties: { lineSpacingPoints: 14, bullet: { type: "none" } },
                  children: [{ text: "Exactly 14pt", fontSize: 14 }],
                },
                {
                  properties: { lineSpacingPoints: 20, bullet: { type: "none" } },
                  children: [{ text: "Exactly 20pt", fontSize: 14 }],
                },
                {
                  properties: { lineSpacingPoints: 28, bullet: { type: "none" } },
                  children: [{ text: "Exactly 28pt", fontSize: 14 }],
                },
              ],
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
