import * as fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const options: PresentationOptions = {
  title: "My Presentation",
  creator: "Demo",
  show: {
    type: "present",
    loop: false,
    penColor: "FF0000",
  },
  view: {
    lastView: "slideSorterView",
    gridSpacing: { cx: 50800, cy: 50800 },
  },
  slides: [
    // Slide 1: Basic shapes
    {
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 400,
            height: 200,
            textBody: { text: "Hello World" },
            geometry: "rect",
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: 200,
            y: 350,
            width: 500,
            height: 100,
            textBody: { text: "Second shape" },
          },
        },
      ],
    },
    // Slide 2: Full width
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 860,
            height: 440,
            textBody: { text: "Slide 2 - Full Width" },
            geometry: "rect",
          },
        },
      ],
    },
    // Slide 3: Vertical text
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 600,
            height: 50,
            textBody: { text: "Vertical Text" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: 50,
            y: 120,
            width: 120,
            height: 300,
            textBody: {
              vertical: "vert",
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Vertical Text (top to bottom)",
                      fontSize: 14,
                    },
                  ],
                },
              ],
            },
            outline: { color: "4472C4", width: 1 },
          },
        },
        {
          shape: {
            x: 200,
            y: 120,
            width: 120,
            height: 300,
            textBody: {
              vertical: "vert270",
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: "Rotated 270 (bottom to top)", fontSize: 14 }],
                },
              ],
            },
            outline: { color: "ED7D31", width: 1 },
          },
        },
        {
          shape: {
            x: 350,
            y: 120,
            width: 120,
            height: 300,
            textBody: {
              vertical: "horz",
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: "Horizontal (default)", fontSize: 14 }],
                },
              ],
            },
            outline: { color: "70AD47", width: 1 },
          },
        },
      ],
    },
    // Slide 4: Text anchor & auto-fit
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 600,
            height: 50,
            textBody: { text: "Text Anchor & Auto-Fit" },
            fill: "4472C4",
          },
        },
        // Top anchor
        {
          shape: {
            x: 50,
            y: 120,
            width: 200,
            height: 200,
            textBody: { anchor: "top", text: "Top anchored text" },
            outline: { color: "999999", width: 1 },
          },
        },
        // Center anchor
        {
          shape: {
            x: 280,
            y: 120,
            width: 200,
            height: 200,
            textBody: { anchor: "center", text: "Center anchored text" },
            outline: { color: "999999", width: 1 },
          },
        },
        // Bottom anchor
        {
          shape: {
            x: 510,
            y: 120,
            width: 200,
            height: 200,
            textBody: { anchor: "bottom", text: "Bottom anchored text" },
            outline: { color: "999999", width: 1 },
          },
        },
        // Auto-fit normal
        {
          shape: {
            x: 50,
            y: 350,
            width: 250,
            height: 80,
            textBody: {
              autoFit: "normal",
              text: "This is a very long text that should auto-fit to shrink within the shape bounds",
            },
            outline: { color: "4472C4", width: 1 },
          },
        },
        // Auto-fit shape
        {
          shape: {
            x: 330,
            y: 350,
            width: 250,
            height: 80,
            textBody: { autoFit: "shape", text: "Shape auto-fit text" },
            outline: { color: "ED7D31", width: 1 },
          },
        },
      ],
    },
    // Slide 5: Text margins & columns
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 600,
            height: 50,
            textBody: { text: "Text Margins & Columns" },
            fill: "4472C4",
          },
        },
        // Default margins
        {
          shape: {
            x: 50,
            y: 120,
            width: 350,
            height: 150,
            outline: { color: "999999", width: 1 },
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Default margins (no extra padding)",
                      fontSize: 12,
                    },
                  ],
                },
              ],
            },
          },
        },
        // Wide margins
        {
          shape: {
            x: 430,
            y: 120,
            width: 350,
            height: 150,
            textBody: {
              margins: { top: 100000, bottom: 100000, left: 200000, right: 200000 },
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Wide margins (extra padding all around)",
                      fontSize: 12,
                    },
                  ],
                },
              ],
            },
            outline: { color: "ED7D31", width: 1 },
          },
        },
        // 2 columns
        {
          shape: {
            x: 50,
            y: 300,
            width: 730,
            height: 150,
            textBody: {
              columns: 2,
              columnSpacing: 12,
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "This is column 1 text. The shape is divided into 2 columns with spacing between them.",
                      fontSize: 12,
                    },
                  ],
                },
              ],
            },
            outline: { color: "70AD47", width: 1 },
          },
        },
      ],
    },
    // Slide 6: Blip fill (image) + gradient path (radial)
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 600,
            height: 50,
            textBody: { text: "Blip Fill & Gradient Path" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: 50,
            y: 120,
            width: 400,
            height: 200,
            textBody: { text: "Image Fill" },
            fill: {
              type: "blip",
              data: new Uint8Array(
                fs.readFileSync(path.resolve(__dirname, "assets/test-poster.png")),
              ),
              imageType: "png",
            },
          },
        },
        {
          shape: {
            x: 480,
            y: 120,
            width: 300,
            height: 200,
            textBody: { text: "Radial Gradient" },
            fill: {
              type: "gradient",
              path: "circle",
              stops: [
                { position: 0, color: "FFFFFF" },
                { position: 100, color: "4472C4" },
              ],
            },
          },
        },
      ],
    },
    // Slide 7: Shape locking
    {
      children: [
        {
          shape: {
            x: 50,
            y: 50,
            width: 600,
            height: 50,
            textBody: { text: "Shape Locking" },
            fill: "4472C4",
          },
        },
        // Locked shape: cannot select, move, resize, or edit text
        {
          shape: {
            x: 50,
            y: 120,
            width: 350,
            height: 100,
            textBody: { text: "Locked: no select, move, resize, text edit" },
            outline: { color: "ED7D31", width: 1 },
            locking: {
              noSelect: true,
              noMove: true,
              noResize: true,
              noTextEdit: true,
            },
          },
        },
        // Aspect-locked shape: can move/resize but ratio is fixed
        {
          shape: {
            x: 430,
            y: 120,
            width: 350,
            height: 100,
            textBody: { text: "Aspect locked: ratio is fixed" },
            outline: { color: "70AD47", width: 1 },
            locking: {
              noChangeAspect: true,
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
