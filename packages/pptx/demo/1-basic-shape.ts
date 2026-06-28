import * as fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

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
            x: "2cm",
            y: "2cm",
            width: "8cm",
            height: "4cm",
            textBody: { text: "Hello World" },
            geometry: "rect",
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "4cm",
            y: "7cm",
            width: "10cm",
            height: "2cm",
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
            x: "1cm",
            y: "1cm",
            width: "17cm",
            height: "9cm",
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
            x: "1cm",
            y: "1cm",
            width: "12cm",
            height: "1cm",
            textBody: { text: "Vertical Text" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "3.2cm",
            height: "7.9cm",
            textBody: {
              vertical: "vert",
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Vertical Text (top to bottom)",
                      size: 14,
                    },
                  ],
                },
              ],
            },
            outline: { color: "4472C4", width: "1pt" },
          },
        },
        {
          shape: {
            x: "5.3cm",
            y: "3.2cm",
            width: "3.2cm",
            height: "7.9cm",
            textBody: {
              vertical: "vert270",
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: "Rotated 270 (bottom to top)", size: 14 }],
                },
              ],
            },
            outline: { color: "ED7D31", width: "1pt" },
          },
        },
        {
          shape: {
            x: "9.3cm",
            y: "3.2cm",
            width: "3.2cm",
            height: "7.9cm",
            textBody: {
              vertical: "horz",
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: "Horizontal (default)", size: 14 }],
                },
              ],
            },
            outline: { color: "70AD47", width: "1pt" },
          },
        },
      ],
    },
    // Slide 4: Text anchor & auto-fit
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "1.3cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Text Anchor & Auto-Fit" },
            fill: "4472C4",
          },
        },
        // Top anchor
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "5.3cm",
            textBody: { anchor: "top", text: "Top anchored text" },
            outline: { color: "999999", width: "1pt" },
          },
        },
        // Center anchor
        {
          shape: {
            x: "7.4cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "5.3cm",
            textBody: { anchor: "center", text: "Center anchored text" },
            outline: { color: "999999", width: "1pt" },
          },
        },
        // Bottom anchor
        {
          shape: {
            x: "13.5cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "5.3cm",
            textBody: { anchor: "bottom", text: "Bottom anchored text" },
            outline: { color: "999999", width: "1pt" },
          },
        },
        // Auto-fit normal
        {
          shape: {
            x: "1.3cm",
            y: "9.3cm",
            width: "6.6cm",
            height: "2.1cm",
            textBody: {
              autoFit: "normal",
              text: "This is a very long text that should auto-fit to shrink within the shape bounds",
            },
            outline: { color: "4472C4", width: "1pt" },
          },
        },
        // Auto-fit shape
        {
          shape: {
            x: "8.7cm",
            y: "9.3cm",
            width: "6.6cm",
            height: "2.1cm",
            textBody: { autoFit: "shape", text: "Shape auto-fit text" },
            outline: { color: "ED7D31", width: "1pt" },
          },
        },
      ],
    },
    // Slide 5: Text margins & columns
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "1.3cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Text Margins & Columns" },
            fill: "4472C4",
          },
        },
        // Default margins
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "9.3cm",
            height: "4.0cm",
            outline: { color: "999999", width: "1pt" },
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Default margins (no extra padding)",
                      size: 12,
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
            x: "11.4cm",
            y: "3.2cm",
            width: "9.3cm",
            height: "4.0cm",
            textBody: {
              margins: { top: 100000, bottom: 100000, left: 200000, right: 200000 },
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "Wide margins (extra padding all around)",
                      size: 12,
                    },
                  ],
                },
              ],
            },
            outline: { color: "ED7D31", width: "1pt" },
          },
        },
        // 2 columns
        {
          shape: {
            x: "1.3cm",
            y: "7.9cm",
            width: "19.3cm",
            height: "4.0cm",
            textBody: {
              columns: 2,
              columnSpacing: 12,
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: "This is column 1 text. The shape is divided into 2 columns with spacing between them.",
                      size: 12,
                    },
                  ],
                },
              ],
            },
            outline: { color: "70AD47", width: "1pt" },
          },
        },
      ],
    },
    // Slide 6: Blip fill (image) + gradient path (radial)
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "1.3cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Blip Fill & Gradient Path" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "10.6cm",
            height: "5.3cm",
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
            x: "12.7cm",
            y: "3.2cm",
            width: "7.9cm",
            height: "5.3cm",
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
            x: "1.3cm",
            y: "1.3cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Shape Locking" },
            fill: "4472C4",
          },
        },
        // Locked shape: cannot select, move, resize, or edit text
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "9.3cm",
            height: "2.6cm",
            textBody: { text: "Locked: no select, move, resize, text edit" },
            outline: { color: "ED7D31", width: "1pt" },
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
            x: "11.4cm",
            y: "3.2cm",
            width: "9.3cm",
            height: "2.6cm",
            textBody: { text: "Aspect locked: ratio is fixed" },
            outline: { color: "70AD47", width: "1pt" },
            locking: {
              noChangeAspect: true,
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
