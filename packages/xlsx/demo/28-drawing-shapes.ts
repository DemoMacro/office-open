import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

// Demonstrates anchored DrawingML objects on a worksheet: a shape with a text
// body, a connector line, and a group containing a nested shape. These reuse
// the shared core DrawingML descriptors (shapePropertiesDesc / textBodyDesc)
// also used by the docx and pptx packages.

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Shapes",
      rows: [{ cells: [{ value: "Anchored drawing objects" }] }],
      shapes: [
        {
          col: 2,
          row: 2,
          toCol: 5,
          toRow: 8,
          name: "Rectangle",
          spPr: { geometry: "rect", fill: { type: "solid", color: "4472C4" } },
          textBody: { paragraphs: [{ text: "Box" }] },
        },
      ],
      connectors: [{ col: 6, row: 4, toCol: 10, toRow: 4, spPr: { geometry: "line" } }],
      groups: [
        {
          col: 2,
          row: 10,
          toCol: 7,
          toRow: 16,
          name: "Group",
          grpSpPr: {
            x: 0,
            y: 0,
            width: 3000000,
            height: 1500000,
            childOffsetX: 0,
            childOffsetY: 0,
            childExtentWidth: 3000000,
            childExtentHeight: 1500000,
          },
          shapes: [
            {
              name: "Ellipse",
              spPr: { geometry: "ellipse", fill: { type: "solid", color: "ED7D31" } },
            },
          ],
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/28-drawing-shapes.xlsx", buffer);
