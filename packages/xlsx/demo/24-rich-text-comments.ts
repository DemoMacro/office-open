// Rich text in cells and comments, comment properties with anchors.

import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Rich Text",
      rows: [
        {
          cells: [{ value: "Product" }, { value: "Description" }],
        },
        {
          cells: [
            { value: "Widget A" },
            {
              value: {
                runs: [
                  { text: "High " },
                  { text: "quality", properties: { bold: true, color: "FF0000" } },
                  { text: " product" },
                ],
              },
            },
          ],
        },
        {
          cells: [
            { value: "Widget B" },
            {
              value: {
                runs: [
                  { text: "Affordable", properties: { italic: true, underline: "single" } },
                  { text: " & " },
                  { text: "reliable", properties: { bold: true, size: 14 } },
                ],
              },
            },
          ],
        },
      ],
      comments: [
        {
          cell: "B2",
          author: "Alice",
          text: {
            runs: [
              { text: "Note: ", properties: { bold: true } },
              { text: "Premium grade material" },
            ],
          },
          commentPr: {
            locked: false,
            autoFill: true,
            anchor: {
              moveWithCells: true,
              sizeWithCells: false,
            },
          },
        },
        {
          cell: "B3",
          author: "Bob",
          text: {
            runs: [
              { text: "Best seller! ", properties: { color: "008000", bold: true } },
              { text: "Consider bulk discount." },
            ],
          },
        },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
