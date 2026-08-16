// Rich text in cells and comments, comment properties with anchors.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
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
          // commentPr is parsed but never re-emitted: Excel rejects a commentPr
          // alongside the sheet's legacy VML note drawing (the two are rival
          // property systems for the same note; Excel reads note properties
          // from the VML shape's x:ClientData).
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

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/24-rich-text-comments.xlsx", buffer);
