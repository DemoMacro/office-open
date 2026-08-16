import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "With Comments",
      rows: [
        { cells: [{ value: "Product" }, { value: "Price" }, { value: "Stock" }] },
        { cells: [{ value: "Widget A" }, { value: 9.99 }, { value: 150 }] },
        { cells: [{ value: "Widget B" }, { value: 14.99 }, { value: 75 }] },
        { cells: [{ value: "Widget C" }, { value: 24.99 }, { value: 0 }] },
      ],
      comments: [
        { cell: "B2", author: "Alice", text: "Discounted from 12.99" },
        { cell: "C3", author: "Bob", text: "Reorder soon — running low" },
        { cell: "C4", author: "Alice", text: "Out of stock! Contact supplier." },
        {
          // Custom placement — anchored at D5, widened to 160×90 pt, pinned
          // visible instead of Excel's hover-reveal default.
          cell: "D5",
          author: "Alice",
          text: "Pinned note with a custom anchor",
          anchor: {
            from: { col: 3, colOff: 285750, row: 4, rowOff: 228600 },
            to: { col: 6, colOff: 571500, row: 9, rowOff: 114300 },
          },
          visible: true,
          size: { width: 160, height: 90 },
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/15-comments.xlsx", buffer);
