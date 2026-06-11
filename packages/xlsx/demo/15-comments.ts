import { writeFileSync } from "node:fs";

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
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
