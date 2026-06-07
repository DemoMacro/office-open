import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  worksheets: [
    {
      name: "Sales",
      rows: [
        {
          cells: [
            { value: "Product" },
            { value: "Amount" },
            { value: "Region" },
            { value: "Quarter" },
          ],
        },
        { cells: [{ value: "Widget A" }, { value: 1500 }, { value: "North" }, { value: "Q1" }] },
        { cells: [{ value: "Widget B" }, { value: 2300 }, { value: "South" }, { value: "Q1" }] },
        { cells: [{ value: "Widget C" }, { value: 890 }, { value: "East" }, { value: "Q1" }] },
        { cells: [{ value: "Widget A" }, { value: 1750 }, { value: "West" }, { value: "Q2" }] },
        { cells: [{ value: "Widget B" }, { value: 2100 }, { value: "North" }, { value: "Q2" }] },
        { cells: [{ value: "Widget C" }, { value: 950 }, { value: "South" }, { value: "Q2" }] },
        { cells: [{ value: "Widget A" }, { value: 1200 }, { value: "East" }, { value: "Q3" }] },
        { cells: [{ value: "Widget B" }, { value: 2800 }, { value: "West" }, { value: "Q3" }] },
        { cells: [{ value: "Widget C" }, { value: 1100 }, { value: "North" }, { value: "Q3" }] },
        // Totals row
        { cells: [{ value: "Total" }, { value: 14590 }, { value: "" }, { value: "" }] },
      ],
      tables: [
        {
          id: 1,
          displayName: "SalesTable",
          name: "SalesTable",
          ref: "A1:D11",
          autoFilter: "A1:D10",
          columns: [
            { name: "Product" },
            { name: "Amount", totalsRowFunction: "sum" },
            { name: "Region" },
            { name: "Quarter" },
          ],
          totalsRowCount: 1,
          style: {
            name: "TableStyleMedium9",
            showRowStripes: true,
          },
        },
      ],
    },
    {
      name: "Simple",
      rows: [
        { cells: [{ value: "Name" }, { value: "Score" }] },
        { cells: [{ value: "Alice" }, { value: 92 }] },
        { cells: [{ value: "Bob" }, { value: 78 }] },
        { cells: [{ value: "Charlie" }, { value: 85 }] },
      ],
      tables: [
        {
          id: 2,
          displayName: "ScoreTable",
          ref: "A1:B4",
          columns: [{ name: "Name" }, { name: "Score" }],
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
