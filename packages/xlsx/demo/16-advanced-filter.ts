import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Advanced Filter",
      rows: [
        { cells: [{ value: "Product" }, { value: "Sales" }, { value: "Region" }] },
        { cells: [{ value: "Widget A" }, { value: 500 }, { value: "North" }] },
        { cells: [{ value: "Widget B" }, { value: 200 }, { value: "South" }] },
        { cells: [{ value: "Widget C" }, { value: 800 }, { value: "East" }] },
        { cells: [{ value: "Widget D" }, { value: 300 }, { value: "West" }] },
        { cells: [{ value: "Widget E" }, { value: 100 }, { value: "North" }] },
      ],
      autoFilter: {
        ref: "A1:C6",
        columns: [
          {
            colId: 1,
            hiddenButton: true,
            customFilters: { entries: [{ operator: "greaterThan", val: "200" }] },
          },
        ],
        sort: [{ ref: "B1", descending: true }],
      },
    },
    {
      name: "Top 10 Filter",
      rows: [
        { cells: [{ value: "Student" }, { value: "Score" }] },
        { cells: [{ value: "Alice" }, { value: 95 }] },
        { cells: [{ value: "Bob" }, { value: 82 }] },
        { cells: [{ value: "Charlie" }, { value: 78 }] },
        { cells: [{ value: "Diana" }, { value: 91 }] },
        { cells: [{ value: "Eve" }, { value: 65 }] },
      ],
      autoFilter: {
        ref: "A1:B6",
        columns: [{ colId: 1, top10: { val: 3, filterVal: 65 } }],
      },
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/16-advanced-filter.xlsx", buffer);
