import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
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
        customFilters: [{ colId: 1, operator: "greaterThan", val: "200", hiddenButton: true }],
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
        top10: [{ colId: 1, val: 3, filterVal: 65 }],
      },
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
