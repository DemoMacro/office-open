import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Sheet1",
      rows: [
        {
          cells: [{ value: "Name" }, { value: "Age" }, { value: "City" }],
        },
        {
          cells: [{ value: "Alice" }, { value: 30 }, { value: "New York" }],
        },
        {
          cells: [{ value: "Bob" }, { value: 25 }, { value: "London" }],
        },
        {
          cells: [{ value: "Charlie" }, { value: true }, { value: "Paris" }],
        },
      ],
    },
    {
      name: "Sheet2",
      rows: [
        {
          cells: [{ value: "Hello" }, { value: 42 }],
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
