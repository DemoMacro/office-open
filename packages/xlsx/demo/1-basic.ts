import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
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

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
