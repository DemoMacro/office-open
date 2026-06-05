import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Freeze & Filter",
      freezePanes: { row: 1 },
      autoFilter: {
        ref: "A1:D5",
        // Simple value filters with calendar type
        filters: [{ colId: 2, values: ["NYC", "London"], calendarType: "gregorian" }],
      },
      rows: [
        {
          cells: [
            { value: "Name", style: { font: { bold: true } } },
            { value: "Age", style: { font: { bold: true } } },
            { value: "City", style: { font: { bold: true } } },
            { value: "Score", style: { font: { bold: true } } },
          ],
        },
        { cells: [{ value: "Alice" }, { value: 30 }, { value: "NYC" }, { value: 95 }] },
        { cells: [{ value: "Bob" }, { value: 25 }, { value: "London" }, { value: 87 }] },
        { cells: [{ value: "Charlie" }, { value: 28 }, { value: "Paris" }, { value: 72 }] },
        { cells: [{ value: "Diana" }, { value: 32 }, { value: "Tokyo" }, { value: 91 }] },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
