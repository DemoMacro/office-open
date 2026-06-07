// Dialogsheet, Revision Log, QueryTable, and Metadata modules.

import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  worksheets: [
    {
      name: "Data",
      rows: [
        { cells: [{ value: "Product" }, { value: "Price" }] },
        { cells: [{ value: "Widget" }, { value: 9.99 }] },
        { cells: [{ value: "Gadget" }, { value: 19.99 }] },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
