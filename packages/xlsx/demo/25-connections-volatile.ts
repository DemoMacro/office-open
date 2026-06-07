// Volatile dependencies and web publish objects.

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
  volTypes: [
    {
      type: "realTimeData",
      mains: [
        {
          first: "rId1",
          topics: [
            {
              value: "StockPrice",
              stringTopics: ["Topic1"],
              refs: [{ reference: "A1", sheetIndex: 0 }],
            },
          ],
        },
      ],
    },
  ],
  webPublishObjects: [
    {
      rId: "rId1",
      destinationFile: "report.htm",
      title: "Price List",
      autoRepublish: true,
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
