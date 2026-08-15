// Volatile dependencies, web publish objects, and workbook connections with query tables.

import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Data",
      queryTables: [
        {
          name: "PriceQuery",
          fillFormulas: true,
          connectionId: 1,
          queryTableRefresh: {
            nextId: 3,
            queryTableFields: [
              { id: 1, name: "Product" },
              { id: 2, name: "Price" },
            ],
          },
        },
      ],
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
  connections: [
    {
      id: 1,
      name: "Price source",
      type: 2,
      refreshedVersion: 6,
      keepAlive: true,
      interval: 5,
      dbPr: {
        connection: "Provider=SQLOLEDB;Data Source=localhost",
        command: "SELECT Product, Price FROM Prices",
        commandType: 2,
      },
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
