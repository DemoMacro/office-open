// Workbook connections with query tables and volatile (RTD) dependencies.

import { mkdirSync, writeFileSync } from "node:fs";

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
  // volType/main/@first is a relationship id that must point at the workbook's
  // connections relationship — rId5 is what the compiler allocates for
  // xl/connections.xml in this workbook (sheets rId1-4 precede it).
  volTypes: [
    {
      type: "realTimeData",
      mains: [
        {
          first: "rId5",
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
      // type 6 = text file source; Excel rejects a dbPr (OLE DB) connection
      // without a live provider, so the demo queries a CSV-style text source.
      type: 6,
      refreshedVersion: 8,
      text: {
        prompt: false,
        sourceFile: "prices.csv",
        textFields: [{}],
      },
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/25-connections-volatile.xlsx", buffer);
