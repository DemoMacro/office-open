import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  externalLinks: [
    {
      externalBook: {
        target: "data.xlsx",
        sheetNames: ["Sheet1", "Sheet2"],
        definedNames: [
          { name: "SalesData", refersTo: "Sheet1!$A$1:$D$10" },
          { name: "TotalRevenue", refersTo: "Sheet1!$E$1" },
        ],
        sheetDataSet: [
          {
            sheetId: 0,
            rows: [
              {
                rowNumber: 1,
                cells: [
                  { reference: "A1", type: "s", value: "Product" },
                  { reference: "B1", type: "n", value: "100" },
                ],
              },
            ],
          },
        ],
      },
    },
  ],
  worksheets: [
    {
      name: "Summary",
      rows: [
        {
          cells: [{ value: "External Data from data.xlsx" }],
        },
        {
          cells: [
            { value: "Product:" },
            { formula: { formula: "'[data.xlsx]Sheet1'!A1" }, value: "Product" },
          ],
        },
        {
          cells: [
            { value: "Value:" },
            { formula: { formula: "'[data.xlsx]Sheet1'!B1" }, value: 100 },
          ],
        },
        {
          cells: [
            { value: "Total Revenue:" },
            { formula: { formula: "data.xlsx!TotalRevenue" }, value: 0 },
          ],
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
