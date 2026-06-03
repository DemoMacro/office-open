import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
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
                r: 1,
                cells: [
                  { r: "A1", t: "s", v: "Product" },
                  { r: "B1", t: "n", v: "100" },
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
      rows: [{ cells: [{ value: "External Data Reference" }, { value: "Source: data.xlsx" }] }],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
