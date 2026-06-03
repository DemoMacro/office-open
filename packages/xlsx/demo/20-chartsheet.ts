import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        {
          cells: [{ value: "Month" }, { value: "Sales" }, { value: "Expenses" }],
        },
        ...[
          ["Jan", 1000, 400],
          ["Feb", 1200, 500],
          ["Mar", 1500, 600],
          ["Apr", 1300, 450],
          ["May", 1700, 700],
          ["Jun", 2000, 800],
        ].map(([month, sales, expenses]) => ({
          cells: [{ value: month }, { value: sales }, { value: expenses }],
        })),
      ],
    },
  ],
  chartsheets: [
    {
      name: "Sales Chart",
      tabColor: "FF4472C4",
      chart: {
        type: "column",
        title: "Monthly Sales vs Expenses",
        categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        series: [
          { name: "Sales", values: [1000, 1200, 1500, 1300, 1700, 2000] },
          { name: "Expenses", values: [400, 500, 600, 450, 700, 800] },
        ],
      },
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
