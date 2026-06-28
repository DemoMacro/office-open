import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
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
      // Chartsheet extensions: pageMargins, pageSetup, headerFooter, sheetProtection
      pageMargins: {
        left: "1.3cm",
        right: "1.3cm",
        top: "1.9cm",
        bottom: "1.9cm",
        header: "0.8cm",
        footer: "0.8cm",
      },
      pageSetup: {
        paperSize: 9, // A4
        orientation: "landscape",
        horizontalDpi: 300,
        verticalDpi: 300,
      },
      headerFooter: {
        differentFirst: true,
        oddHeader: "&CChartsheet Demo",
        oddFooter: "&CPage &P of &N",
      },
      sheetProtection: {
        content: true,
        objects: true,
      },
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

writeFileSync("My Workbook.xlsx", buffer);
