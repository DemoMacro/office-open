import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  creator: "office-open",
  title: "Style Demo",
  worksheets: [
    {
      name: "Styles",
      rows: [
        // Header row with bold + background
        {
          cells: [
            {
              value: "Name",
              style: { font: { bold: true, color: "FFFFFF" }, fill: { color: "4472C4" } },
            },
            {
              value: "Age",
              style: { font: { bold: true, color: "FFFFFF" }, fill: { color: "4472C4" } },
            },
            {
              value: "Score",
              style: { font: { bold: true, color: "FFFFFF" }, fill: { color: "4472C4" } },
            },
          ],
        },
        // Data rows
        {
          cells: [
            { value: "Alice", style: { font: { italic: true } } },
            { value: 30, style: { alignment: { horizontal: "center" } } },
            { value: 95.5, style: { numFmt: "0.00" } },
          ],
        },
        {
          cells: [
            { value: "Bob", style: { border: { bottom: { style: "thin", color: "000000" } } } },
            { value: 25, style: { alignment: { horizontal: "center" } } },
            { value: 87.3, style: { numFmt: "0.00", font: { bold: true } } },
          ],
        },
        {
          cells: [
            { value: "Charlie" },
            { value: 28, style: { alignment: { horizontal: "center" } } },
            { value: 0.92, style: { numFmt: "0.00%" } },
          ],
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
