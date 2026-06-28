import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Merge & Size",
      rows: [
        // Row 1: merged title
        {
          cells: [{ value: "Quarterly Report", style: { font: { bold: true, size: 14 } } }],
        },
        // Row 2: headers
        {
          cells: [
            { value: "Q1", style: { font: { bold: true }, alignment: { horizontal: "center" } } },
            { value: "Q2", style: { font: { bold: true }, alignment: { horizontal: "center" } } },
            { value: "Q3", style: { font: { bold: true }, alignment: { horizontal: "center" } } },
            { value: "Q4", style: { font: { bold: true }, alignment: { horizontal: "center" } } },
          ],
        },
        // Row 3: data
        {
          cells: [{ value: 100 }, { value: 200 }, { value: 150 }, { value: 300 }],
        },
        // Row 4: taller row
        {
          cells: [{ value: "Total", style: { font: { bold: true } } }, { value: 750 }],
          height: "1.1cm",
        },
      ],
      columns: [
        { min: 1, max: 1, width: 20 },
        { min: 2, max: 4, width: 12 },
        { min: 5, max: 5, width: 12 },
      ],
      mergeCells: [{ from: { row: 1, col: 1 }, to: { row: 1, col: 4 } }],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
