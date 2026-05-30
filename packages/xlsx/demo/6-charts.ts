import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Data",
      children: [
        { cells: [{ value: "Quarter" }, { value: "2023" }, { value: "2024" }] },
        { cells: [{ value: "Q1" }, { value: 100 }, { value: 120 }] },
        { cells: [{ value: "Q2" }, { value: 150 }, { value: 170 }] },
        { cells: [{ value: "Q3" }, { value: 200 }, { value: 220 }] },
        { cells: [{ value: "Q4" }, { value: 180 }, { value: 200 }] },
      ],
      charts: [
        {
          type: "column",
          title: "Quarterly Revenue",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2023", values: [100, 150, 200, 180] },
            { name: "2024", values: [120, 170, 220, 200] },
          ],
          col: 5,
          row: 1,
        },
        {
          type: "line",
          title: "Growth Trend",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2023", values: [100, 150, 200, 180] },
            { name: "2024", values: [120, 170, 220, 200] },
          ],
          col: 5,
          row: 20,
        },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
