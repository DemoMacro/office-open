import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Charts",
      rows: [
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
          row: 18,
        },
        {
          type: "pie",
          title: "Q4 Share",
          categories: ["2023 Q4", "2024 Q4"],
          series: [{ name: "Revenue", values: [180, 200] }],
          col: 14,
          row: 1,
        },
        {
          type: "area",
          title: "Revenue Area",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2023", values: [100, 150, 200, 180] },
            { name: "2024", values: [120, 170, 220, 200] },
          ],
          col: 14,
          row: 18,
        },
        {
          type: "scatter",
          title: "Revenue Scatter",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [{ name: "2024", values: [120, 170, 220, 200] }],
          col: 5,
          row: 35,
        },
        {
          type: "doughnut",
          title: "Annual Share",
          categories: ["2023", "2024"],
          series: [{ name: "Total", values: [630, 710] }],
          col: 14,
          row: 35,
        },
        {
          type: "radar",
          title: "Quarterly Radar",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2023", values: [100, 150, 200, 180] },
            { name: "2024", values: [120, 170, 220, 200] },
          ],
          col: 5,
          row: 52,
        },
        {
          type: "stock",
          title: "Daily Stock",
          categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
          series: [
            { name: "High", values: [105, 108, 112, 110, 115] },
            { name: "Low", values: [98, 103, 106, 105, 110] },
            { name: "Close", values: [103, 106, 110, 108, 113] },
          ],
          col: 14,
          row: 52,
        },
        {
          type: "surface",
          title: "Revenue Surface",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2023", values: [100, 150, 200, 180] },
            { name: "2024", values: [120, 170, 220, 200] },
          ],
          col: 5,
          row: 69,
        },
        {
          type: "column",
          threeD: true,
          title: "Quarterly Revenue (3D)",
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2023", values: [100, 150, 200, 180] },
            { name: "2024", values: [120, 170, 220, 200] },
          ],
          col: 14,
          row: 69,
        },
        {
          type: "pie",
          threeD: true,
          title: "Q4 Share (3D)",
          categories: ["2023 Q4", "2024 Q4"],
          series: [{ name: "Revenue", values: [180, 200] }],
          col: 5,
          row: 86,
        },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
