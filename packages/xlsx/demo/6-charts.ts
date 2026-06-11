import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
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
    {
      name: "AdvancedCharts",
      rows: [
        { cells: [{ value: "Month" }, { value: "Sales" }, { value: "Target" }] },
        { cells: [{ value: "Jan" }, { value: 120 }, { value: 100 }] },
        { cells: [{ value: "Feb" }, { value: 145 }, { value: 110 }] },
        { cells: [{ value: "Mar" }, { value: 160 }, { value: 120 }] },
        { cells: [{ value: "Apr" }, { value: 135 }, { value: 130 }] },
        { cells: [{ value: "May" }, { value: 180 }, { value: 140 }] },
        { cells: [{ value: "Jun" }, { value: 210 }, { value: 150 }] },
      ],
      charts: [
        // Line chart with trendline, error bars, and data labels
        {
          type: "line",
          title: "Sales with Trendline",
          categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
          series: [
            {
              name: "Sales",
              values: [120, 145, 160, 135, 180, 210],
              trendlines: [{ type: "linear", dispRSqr: true, dispEq: true }],
              errorBars: {
                direction: "y",
                barType: "both",
                valueType: "stdErr",
              },
              dataLabels: {
                showVal: true,
                position: "t",
              },
            },
            {
              name: "Target",
              values: [100, 110, 120, 130, 140, 150],
              trendlines: [{ type: "poly", order: 2 }],
            },
          ],
          col: 5,
          row: 1,
        },
        // Bar chart with data labels and moving average trendline
        {
          type: "column",
          title: "Sales vs Target",
          categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
          series: [
            {
              name: "Sales",
              values: [120, 145, 160, 135, 180, 210],
              dataLabels: {
                showVal: true,
                position: "outEnd",
              },
            },
            {
              name: "Target",
              values: [100, 110, 120, 130, 140, 150],
              trendlines: [{ type: "movingAvg", period: 2 }],
            },
          ],
          col: 5,
          row: 18,
        },
        // Scatter chart with linear trendline and fixed error bars
        {
          type: "scatter",
          title: "Sales Scatter with Trend",
          categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
          series: [
            {
              name: "Sales",
              values: [120, 145, 160, 135, 180, 210],
              trendlines: [{ type: "linear", forward: 2 }],
              errorBars: {
                direction: "y",
                barType: "both",
                valueType: "fixedVal",
                value: 10,
              },
            },
          ],
          col: 5,
          row: 35,
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
