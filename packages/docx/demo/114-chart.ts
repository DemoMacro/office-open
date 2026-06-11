// Demo: Chart - all chart types
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "Chart Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        {
          paragraph: {
            children: [{ bold: true, text: "Column Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "column",
                  categories: ["Q1", "Q2", "Q3", "Q4"],
                  series: [
                    { name: "2024", values: [120, 150, 180, 200] },
                    { name: "2025", values: [140, 170, 210, 250] },
                  ],
                  title: "Quarterly Revenue",
                  transformation: { width: 500, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Pie Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "pie",
                  categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
                  series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
                  title: "Market Share",
                  transformation: { width: 400, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Line Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "line",
                  categories: ["Jan", "Feb", "Mar", "Apr", "May"],
                  series: [
                    { name: "2024", values: [10, 15, 13, 20, 25] },
                    { name: "2025", values: [12, 18, 16, 22, 28] },
                  ],
                  title: "Monthly Revenue",
                  transformation: { width: 500, height: 300 },
                  style: 2,
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Area Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "area",
                  categories: ["Q1", "Q2", "Q3", "Q4"],
                  series: [
                    { name: "2024", values: [100, 200, 300, 400] },
                    { name: "2025", values: [120, 220, 350, 500] },
                  ],
                  transformation: { width: 500, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Scatter Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "scatter",
                  categories: ["A", "B", "C", "D"],
                  series: [{ name: "Data", values: [160, 175, 155, 170] }],
                  title: "Height vs Weight",
                  transformation: { width: 500, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Doughnut Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "doughnut",
                  categories: ["North", "South", "East", "West"],
                  series: [{ name: "Revenue", values: [35, 25, 22, 18] }],
                  title: "Revenue by Region",
                  transformation: { width: 400, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Radar Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "radar",
                  categories: ["Speed", "Strength", "Agility", "Endurance", "Strategy"],
                  series: [
                    { name: "Team A", values: [85, 70, 90, 60, 80] },
                    { name: "Team B", values: [70, 90, 65, 85, 75] },
                  ],
                  title: "Team Skills",
                  transformation: { width: 400, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Stock Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "stock",
                  categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
                  series: [
                    { name: "High", values: [105, 108, 112, 110, 115] },
                    { name: "Low", values: [98, 103, 106, 105, 110] },
                    { name: "Close", values: [103, 106, 110, 108, 113] },
                  ],
                  title: "Daily Stock",
                  transformation: { width: 500, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "Surface Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "surface",
                  categories: ["Jan", "Feb", "Mar", "Apr", "May"],
                  series: [
                    { name: "North", values: [5, 8, 15, 20, 25] },
                    { name: "South", values: [20, 22, 25, 30, 35] },
                  ],
                  title: "Temperature Surface",
                  transformation: { width: 500, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "3D Column Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "column",
                  threeD: true,
                  categories: ["Q1", "Q2", "Q3", "Q4"],
                  series: [
                    { name: "2024", values: [120, 150, 180, 200] },
                    { name: "2025", values: [140, 170, 210, 250] },
                  ],
                  title: "Quarterly Revenue (3D)",
                  transformation: { width: 500, height: 300 },
                  style: 2,
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [{ bold: true, text: "3D Pie Chart", size: 14 }],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                chart: {
                  type: "pie",
                  threeD: true,
                  categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
                  series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
                  title: "Market Share (3D)",
                  transformation: { width: 400, height: 300 },
                },
              },
            ],
          },
        },
        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [
              {
                text: "Note: Open this document in Microsoft Word to see the charts rendered.",
                italic: true,
                color: "888888",
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
