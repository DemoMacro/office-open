// Demo: Chart - all chart types
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, ChartRun } from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun({ text: "Chart Demo", bold: true, size: 32 })],
          spacing: { after: 400 },
        }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Column Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "column",
              categories: ["Q1", "Q2", "Q3", "Q4"],
              series: [
                { name: "2024", values: [120, 150, 180, 200] },
                { name: "2025", values: [140, 170, 210, 250] },
              ],
              title: "Quarterly Revenue",
              transformation: { width: 500, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Pie Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "pie",
              categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
              series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
              title: "Market Share",
              transformation: { width: 400, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Line Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "line",
              categories: ["Jan", "Feb", "Mar", "Apr", "May"],
              series: [
                { name: "2024", values: [10, 15, 13, 20, 25] },
                { name: "2025", values: [12, 18, 16, 22, 28] },
              ],
              title: "Monthly Revenue",
              transformation: { width: 500, height: 300 },
              style: 2,
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Area Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "area",
              categories: ["Q1", "Q2", "Q3", "Q4"],
              series: [
                { name: "2024", values: [100, 200, 300, 400] },
                { name: "2025", values: [120, 220, 350, 500] },
              ],
              transformation: { width: 500, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Scatter Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "scatter",
              categories: ["A", "B", "C", "D"],
              series: [{ name: "Data", values: [160, 175, 155, 170] }],
              title: "Height vs Weight",
              transformation: { width: 500, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Doughnut Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "doughnut",
              categories: ["North", "South", "East", "West"],
              series: [{ name: "Revenue", values: [35, 25, 22, 18] }],
              title: "Revenue by Region",
              transformation: { width: 400, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Radar Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "radar",
              categories: ["Speed", "Strength", "Agility", "Endurance", "Strategy"],
              series: [
                { name: "Team A", values: [85, 70, 90, 60, 80] },
                { name: "Team B", values: [70, 90, 65, 85, 75] },
              ],
              title: "Team Skills",
              transformation: { width: 400, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Stock Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "stock",
              categories: ["Mon", "Tue", "Wed", "Thu", "Fri"],
              series: [
                { name: "High", values: [105, 108, 112, 110, 115] },
                { name: "Low", values: [98, 103, 106, 105, 110] },
                { name: "Close", values: [103, 106, 110, 108, 113] },
              ],
              title: "Daily Stock",
              transformation: { width: 500, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "Surface Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "surface",
              categories: ["Jan", "Feb", "Mar", "Apr", "May"],
              series: [
                { name: "North", values: [5, 8, 15, 20, 25] },
                { name: "South", values: [20, 22, 25, 30, 35] },
              ],
              title: "Temperature Surface",
              transformation: { width: 500, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "3D Column Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
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
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [new TextRun({ bold: true, text: "3D Pie Chart", size: 28 })],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ChartRun({
              type: "pie",
              threeD: true,
              categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
              series: [{ name: "Browser", values: [65, 18, 3, 5, 9] }],
              title: "Market Share (3D)",
              transformation: { width: 400, height: 300 },
            }),
          ],
        }),
        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [
            new TextRun({
              text: "Note: Open this document in Microsoft Word to see the charts rendered.",
              italics: true,
              color: "888888",
            }),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
