import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Page Setup Demo",
      tabColor: { rgb: "FF4472C4" },
      rows: [
        {
          cells: [
            { value: "Product" },
            { value: "Q1" },
            { value: "Q2" },
            { value: "Q3" },
            { value: "Q4" },
          ],
        },
        {
          cells: [
            { value: "Widget A" },
            { value: 100 },
            { value: 200 },
            { value: 300 },
            { value: 400 },
          ],
        },
        {
          cells: [
            { value: "Widget B" },
            { value: 150 },
            { value: 250 },
            { value: 350 },
            { value: 450 },
          ],
        },
        {
          cells: [
            { value: "Widget C" },
            { value: 200 },
            { value: 300 },
            { value: 400 },
            { value: 500 },
          ],
        },
      ],
      pageSetup: {
        orientation: "landscape",
        paperSize: 1,
        fitToWidth: 1,
        fitToHeight: 0,
      },
      headerFooter: {
        oddHeader: "Sales Report - &P",
        oddFooter: "Confidential",
        evenHeader: "Sales Report (Even) - &P",
        evenFooter: "Confidential (Even)",
      },
    },
    {
      name: "Tab Colors",
      tabColor: { rgb: "FFED7D31" },
      rows: [{ cells: [{ value: "Orange tab" }] }],
    },
    {
      name: "Portrait",
      tabColor: { theme: 3, tint: 0.3 },
      pageSetup: { orientation: "portrait", scale: 80 },
      headerFooter: { oddHeader: "Centered Header", oddFooter: "Page &P of &N" },
      rows: [{ cells: [{ value: "Portrait with scale 80%" }] }],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
