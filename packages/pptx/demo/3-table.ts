import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const CUSTOM_STYLE_ID = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}";

const options: PresentationOptions = {
  title: "Phase 3 Demo",
  creator: "Demo",
  tableStyles: {
    defaultStyleId: CUSTOM_STYLE_ID,
    styles: [
      {
        styleId: CUSTOM_STYLE_ID,
        styleName: "Custom Blue Stripe",
        regions: {
          wholeTbl: {
            text: { bold: "off" },
          },
          firstRow: {
            text: { bold: "on" },
          },
          band1H: {
            text: { bold: "off" },
          },
        },
      },
    ],
  },
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.6cm",
            textBody: { text: "Table Demo" },
            fill: "4472C4",
          },
        },
        {
          table: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "6.6cm",
            rows: [
              {
                cells: [
                  {
                    text: "Name",
                    fill: "4472C4",
                  },
                  { text: "Age" },
                  { text: "City" },
                ],
              },
              {
                cells: [{ text: "Alice" }, { text: "30" }, { text: "Beijing" }],
              },
              {
                cells: [{ text: "Bob" }, { text: "25" }, { text: "Shanghai" }],
              },
              {
                cells: [{ text: "Charlie" }, { text: "35" }, { text: "Shenzhen" }],
              },
            ],
            columnWidths: [2000000, 1500000, 2500000],
            firstRow: true,
            bandRow: true,
          },
        },
      ],
    },
    {
      background: { fill: "F5F5F5" },
      children: [
        {
          table: {
            x: "2.6cm",
            y: "2.1cm",
            width: "13.2cm",
            height: "5.3cm",
            rows: [
              {
                cells: [
                  {
                    text: "Header 1",
                    fill: "ED7D31",
                  },
                  { text: "Header 2", fill: "ED7D31" },
                ],
              },
              {
                cells: [{ text: "A" }, { text: "B" }],
              },
              {
                cells: [{ text: "C" }, { text: "D" }],
              },
            ],
            columnWidths: [2500000, 2500000],
          },
        },
      ],
    },
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Vertical Align & Cell Margins" },
            fill: "4472C4",
          },
        },
        {
          table: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "6.6cm",
            rows: [
              {
                height: "18520.8cm",
                cells: [
                  {
                    text: "Top",
                    verticalAlign: "top",
                    fill: "E8F0FE",
                  },
                  {
                    text: "Center",
                    verticalAlign: "center",
                    fill: "E8F0FE",
                  },
                  {
                    text: "Bottom",
                    verticalAlign: "bottom",
                    fill: "E8F0FE",
                  },
                ],
              },
              {
                height: "13229.2cm",
                cells: [
                  { text: "Default" },
                  { text: "Wide L/R", margins: { left: 300000, right: 300000 } },
                  { text: "Wide T/B", margins: { top: 80000, bottom: 80000 } },
                ],
              },
            ],
          },
        },
      ],
    },
    // Slide 4: Merged cells
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Merged Cells" },
            fill: "4472C4",
          },
        },
        {
          table: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "5.3cm",
            rows: [
              {
                cells: [
                  {
                    text: "A",
                    columnSpan: 2,
                    fill: "E8F0FE",
                  },
                  { text: "C" },
                ],
              },
              {
                cells: [{ text: "D" }, { text: "E" }, { text: "F" }],
              },
              {
                cells: [
                  {
                    text: "Merged",
                    rowSpan: 2,
                    fill: "FFF2CC",
                  },
                  { text: "H" },
                  { text: "I" },
                ],
              },
              {
                cells: [{ text: "K" }, { text: "L" }],
              },
            ],
          },
        },
      ],
    },
    // Slide 5: Custom table style
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "15.9cm",
            height: "1.3cm",
            textBody: { text: "Custom Table Style" },
            fill: "4472C4",
          },
        },
        {
          table: {
            x: "1.3cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "6.6cm",
            rows: [
              {
                cells: [{ text: "Product" }, { text: "Q1" }, { text: "Q2" }],
              },
              {
                cells: [{ text: "Widget" }, { text: "120" }, { text: "150" }],
              },
              {
                cells: [{ text: "Gadget" }, { text: "80" }, { text: "95" }],
              },
              {
                cells: [{ text: "Doohickey" }, { text: "200" }, { text: "180" }],
              },
            ],
            columnWidths: [2500000, 1750000, 1750000],
            firstRow: true,
            bandRow: true,
            tableStyleId: CUSTOM_STYLE_ID,
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
