import { readFileSync, writeFileSync } from "node:fs";

// Example of using tableLook to control conditional table formatting
import { generateDocument, WidthType } from "@office-open/docx";

const styles = readFileSync("./demo/assets/custom-styles.xml", "utf8");

const buffer = await generateDocument({
  externalStyles: styles,
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ bold: true, text: "Table 1: Table Look Default Values" }],
          },
        },
        { paragraph: { text: "" } },
        {
          table: {
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Header 1" }] },
                  { children: [{ paragraph: "Header 2" }] },
                  { children: [{ paragraph: "Header 3" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Row 1, Col 1" }] },
                  { children: [{ paragraph: "Row 1, Col 2" }] },
                  { children: [{ paragraph: "Row 1, Col 3" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Row 2, Col 1" }] },
                  { children: [{ paragraph: "Row 2, Col 2" }] },
                  { children: [{ paragraph: "Row 2, Col 3" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Row 3, Col 1" }] },
                  { children: [{ paragraph: "Row 3, Col 2" }] },
                  { children: [{ paragraph: "Row 3, Col 3" }] },
                ],
              },
            ],
            style: "MyCustomTableStyle",
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
        { paragraph: { text: "" } },
        { paragraph: { text: "" } },
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "Table 2: Table Look All Look Values Enabled",
              },
            ],
          },
        },
        { paragraph: { text: "" } },
        {
          table: {
            rows: [
              {
                cells: [
                  { children: [{ paragraph: "Header 1" }] },
                  { children: [{ paragraph: "Header 2" }] },
                  { children: [{ paragraph: "Header 3" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Row 1, Col 1" }] },
                  { children: [{ paragraph: "Row 1, Col 2" }] },
                  { children: [{ paragraph: "Row 1, Col 3" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Row 2, Col 1" }] },
                  { children: [{ paragraph: "Row 2, Col 2" }] },
                  { children: [{ paragraph: "Row 2, Col 3" }] },
                ],
              },
              {
                cells: [
                  { children: [{ paragraph: "Row 3, Col 1" }] },
                  { children: [{ paragraph: "Row 3, Col 2" }] },
                  { children: [{ paragraph: "Row 3, Col 3" }] },
                ],
              },
            ],
            style: "MyCustomTableStyle",
            tableLook: {
              firstColumn: true,
              firstRow: true,
              lastColumn: true,
              lastRow: true,
              noHBand: false,
              noVBand: false,
            },
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
