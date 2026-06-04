// Custom XML elements in tables: CustomXmlRow and CustomXmlCell.
//
// CustomXmlRow wraps table rows with custom XML metadata.
// CustomXmlCell wraps table cells with custom XML metadata.

import * as fs from "fs";

import {
  CustomXmlCell,
  CustomXmlRow,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "Custom XML in Tables",
              size: 32,
            }),
          ],
          spacing: { after: 200 },
        }),

        // Table with CustomXmlRow wrapping a regular TableRow
        new Table({
          rows: [
            new TableRow({
              cells: [
                new TableCell({ children: [new Paragraph("Header 1")] }),
                new TableCell({ children: [new Paragraph("Header 2")] }),
              ],
            }),
            new CustomXmlRow({
              element: "dataRow",
              uri: "http://example.com/ns",
              children: [
                new TableRow({
                  cells: [
                    new TableCell({ children: [new Paragraph("Row 2, Cell 1")] }),
                    new TableCell({ children: [new Paragraph("Row 2, Cell 2")] }),
                  ],
                }),
              ],
            }),
            new TableRow({
              cells: [
                // CustomXmlCell wrapping a TableCell
                new CustomXmlCell({
                  element: "importantValue",
                  uri: "http://example.com/ns",
                  children: [new TableCell({ children: [new Paragraph("Important")] })],
                }),
                new TableCell({ children: [new Paragraph("Normal cell")] }),
              ],
            }),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
