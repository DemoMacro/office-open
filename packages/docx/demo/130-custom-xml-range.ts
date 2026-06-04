// Custom XML range markers: insert, delete, and move ranges for
// track changes on custom XML markup.

import * as fs from "fs";

import {
  CustomXmlDelRangeEnd,
  CustomXmlDelRangeStart,
  CustomXmlInsRangeEnd,
  CustomXmlInsRangeStart,
  CustomXmlMoveFromRangeEnd,
  CustomXmlMoveFromRangeStart,
  CustomXmlMoveToRangeEnd,
  CustomXmlMoveToRangeStart,
  Document,
  Packer,
  Paragraph,
  TextRun,
} from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun({ bold: true, text: "Custom XML Range Markers", size: 32 })],
        }),

        // Insert range
        new Paragraph({
          children: [
            new TextRun("Custom XML insert: "),
            new CustomXmlInsRangeStart(100, "Alice", "2026-06-04T10:00:00Z"),
            new TextRun("newly inserted content"),
            new CustomXmlInsRangeEnd(100),
            new TextRun("."),
          ],
        }),

        // Delete range
        new Paragraph({
          children: [
            new TextRun("Custom XML delete: "),
            new CustomXmlDelRangeStart(101, "Bob", "2026-06-04T11:00:00Z"),
            new TextRun("deleted content"),
            new CustomXmlDelRangeEnd(101),
            new TextRun("."),
          ],
        }),

        // Move range
        new Paragraph({
          children: [
            new TextRun("Move from: "),
            new CustomXmlMoveFromRangeStart(102, "Alice", "2026-06-04T12:00:00Z"),
            new TextRun("original location"),
            new CustomXmlMoveFromRangeEnd(102),
            new TextRun(" -> moved to: "),
            new CustomXmlMoveToRangeStart(103, "Alice", "2026-06-04T12:00:00Z"),
            new TextRun("new location"),
            new CustomXmlMoveToRangeEnd(103),
            new TextRun("."),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
