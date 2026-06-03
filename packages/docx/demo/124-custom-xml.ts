// Custom XML elements: embed structured XML markup within a Word document.
//
// NOTE: Microsoft Word shows a warning "XML elements that are no longer supported"
// when opening documents with w:customXml elements. This is due to the i4i patent
// ruling (2009). The OOXML spec still defines these elements, but Microsoft Office
// no longer processes them. Use content controls (SDT) for structured data instead.
//
// Use cases: interoperability with non-Microsoft tools that support custom XML.

import * as fs from "fs";

import {
  CustomXmlBlock,
  CustomXmlRun,
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
          children: [
            new TextRun({
              bold: true,
              text: "Custom XML Elements",
              size: 32,
            }),
          ],
          spacing: { after: 200 },
        }),

        new Paragraph({
          children: [
            new TextRun("Inline custom XML: Price = "),
            new CustomXmlRun({
              element: "price",
              uri: "http://example.com/ns",
              children: [new TextRun("99.99")],
            }),
          ],
        }),

        new Paragraph({ children: [new TextRun("")] }),

        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "Block-level custom XML:",
              size: 28,
            }),
          ],
          spacing: { after: 200 },
        }),
        new CustomXmlBlock({
          element: "invoiceItems",
          uri: "http://example.com/ns",
          customXmlPr: {
            placeholder: "Invoice items",
            attributes: [
              { name: "status", val: "draft" },
              { name: "version", val: "1.0", uri: "http://example.com/ns" },
            ],
          },
          children: [
            new Paragraph({
              children: [new TextRun("Item 1: Widget A - $50.00")],
            }),
            new Paragraph({
              children: [new TextRun("Item 2: Widget B - $49.99")],
            }),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
