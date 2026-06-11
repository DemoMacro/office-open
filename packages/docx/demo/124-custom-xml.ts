// Custom XML elements: embed structured XML markup within a Word document.
//
// NOTE: Microsoft Word shows a warning "XML elements that are no longer supported"
// when opening documents with w:customXml elements. This is due to the i4i patent
// ruling (2009). The OOXML spec still defines these elements, but Microsoft Office
// no longer processes them. Use content controls (SDT) for structured data instead.
//
// Use cases: interoperability with non-Microsoft tools that support custom XML.

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "Custom XML Elements",
                size: 32,
              },
            ],
            spacing: { after: 200 },
          },
        },

        {
          paragraph: {
            children: [
              "Inline custom XML: Price = ",
              {
                customXml: {
                  element: "price",
                  uri: "http://example.com/ns",
                  children: ["99.99"],
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
                bold: true,
                text: "Block-level custom XML:",
                size: 28,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          customXml: {
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
              {
                paragraph: {
                  children: ["Item 1: Widget A - $50.00"],
                },
              },
              {
                paragraph: {
                  children: ["Item 2: Widget B - $49.99"],
                },
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
