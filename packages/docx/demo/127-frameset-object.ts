// Frameset (for webSettings.xml) and Object embedding (w:object with OLE).
//
// Frameset defines a frames-based layout for web-published documents.
// ObjectElement embeds an OLE object (e.g., Excel worksheet) using VML shape.

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, Frameset, ObjectElement } from "@office-open/docx";

// Demo: Generate Frameset XML (for webSettings)
const frameset = new Frameset({
  layout: "rows",
  size: "100%",
  splitbar: {
    width: 120,
    color: "auto",
  },
  title: "Main Frameset",
  children: [
    {
      size: "60%",
      name: "content",
      sourceRId: "rId1",
      scrollbar: "auto",
      linkedToFile: true,
    },
    {
      size: "40%",
      name: "nav",
      sourceRId: "rId2",
      scrollbar: "on",
      noResizeAllowed: true,
    },
  ],
});

// Demo: Generate ObjectElement XML (w:object with OLE embed)
const oleObject = new ObjectElement({
  shapeId: "_x0000_i1025",
  dxaOrig: 4320,
  dyaOrig: 2592,
  style: {
    width: "3in",
    height: "1.8in",
    position: "absolute",
    left: "1in",
    top: "1in",
  },
  embed: {
    rId: "rIdEmbed1",
    progId: "Excel.Sheet.12",
    drawAspect: "Content",
  },
});

// Log the standalone XML output (using toXml directly, no XML declaration)
const ctx = { fileData: null, stack: [] };
console.log("Frameset XML:", frameset.toXml(ctx).slice(0, 300));
console.log("ObjectElement XML:", oleObject.toXml(ctx).slice(0, 300));

// Also produce a regular document
const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "Frameset and OLE Object Demo",
              size: 32,
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new TextRun(
              "This document demonstrates Frameset (webSettings) and OLE Object embedding components.",
            ),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
