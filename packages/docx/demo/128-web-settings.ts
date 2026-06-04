// Web Settings: encoding, optimizeForBrowser, pixelsPerInch, targetScreenSz,
// frameset layout with splitbar, and div elements with borders.

import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
  webSettings: {
    encoding: "utf-8",
    optimizeForBrowser: true,
    pixelsPerInch: 96,
    allowPNG: true,
    doNotRelyOnCSS: false,
    targetScreenSz: "1024x768",
    divs: [
      {
        marginLeft: 100,
        marginRight: 100,
        marginTop: 50,
        marginBottom: 50,
        blockQuote: true,
        border: {
          top: { style: "single", color: "000000", size: 4 },
          left: { style: "single", color: "000000", size: 4 },
          bottom: { style: "single", color: "000000", size: 4 },
          right: { style: "single", color: "000000", size: 4 },
        },
      },
    ],
  },

  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "Web Settings Demo",
              size: 32,
            }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun(
              "This document includes web settings for browser rendering optimization, encoding, and div formatting.",
            ),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
