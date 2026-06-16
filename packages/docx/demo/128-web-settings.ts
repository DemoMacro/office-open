// Web Settings: encoding, optimizeForBrowser, pixelsPerInch, targetScreenSize,
// frameset layout with splitbar, and div elements with borders.

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  webSettings: {
    encoding: "utf-8",
    optimizeForBrowser: true,
    pixelsPerInch: 96,
    allowPNG: true,
    doNotRelyOnCSS: false,
    targetScreenSize: "1024x768",
    divs: [
      {
        id: 100,
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
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "Web Settings Demo",
                size: 16,
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "This document includes web settings for browser rendering optimization, encoding, and div formatting.",
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
