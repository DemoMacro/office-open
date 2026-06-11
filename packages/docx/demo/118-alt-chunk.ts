// Demo: AltChunk (Alternative Format Content)
// Demonstrates embedding HTML/RTF/plain-text content into a Word document via w:altChunk.
// Note: HTML fragments are automatically wrapped in a full document structure.
// Word only supports a limited HTML subset (no CSS, no div/span, no script/style).

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
                text: "AltChunk (Alternative Format Content) Demo",
                bold: true,
                size: 16,
              },
            ],
            spacing: { after: 400 },
          },
        },

        {
          paragraph: {
            children: [{ bold: true, text: "1. Embedded HTML Content", size: 14 }],
            spacing: { after: 200 },
          },
        },

        {
          altChunk: {
            data: `<p><b>Embedded HTML Heading</b></p>
<p>This content was embedded from HTML. It supports <b>bold</b>, <i>italic</i>, and other HTML formatting.</p>
<ul>
  <li>Item one</li>
  <li>Item two</li>
  <li>Item three</li>
</ul>`,
            contentType: "text/html",
            extension: "html",
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. HTML Table with Match Source Formatting",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },

        {
          altChunk: {
            data: `<table border="1">
  <tr>
    <th><b>Name</b></th>
    <th><b>Value</b></th>
  </tr>
  <tr>
    <td>Item A</td>
    <td>100</td>
  </tr>
  <tr>
    <td>Item B</td>
    <td>200</td>
  </tr>
</table>`,
            contentType: "text/html",
            extension: "html",
            matchSrc: true,
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "3. RTF Content",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },

        {
          altChunk: {
            data: "{\\rtf1\\ansi\\b RTF bold text\\b0  and normal text.}",
            contentType: "application/rtf",
            extension: "rtf",
          },
        },

        { paragraph: "" },

        {
          paragraph: {
            children: ["This is native docx content after the alt chunk."],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
