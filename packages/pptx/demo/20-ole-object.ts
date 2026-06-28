// OLE Object on a slide, docx Frameset, and docx Object embedding.

import { writeFileSync } from "node:fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// PPTX: OLE Object on a slide
// Note: This generates the XML structure; a real OLE embed would need
// actual OLE binary data registered via relationships.
const options: PresentationOptions = {
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.5cm",
            width: "15.9cm",
            height: "1.1cm",
            textBody: { text: "OLE Object Demo" },
            fill: "4472C4",
          },
        },
        {
          ole: {
            x: "2.6cm",
            y: "2.6cm",
            width: "10.6cm",
            height: "7.9cm",
            progId: "Excel.Sheet.12",
            spid: "sp1025",
            name: "Embedded Worksheet",
            showAsIcon: false,
            embed: {
              rId: "rIdOle1",
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
writeFileSync("My Presentation.pptx", buffer);
