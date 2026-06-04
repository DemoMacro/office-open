// OLE Object on a slide, docx Frameset, and docx Object embedding.

import { writeFileSync } from "node:fs";

import { Packer, Presentation, Shape } from "@office-open/pptx";

// PPTX: OLE Object on a slide
// Note: This generates the XML structure; a real OLE embed would need
// actual OLE binary data registered via relationships.
const pres = new Presentation({
  slides: [
    {
      children: [
        new Shape({
          x: 50,
          y: 20,
          width: 600,
          height: 40,
          textBody: { text: "OLE Object Demo" },
          fill: "4472C4",
        }),
        {
          ole: {
            x: 100,
            y: 100,
            width: 400,
            height: 300,
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
});

const buffer = await Packer.toBuffer(pres);
writeFileSync("My Presentation.pptx", buffer);
