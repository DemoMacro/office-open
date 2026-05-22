// Image effects (shadow, glow, reflection, fillOverlay) applied via ImageRun
import * as fs from "fs";

import { Document, ImageRun, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              size: 32,
              text: "Image Effects (via ImageRun)",
            }),
          ],
          spacing: { after: 400 },
        }),

        // 1. Image with outer shadow
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "1. Outer Shadow",
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync("./demo/images/cat.jpg"),
              effects: {
                outerShadow: {
                  blurRadius: 50800,
                  color: { value: "000000" },
                  direction: 5400000,
                  distance: 38100,
                },
              },
              transformation: {
                height: 150,
                width: 150,
              },
              type: "jpg",
            }),
          ],
        }),

        new Paragraph({ children: [new TextRun("")] }),

        // 2. Image with glow
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "2. Glow Effect",
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync("./demo/images/cat.jpg"),
              effects: {
                glow: {
                  color: { value: "00B0F0" },
                  radius: 76200,
                },
              },
              transformation: {
                height: 150,
                width: 150,
              },
              type: "jpg",
            }),
          ],
        }),

        new Paragraph({ children: [new TextRun("")] }),

        // 3. Image with reflection
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "3. Reflection",
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync("./demo/images/cat.jpg"),
              effects: {
                reflection: {
                  blurRadius: 6350,
                  startAlpha: 40000,
                },
              },
              transformation: {
                height: 150,
                width: 150,
              },
              type: "jpg",
            }),
          ],
        }),

        new Paragraph({ children: [new TextRun("")] }),

        // 4. Image with soft edge
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "4. Soft Edge",
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync("./demo/images/cat.jpg"),
              effects: {
                softEdge: 50800,
              },
              transformation: {
                height: 150,
                width: 150,
              },
              type: "jpg",
            }),
          ],
        }),

        new Paragraph({ children: [new TextRun("")] }),

        // 5. Image with combined effects
        new Paragraph({
          children: [
            new TextRun({
              bold: true,
              text: "5. Combined (Shadow + Glow + Reflection)",
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync("./demo/images/cat.jpg"),
              effects: {
                glow: {
                  color: { value: "FFD700" },
                  radius: 50800,
                },
                outerShadow: {
                  blurRadius: 40000,
                  color: { value: "000000" },
                  direction: 5400000,
                  distance: 30000,
                },
                reflection: {
                  blurRadius: 4000,
                  startAlpha: 30000,
                },
              },
              transformation: {
                height: 150,
                width: 150,
              },
              type: "jpg",
            }),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
