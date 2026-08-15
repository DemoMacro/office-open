import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Demonstrates bullet color/size/font variants (buClrTx/buSzPts/buFontTx) and
// autoNum bullets, exercising the shared core DrawingML text descriptors.

const options: PresentationOptions = {
  title: "Bullets",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1cm",
            y: "1cm",
            width: "18cm",
            height: "1.5cm",
            textBody: { text: "Bullet variants" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1cm",
            y: "3cm",
            width: "18cm",
            height: "12cm",
            textBody: {
              paragraphs: [
                {
                  properties: {
                    bullet: { type: "char", char: "•", color: "FF0000", size: 120 },
                  },
                  children: [{ text: "Red bullet, 120% size", size: 18 }],
                },
                {
                  properties: {
                    bullet: { type: "char", char: "►", colorFollowsText: true },
                  },
                  children: [{ text: "Color follows text run", size: 18, fill: "00B050" }],
                },
                {
                  properties: {
                    bullet: { type: "char", char: "◆", sizePoints: 24 },
                  },
                  children: [{ text: "Fixed 24pt bullet size", size: 18 }],
                },
                {
                  properties: {
                    bullet: { type: "char", char: "✓", fontFollowsText: true },
                  },
                  children: [{ text: "Font follows text run", size: 18 }],
                },
                {
                  properties: {
                    bullet: { type: "autoNum", format: "arabicPeriod", startAt: 1 },
                  },
                  children: [{ text: "Numbered item", size: 18 }],
                },
              ],
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/30-bullets.pptx", buffer);
