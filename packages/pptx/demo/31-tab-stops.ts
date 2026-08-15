import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

// Demonstrates paragraph tab stops (a:tabLst) — explicit tab positions and
// alignments, exercised through the shared core DrawingML text descriptors.

const options: PresentationOptions = {
  title: "Tab stops",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1cm",
            y: "1cm",
            width: "18cm",
            height: "1.5cm",
            textBody: { text: "Tab stops" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1cm",
            y: "3cm",
            width: "18cm",
            height: "5cm",
            textBody: {
              paragraphs: [
                {
                  properties: {
                    tabStops: [
                      { position: 4572000, alignment: "l" },
                      { position: 9144000, alignment: "dec" },
                    ],
                  },
                  children: [{ text: "Left tab then decimal tab", size: 18 }],
                },
                {
                  properties: {
                    tabStops: [{ position: 6858000, alignment: "r" }],
                  },
                  children: [{ text: "Right-aligned tab stop", size: 18 }],
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
fs.writeFileSync(".temp/31-tab-stops.pptx", buffer);
