import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Locked Canvas Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: {
              paragraphs: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [{ text: "Locked Canvas Demo", size: 32, bold: true }],
                },
              ],
            },
          },
        },
        {
          lockedCanvas: {
            x: "2.6cm",
            y: "3.2cm",
            width: "15.9cm",
            height: "7.9cm",
            children: [
              {
                x: "0.5cm",
                y: "0.5cm",
                width: "6.6cm",
                height: "2.6cm",
                fill: "4472C4",
                textBody: "Locked Shape 1",
              },
              {
                x: "7.9cm",
                y: "0.5cm",
                width: "6.6cm",
                height: "2.6cm",
                fill: "ED7D31",
                textBody: "Locked Shape 2",
              },
            ],
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/26-locked-canvas.pptx", buffer);
