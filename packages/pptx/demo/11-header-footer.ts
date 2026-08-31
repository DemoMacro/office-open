import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Header Footer Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Slide 1 - Default footer" },
            properties: {
              fill: "4472C4",
            },
          },
        },
      ],
      headerFooter: { slideNumber: true, dateTime: true, footer: "Confidential" },
    },
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Slide 2 - No date" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
      ],
      headerFooter: { slideNumber: true, dateTime: false, footer: "My Presentation" },
    },
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Slide 3 - Only slide number" },
            properties: {
              fill: "70AD47",
            },
          },
        },
      ],
      headerFooter: { slideNumber: true, dateTime: false, footer: false },
    },
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Slide 4 - No header/footer" },
            properties: {
              fill: "FFC000",
            },
          },
        },
      ],
      headerFooter: { slideNumber: false, dateTime: false, footer: false },
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/11-header-footer.pptx", buffer);
