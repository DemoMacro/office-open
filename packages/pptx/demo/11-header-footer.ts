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
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Slide 1 - Default footer" },
            fill: "4472C4",
          },
        },
      ],
      headerFooter: { slideNumber: true, dateTime: true, footer: "Confidential" },
    },
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Slide 2 - No date" },
            fill: "ED7D31",
          },
        },
      ],
      headerFooter: { slideNumber: true, dateTime: false, footer: "My Presentation" },
    },
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Slide 3 - Only slide number" },
            fill: "70AD47",
          },
        },
      ],
      headerFooter: { slideNumber: true, dateTime: false, footer: false },
    },
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 500,
            height: 60,
            textBody: { text: "Slide 4 - No header/footer" },
            fill: "FFC000",
          },
        },
      ],
      headerFooter: { slideNumber: false, dateTime: false, footer: false },
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
