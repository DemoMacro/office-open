import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Theme Demo",
  creator: "Demo",
  masters: [
    {
      theme: {
        name: "Brand Theme",
        colors: {
          dark1: "1A1A2E",
          light1: "FFFFFF",
          dark2: "16213E",
          light2: "E8E8E8",
          accent1: "0F3460",
          accent2: "E94560",
          accent3: "533483",
          accent4: "2B2D42",
          accent5: "8D99AE",
          accent6: "EDF2F4",
        },
        fonts: {
          majorFont: "Segoe UI",
          minorFont: "Segoe UI",
        },
      },
    },
  ],
  slides: [
    {
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "7.9cm",
            textBody: { text: "Custom Theme" },
            geometry: "rect",
            fill: "0F3460",
          },
        },
        {
          shape: {
            x: "2.6cm",
            y: "11.9cm",
            width: "15.9cm",
            height: "2.6cm",
            textBody: { text: "Accent2 highlight" },
            geometry: "rect",
            fill: "E94560",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
