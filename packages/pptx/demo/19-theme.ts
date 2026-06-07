import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

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
            x: 100,
            y: 100,
            width: 600,
            height: 300,
            textBody: { text: "Custom Theme" },
            geometry: "rect",
            fill: "0F3460",
          },
        },
        {
          shape: {
            x: 100,
            y: 450,
            width: 600,
            height: 100,
            textBody: { text: "Accent2 highlight" },
            geometry: "rect",
            fill: "E94560",
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
