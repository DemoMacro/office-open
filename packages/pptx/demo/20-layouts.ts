import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Layouts Demo",
  creator: "Demo",
  masters: [
    {
      background: {
        fill: "1B2A4A",
      },
      children: [
        {
          shape: {
            x: "0.0cm",
            y: "18.4cm",
            width: "33.9cm",
            height: "0.7cm",
            properties: {
              geometry: "rect",

              fill: "4472C4",
            },
          },
        },
      ],
      placeholders: {
        slideNumber: {
          x: "23.9cm",
          y: "17.6cm",
          width: "7.6cm",
          height: "1.0cm",
        },
      },
      layouts: [
        // Preset layouts
        { type: "title" },
        { type: "object" },
        { type: "twoColumnText" },
        { type: "blank" },
        { type: "titleOnly" },
        // Custom layout with decorative shapes + positioned placeholders
        {
          name: "Hero",
          type: "blank",
          // Per-layout theme deviation — emits ppt/theme/themeOverride1.xml
          themeOverride: {
            colorScheme: { accent1: "C0504D", dark1: "1B2A4A" },
          },
          children: [
            {
              shape: {
                x: "0.0cm",
                y: "0.0cm",
                width: "33.9cm",
                height: "7.4cm",
                properties: {
                  geometry: "rect",

                  fill: "4472C4",
                },
              },
            },
          ],
          placeholders: {
            title: {
              x: "2.6cm",
              y: "1.3cm",
              width: "28.6cm",
              height: "4.8cm",
            },
            body: {
              x: "2.6cm",
              y: "8.5cm",
              width: "28.6cm",
              height: "7.9cm",
            },
            slideNumber: {
              x: "23.9cm",
              y: "17.6cm",
              width: "7.6cm",
              height: "1.0cm",
            },
          },
        },
      ],
    },
  ],
  slides: [
    {
      layout: "title",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: { text: "Title Slide Layout" },
            properties: {
              geometry: "rect",

              fill: "0F3460",
            },
          },
        },
      ],
    },
    {
      layout: "object",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "2.6cm",
            textBody: { text: "Title and Content Layout" },
            properties: {
              fill: "4472C4",
            },
          },
        },
        {
          shape: {
            x: "2.6cm",
            y: "6.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: { text: "Content area" },
            properties: {
              fill: "E8E8E8",
            },
          },
        },
      ],
    },
    {
      layout: "twoColumnText",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "2.6cm",
            textBody: { text: "Two Column Layout" },
            properties: {
              fill: "70AD47",
            },
          },
        },
      ],
    },
    {
      layout: "blank",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "7.9cm",
            textBody: { text: "Blank Layout" },
            properties: {
              fill: "FFC000",
            },
          },
        },
      ],
    },
    {
      layout: "titleOnly",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "2.6cm",
            textBody: { text: "Title Only Layout" },
            properties: {
              fill: "ED7D31",
            },
          },
        },
      ],
    },
    {
      layout: "Hero",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "1.3cm",
            width: "28.6cm",
            height: "4.8cm",
            textBody: { text: "Custom Hero Layout" },
            properties: {
              fill: "2E75B6",
            },
          },
        },
        {
          shape: {
            x: "2.6cm",
            y: "8.5cm",
            width: "28.6cm",
            height: "7.9cm",
            textBody: {
              text: "This slide uses a custom layout with decorative header bar and custom placeholder positions.",
            },
            properties: {
              fill: "E8E8E8",
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.mkdirSync(".temp", { recursive: true });
fs.writeFileSync(".temp/20-layouts.pptx", buffer);
