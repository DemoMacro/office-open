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
            geometry: "rect",
            fill: "4472C4",
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
        { type: "obj" },
        { type: "twoColTx" },
        { type: "blank" },
        { type: "titleOnly" },
        // Custom layout with decorative shapes + positioned placeholders
        {
          name: "Hero",
          type: "blank",
          children: [
            {
              shape: {
                x: "0.0cm",
                y: "0.0cm",
                width: "33.9cm",
                height: "7.4cm",
                geometry: "rect",
                fill: "4472C4",
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
            geometry: "rect",
            fill: "0F3460",
          },
        },
      ],
    },
    {
      layout: "obj",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "2.6cm",
            textBody: { text: "Title and Content Layout" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "2.6cm",
            y: "6.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: { text: "Content area" },
            fill: "E8E8E8",
          },
        },
      ],
    },
    {
      layout: "twoColTx",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "2.6cm",
            textBody: { text: "Two Column Layout" },
            fill: "70AD47",
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
            fill: "FFC000",
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
            fill: "ED7D31",
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
            fill: "2E75B6",
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
            fill: "E8E8E8",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
