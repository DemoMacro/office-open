import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Multi-Master Demo",
  creator: "Demo",
  masters: [
    {
      name: "light",
      theme: {
        colors: {
          accent1: "4472C4",
          accent2: "ED7D31",
        },
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
    },
    {
      name: "dark",
      theme: {
        colors: {
          accent1: "FFC000",
          accent2: "70AD47",
        },
      },
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
            fill: "FFC000",
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
    },
  ],
  slides: [
    {
      master: "light",
      layout: "title",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: { text: "Slide on Light Master" },
            fill: "4472C4",
          },
        },
      ],
    },
    {
      master: "dark",
      layout: "title",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "15.9cm",
            height: "5.3cm",
            textBody: { text: "Slide on Dark Master" },
            fill: "FFC000",
          },
        },
      ],
    },
    {
      master: "light",
      layout: "blank",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "5.3cm",
            width: "15.9cm",
            height: "7.9cm",
            textBody: { text: "Blank Layout on Light" },
            fill: "E8E8E8",
          },
        },
      ],
    },
    {
      master: "dark",
      layout: "obj",
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "5.3cm",
            width: "15.9cm",
            height: "7.9cm",
            textBody: { text: "Content Layout on Dark" },
            fill: "2C3E50",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
