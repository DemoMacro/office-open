import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

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
            x: 0,
            y: 695,
            width: 1280,
            height: 25,
            geometry: "rect",
            fill: "4472C4",
          },
        },
      ],
      placeholders: {
        slideNumber: {
          x: 904,
          y: 667,
          width: 288,
          height: 38,
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
            x: 0,
            y: 695,
            width: 1280,
            height: 25,
            geometry: "rect",
            fill: "FFC000",
          },
        },
      ],
      placeholders: {
        slideNumber: {
          x: 904,
          y: 667,
          width: 288,
          height: 38,
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
            x: 100,
            y: 100,
            width: 600,
            height: 200,
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
            x: 100,
            y: 100,
            width: 600,
            height: 200,
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
            x: 100,
            y: 200,
            width: 600,
            height: 300,
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
            x: 100,
            y: 200,
            width: 600,
            height: 300,
            textBody: { text: "Content Layout on Dark" },
            fill: "2C3E50",
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
