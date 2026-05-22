import * as fs from "fs";

import { Presentation, Shape, Packer } from "@office-open/pptx";

const pres = new Presentation({
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
        new Shape({
          x: 0,
          y: 695,
          width: 1280,
          height: 25,
          geometry: "rect",
          fill: "4472C4",
        }),
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
        new Shape({
          x: 0,
          y: 695,
          width: 1280,
          height: 25,
          geometry: "rect",
          fill: "FFC000",
        }),
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
        new Shape({
          x: 100,
          y: 100,
          width: 600,
          height: 200,
          text: "Slide on Light Master",
          fill: "4472C4",
        }),
      ],
    },
    {
      master: "dark",
      layout: "title",
      children: [
        new Shape({
          x: 100,
          y: 100,
          width: 600,
          height: 200,
          text: "Slide on Dark Master",
          fill: "FFC000",
        }),
      ],
    },
    {
      master: "light",
      layout: "blank",
      children: [
        new Shape({
          x: 100,
          y: 200,
          width: 600,
          height: 300,
          text: "Blank Layout on Light",
          fill: "E8E8E8",
        }),
      ],
    },
    {
      master: "dark",
      layout: "obj",
      children: [
        new Shape({
          x: 100,
          y: 200,
          width: 600,
          height: 300,
          text: "Content Layout on Dark",
          fill: "2C3E50",
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
fs.writeFileSync("My Presentation.pptx", buffer);
