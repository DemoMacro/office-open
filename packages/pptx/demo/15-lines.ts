import * as fs from "fs";

import type { PresentationOptions } from "@file/file";
import { generate } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Lines & Connectors Demo",
  creator: "Demo",
  slides: [
    // Slide 1: Basic lines
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 50,
            textBody: { text: "Lines & Connectors" },
            fill: "4472C4",
          },
        },
        // Horizontal line
        {
          line: {
            x1: 50,
            y1: 120,
            x2: 800,
            y2: 120,
            outline: { color: "4472C4", width: 2 },
          },
        },
        // Vertical line
        {
          line: {
            x1: 200,
            y1: 150,
            x2: 200,
            y2: 450,
            outline: { color: "ED7D31", width: 2 },
          },
        },
        // Diagonal line
        {
          line: {
            x1: 250,
            y1: 150,
            x2: 750,
            y2: 450,
            outline: { color: "70AD47", width: 3 },
          },
        },
        // Reverse diagonal
        {
          line: {
            x1: 750,
            y1: 150,
            x2: 250,
            y2: 450,
            outline: { color: "FFC000", width: 2 },
          },
        },
      ],
    },
    // Slide 2: Connectors with arrowheads
    {
      children: [
        {
          shape: {
            x: 50,
            y: 30,
            width: 400,
            height: 50,
            textBody: { text: "Connectors with Arrowheads" },
            fill: "4472C4",
          },
        },
        // Left-to-right arrow
        {
          connector: {
            x1: 50,
            y1: 130,
            x2: 400,
            y2: 130,
            endArrowhead: "triangle",
            outline: { color: "4472C4", width: 2 },
          },
        },
        // Bidirectional arrow
        {
          connector: {
            x1: 50,
            y1: 200,
            x2: 400,
            y2: 200,
            beginArrowhead: "triangle",
            endArrowhead: "triangle",
            outline: { color: "ED7D31", width: 2 },
          },
        },
        // Stealth arrow
        {
          connector: {
            x1: 50,
            y1: 270,
            x2: 400,
            y2: 270,
            endArrowhead: "stealth",
            outline: { color: "70AD47", width: 2 },
          },
        },
        // Diamond end
        {
          connector: {
            x1: 50,
            y1: 340,
            x2: 400,
            y2: 340,
            endArrowhead: "diamond",
            outline: { color: "FFC000", width: 2 },
          },
        },
        // Oval end
        {
          connector: {
            x1: 50,
            y1: 410,
            x2: 400,
            y2: 410,
            endArrowhead: "oval",
            outline: { color: "7030A0", width: 2 },
          },
        },
        // Open arrow
        {
          connector: {
            x1: 500,
            y1: 130,
            x2: 800,
            y2: 130,
            endArrowhead: "open",
            outline: { color: "C00000", width: 2 },
          },
        },
        // Diagonal with stealth
        {
          connector: {
            x1: 500,
            y1: 200,
            x2: 800,
            y2: 350,
            endArrowhead: "stealth",
            outline: { color: "4472C4", width: 2 },
          },
        },
        // Large arrowhead
        {
          connector: {
            x1: 500,
            y1: 410,
            x2: 800,
            y2: 410,
            endArrowhead: "triangle",
            arrowheadWidth: "large",
            arrowheadLength: "large",
            outline: { color: "ED7D31", width: 2 },
          },
        },
      ],
    },
  ],
};

const buffer = await generate(options);
fs.writeFileSync("My Presentation.pptx", buffer);
