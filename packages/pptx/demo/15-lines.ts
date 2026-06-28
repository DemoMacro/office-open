import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Lines & Connectors Demo",
  creator: "Demo",
  slides: [
    // Slide 1: Basic lines
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.3cm",
            textBody: { text: "Lines & Connectors" },
            fill: "4472C4",
          },
        },
        // Horizontal line
        {
          line: {
            x1: "1.3cm",
            y1: "3.2cm",
            x2: "21.2cm",
            y2: "3.2cm",
            outline: { color: "4472C4", width: "2pt" },
          },
        },
        // Vertical line
        {
          line: {
            x1: "5.3cm",
            y1: "4.0cm",
            x2: "5.3cm",
            y2: "11.9cm",
            outline: { color: "ED7D31", width: "2pt" },
          },
        },
        // Diagonal line
        {
          line: {
            x1: "6.6cm",
            y1: "4.0cm",
            x2: "19.8cm",
            y2: "11.9cm",
            outline: { color: "70AD47", width: "3pt" },
          },
        },
        // Reverse diagonal
        {
          line: {
            x1: "19.8cm",
            y1: "4.0cm",
            x2: "6.6cm",
            y2: "11.9cm",
            outline: { color: "FFC000", width: "2pt" },
          },
        },
      ],
    },
    // Slide 2: Connectors with arrowheads
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "10.6cm",
            height: "1.3cm",
            textBody: { text: "Connectors with Arrowheads" },
            fill: "4472C4",
          },
        },
        // Left-to-right arrow
        {
          connector: {
            x1: "1.3cm",
            y1: "3.4cm",
            x2: "10.6cm",
            y2: "3.4cm",
            endArrowhead: "triangle",
            outline: { color: "4472C4", width: "2pt" },
          },
        },
        // Bidirectional arrow
        {
          connector: {
            x1: "1.3cm",
            y1: "5.3cm",
            x2: "10.6cm",
            y2: "5.3cm",
            beginArrowhead: "triangle",
            endArrowhead: "triangle",
            outline: { color: "ED7D31", width: "2pt" },
          },
        },
        // Stealth arrow
        {
          connector: {
            x1: "1.3cm",
            y1: "7.1cm",
            x2: "10.6cm",
            y2: "7.1cm",
            endArrowhead: "stealth",
            outline: { color: "70AD47", width: "2pt" },
          },
        },
        // Diamond end
        {
          connector: {
            x1: "1.3cm",
            y1: "9.0cm",
            x2: "10.6cm",
            y2: "9.0cm",
            endArrowhead: "diamond",
            outline: { color: "FFC000", width: "2pt" },
          },
        },
        // Oval end
        {
          connector: {
            x1: "1.3cm",
            y1: "10.8cm",
            x2: "10.6cm",
            y2: "10.8cm",
            endArrowhead: "oval",
            outline: { color: "7030A0", width: "2pt" },
          },
        },
        // Open arrow
        {
          connector: {
            x1: "13.2cm",
            y1: "3.4cm",
            x2: "21.2cm",
            y2: "3.4cm",
            endArrowhead: "open",
            outline: { color: "C00000", width: "2pt" },
          },
        },
        // Diagonal with stealth
        {
          connector: {
            x1: "13.2cm",
            y1: "5.3cm",
            x2: "21.2cm",
            y2: "9.3cm",
            endArrowhead: "stealth",
            outline: { color: "4472C4", width: "2pt" },
          },
        },
        // Large arrowhead
        {
          connector: {
            x1: "13.2cm",
            y1: "10.8cm",
            x2: "21.2cm",
            y2: "10.8cm",
            endArrowhead: "triangle",
            arrowheadWidth: "large",
            arrowheadLength: "large",
            outline: { color: "ED7D31", width: "2pt" },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
