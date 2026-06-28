import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Phase 2 Demo",
  creator: "Demo",
  slides: [
    {
      background: { fill: "F2F2F2" },
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "1.3cm",
            width: "10.6cm",
            height: "2.6cm",
            textBody: { text: "With Outline" },
            geometry: "roundRect",
            fill: "FFFFFF",
            outline: { width: 25400, color: "4472C4", dashStyle: "dash" },
          },
        },
        {
          shape: {
            x: "13.2cm",
            y: "1.3cm",
            width: "10.6cm",
            height: "2.6cm",
            textBody: { text: "Gradient Fill" },
            fill: {
              type: "gradient",
              angle: 0,
              stops: [
                { position: 0, color: "4472C4" },
                { position: 100, color: "ED7D31" },
              ],
            },
          },
        },
      ],
    },
    {
      background: {
        fill: {
          type: "gradient",
          angle: 270,
          stops: [
            { position: 0, color: "1a1a2e" },
            { position: 100, color: "16213e" },
          ],
        },
      },
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "5.3cm",
            width: "7.9cm",
            height: "5.3cm",
            textBody: { text: "On Gradient BG" },
            fill: "FFFFFF",
            outline: { width: 12700, color: "FFC000" },
          },
        },
      ],
    },
    {
      background: {
        fill: "1B2A4A",
        effects: {
          outerShadow: {
            blur: 50800,
            distance: 38100,
            direction: 2700000,
            color: "000000",
            alpha: 50,
          },
        },
      },
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "13.2cm",
            height: "2.1cm",
            textBody: { text: "Background with Shadow Effect" },
            fill: "4472C4",
          },
        },
      ],
    },
    {
      background: {
        fill: {
          type: "gradient",
          angle: 90,
          stops: [
            { position: 0, color: "2E4057" },
            { position: 100, color: "048A81" },
          ],
        },
        shadeToTitle: true,
      },
      children: [
        {
          shape: {
            x: "2.6cm",
            y: "2.6cm",
            width: "13.2cm",
            height: "2.1cm",
            textBody: { text: "Background with shadeToTitle" },
            fill: "FFFFFF",
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
