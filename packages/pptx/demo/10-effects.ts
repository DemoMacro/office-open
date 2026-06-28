import * as fs from "fs";

import { generatePresentation } from "@office-open/pptx";
import type { PresentationOptions } from "@office-open/pptx";

const options: PresentationOptions = {
  title: "Effects Demo",
  creator: "Demo",
  slides: [
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.6cm",
            textBody: { text: "Shape Effects Demo" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "3.2cm",
            textBody: { text: "Outer Shadow" },
            fill: "ED7D31",
            effects: {
              outerShadow: {
                blur: 50800,
                distance: 38100,
                direction: 5400000,
                color: "000000",
                alpha: 50,
              },
            },
          },
        },
        {
          shape: {
            x: "7.9cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "3.2cm",
            textBody: { text: "Glow" },
            fill: "70AD47",
            effects: {
              glow: { radius: 152400, color: "92D050", alpha: 60 },
            },
          },
        },
        {
          shape: {
            x: "14.6cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "3.2cm",
            textBody: { text: "Reflection" },
            fill: "FFC000",
            effects: {
              reflection: {
                blurRadius: 6350,
                distance: 38100,
                direction: 5400000,
                startAlpha: 90,
                endAlpha: 0,
              },
            },
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "7.4cm",
            width: "5.3cm",
            height: "3.2cm",
            textBody: { text: "Inner Shadow" },
            fill: "5B9BD5",
            effects: {
              innerShadow: {
                blur: 40000,
                distance: 30000,
                direction: 5400000,
                color: "000000",
                alpha: 40,
              },
            },
          },
        },
        {
          shape: {
            x: "7.9cm",
            y: "7.4cm",
            width: "5.3cm",
            height: "3.2cm",
            textBody: { text: "Soft Edge" },
            fill: "BF8F00",
            effects: {
              softEdge: { radius: 50800 },
            },
          },
        },
        {
          shape: {
            x: "14.6cm",
            y: "7.4cm",
            width: "5.3cm",
            height: "3.2cm",
            textBody: { text: "Shadow + Glow" },
            fill: "7030A0",
            effects: {
              outerShadow: {
                blur: 40000,
                distance: 30000,
                direction: 2700000,
                color: "000000",
                alpha: 40,
              },
              glow: { radius: 101600, color: "B381E7", alpha: 35 },
            },
          },
        },
      ],
    },
    // Slide 2: 3D rotation
    {
      children: [
        {
          shape: {
            x: "1.3cm",
            y: "0.8cm",
            width: "13.2cm",
            height: "1.3cm",
            textBody: { text: "3D Rotation & Extrusion" },
            fill: "4472C4",
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "5.3cm",
            textBody: { text: "X=30 Y=0" },
            fill: "4472C4",
            effects: { rotation3D: { x: 30 } },
          },
        },
        {
          shape: {
            x: "7.9cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "5.3cm",
            textBody: { text: "X=0 Y=45" },
            fill: "ED7D31",
            effects: { rotation3D: { y: 45 } },
          },
        },
        {
          shape: {
            x: "14.6cm",
            y: "3.2cm",
            width: "5.3cm",
            height: "5.3cm",
            textBody: { text: "X=20 Y=30 Z=10" },
            fill: "70AD47",
            effects: { rotation3D: { x: 20, y: 30, z: 10, perspective: 500 } },
          },
        },
        {
          shape: {
            x: "1.3cm",
            y: "9.8cm",
            width: "5.3cm",
            height: "4.0cm",
            textBody: { text: "Extruded" },
            fill: "FFC000",
            effects: {
              rotation3D: { x: 25, y: 15 },
              extrusionH: 50000,
              material: "plastic",
            },
          },
        },
        {
          shape: {
            x: "7.9cm",
            y: "9.8cm",
            width: "5.3cm",
            height: "4.0cm",
            textBody: { text: "Bevel Top" },
            fill: "7030A0",
            effects: {
              rotation3D: { x: 20 },
              bevelTop: { width: 8, height: 8 },
              extrusionH: 25000,
              material: "metal",
            },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
