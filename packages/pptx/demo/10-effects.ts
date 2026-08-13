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
                blurRadius: 50800,
                distance: 38100,
                direction: 5400000,
                color: { value: "000000", transforms: { alpha: 50 } },
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
              glow: { radius: 152400, color: { value: "92D050", transforms: { alpha: 60 } } },
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
                blurRadius: 40000,
                distance: 30000,
                direction: 5400000,
                color: { value: "000000", transforms: { alpha: 40 } },
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
              softEdge: 50800,
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
                blurRadius: 40000,
                distance: 30000,
                direction: 2700000,
                color: { value: "000000", transforms: { alpha: 40 } },
              },
              glow: { radius: 101600, color: { value: "B381E7", transforms: { alpha: 35 } } },
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
            scene3d: {
              camera: { preset: "orthographicFront", rotation: { lat: 1800000, lon: 0, rev: 0 } },
              lightRig: { rig: "threePt", direction: "t" },
            },
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
            scene3d: {
              camera: { preset: "orthographicFront", rotation: { lat: 0, lon: 2700000, rev: 0 } },
              lightRig: { rig: "threePt", direction: "t" },
            },
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
            scene3d: {
              camera: {
                preset: "legacyPerspectiveFront",
                rotation: { lat: 1200000, lon: 1800000, rev: 600000 },
                fov: 500,
              },
              lightRig: { rig: "threePt", direction: "t" },
            },
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
            scene3d: {
              camera: {
                preset: "orthographicFront",
                rotation: { lat: 1500000, lon: 900000, rev: 0 },
              },
              lightRig: { rig: "threePt", direction: "t" },
            },
            shape3d: { extrusionH: 50000, prstMaterial: "plastic" },
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
            scene3d: {
              camera: { preset: "orthographicFront", rotation: { lat: 1200000, lon: 0, rev: 0 } },
              lightRig: { rig: "threePt", direction: "t" },
            },
            shape3d: { bevelT: { w: 8, h: 8 }, extrusionH: 25000, prstMaterial: "metal" },
          },
        },
      ],
    },
  ],
};

const buffer = await generatePresentation(options);
fs.writeFileSync("My Presentation.pptx", buffer);
