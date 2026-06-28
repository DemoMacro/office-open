// Demo: 3D scene (CT_Scene3D) - DrawingML
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "3D Scene Demo", bold: true, size: 16 }],
            spacing: { after: 400 },
          },
        },

        // 1. Shape with 3D scene (camera + light rig)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "1. Isometric Camera + Three-Point Light",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: [{ bold: true, color: "FFFFFF", text: "3D" }],
                    },
                  ],
                  scene3d: {
                    camera: { preset: "isometricTopUp" },
                    lightRig: { direction: "t", rig: "threePt" },
                  },
                  shape3d: {
                    bevelT: { prst: "circle", w: 76200, h: 76200 },
                    extrusionColor: { value: "2F5496" },
                    prstMaterial: "plastic",
                    z: 127000,
                  },
                  fill: "4472C4",
                  transformation: { height: "4.0cm", width: "7.9cm" },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Perspective camera
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. Perspective Camera + Morning Light",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: [
                        {
                          bold: true,
                          color: "FFFFFF",
                          text: "Perspective",
                        },
                      ],
                    },
                  ],
                  scene3d: {
                    camera: {
                      fov: 600000,
                      preset: "perspectiveFront",
                      zoom: "120000",
                    },
                    lightRig: { direction: "t", rig: "morning" },
                  },
                  shape3d: {
                    bevelB: { prst: "circle", w: 50800, h: 25400 },
                    bevelT: { prst: "relaxedInset", w: 50800, h: 50800 },
                    extrusionColor: { value: "7030A0" },
                    prstMaterial: "matte",
                    z: 76200,
                  },
                  fill: "7030A0",
                  transformation: { height: "4.0cm", width: "7.9cm" },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Camera with rotation
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "3. Rotated Camera",
                size: 14,
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      alignment: "center",
                      children: [{ bold: true, text: "Rotated View" }],
                    },
                  ],
                  scene3d: {
                    camera: {
                      preset: "orthographicFront",
                      rotation: {
                        lat: 0,
                        lon: 2700000,
                        rev: 450000,
                      },
                    },
                    lightRig: { direction: "t", rig: "balanced" },
                  },
                  shape3d: {
                    bevelT: { prst: "angle", w: 76200, h: 38100 },
                    contourColor: { value: "000000" },
                    contourW: 12700,
                    extrusionColor: { value: "C00000" },
                    prstMaterial: "warmMatte",
                    z: 50800,
                  },
                  fill: "ED7D31",
                  transformation: { height: "4.0cm", width: "7.9cm" },
                },
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
