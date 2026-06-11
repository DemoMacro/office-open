// Image blip effects (grayscale, luminance, duotone, tint, etc.)
import { readFileSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                bold: true,
                size: 32,
                text: "Image Blip Effects",
              },
            ],
            spacing: { after: 400 },
          },
        },

        // 1. Grayscale
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "1. Grayscale",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  blipEffects: { grayscale: true },
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Brightness + Contrast
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. Brightness (+30%) & Contrast (-20%)",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  blipEffects: {
                    luminance: { bright: 30, contrast: -20 },
                  },
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Duotone
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "3. Duotone (Dark Blue + Light Gray)",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  blipEffects: {
                    duotone: {
                      color1: { value: "002060" },
                      color2: { value: "D0CECE" },
                    },
                  },
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 4. Tint
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "4. Tint (Warm, 40%)",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  blipEffects: {
                    tint: { hue: 6000000, amount: 40 },
                  },
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 5. BiLevel (Black & White threshold)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "5. BiLevel (Threshold 50%)",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  blipEffects: {
                    biLevel: { threshold: 50 },
                  },
                  transformation: {
                    height: 150,
                    width: 150,
                  },
                  type: "jpg",
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
