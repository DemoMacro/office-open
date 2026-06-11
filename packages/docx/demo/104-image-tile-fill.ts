// Image tile fill mode (repeating image pattern)
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
                text: "Image Tile Fill Demo",
              },
            ],
            spacing: { after: 400 },
          },
        },

        // 1. Default stretch (for comparison)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "1. Default Stretch Fill (no tile)",
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
                  transformation: {
                    height: 100,
                    width: 300,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Tile with 50% scale
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. Tile Fill (50% scale)",
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
                  tile: { sx: 50, sy: 50 },
                  transformation: {
                    height: 100,
                    width: 300,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Tile with flip
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "3. Tile Fill (50% scale, XY flip)",
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
                  tile: { flip: "xy", sx: 50, sy: 50 },
                  transformation: {
                    height: 100,
                    width: 300,
                  },
                  type: "jpg",
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 4. Tile with alignment
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "4. Tile Fill (50% scale, center aligned)",
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
                  tile: { align: "center", sx: 50, sy: 50 },
                  transformation: {
                    height: 100,
                    width: 300,
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
