// DrawingML Advanced Features - image cropping, effects, 3D, gradient fill, enhanced outline
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
                text: "DrawingML Advanced Features",
              },
            ],
            spacing: { after: 400 },
          },
        },

        // 1. Image cropping (srcRect)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "1. Image Cropping (srcRect)",
              },
            ],
            spacing: { after: 200 },
          },
        },
        {
          paragraph: {
            children: ["Original image:"],
            spacing: { after: 100 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
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
        {
          paragraph: {
            children: ["Cropped 25% from left and right:"],
            spacing: { after: 100 },
          },
        },
        {
          paragraph: {
            children: [
              {
                image: {
                  data: readFileSync("./demo/images/cat.jpg"),
                  srcRect: { left: 25000, right: 25000 },
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

        // 2. Shadow effects
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "2. Shadow Effects",
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
                      children: ["Outer Shadow"],
                    },
                  ],
                  effects: {
                    outerShadow: {
                      blurRadius: 50800,
                      color: { value: "000000" },
                      distance: 38100,
                      direction: 5400000,
                    },
                  },
                  outline: {
                    color: { value: "4472C4" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  fill: "D9E2F3",
                  transformation: {
                    height: 80,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Glow effect
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "3. Glow Effect",
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
                          text: "Glowing Box",
                        },
                      ],
                    },
                  ],
                  effects: {
                    glow: { color: { value: "00B0F0" }, radius: 76200 },
                  },
                  outline: {
                    color: { value: "0070C0" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  fill: "0070C0",
                  transformation: {
                    height: 80,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 4. Reflection effect
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "4. Reflection Effect",
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
                      children: ["Reflection"],
                    },
                  ],
                  effects: {
                    reflection: {
                      blurRadius: 6350,
                      distance: 38100,
                      fadeDirection: 5400000,
                      startAlpha: 40000,
                    },
                  },
                  outline: {
                    color: { value: "ED7D31" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  fill: "F4B183",
                  transformation: {
                    height: 80,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 5. Gradient fill (linear)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "5. Gradient Fill (Linear)",
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
                          text: "Linear Gradient",
                        },
                      ],
                    },
                  ],
                  fill: {
                    type: "gradient",
                    options: {
                      shade: { angle: 5400000 },
                      stops: [
                        { color: { value: "002060" }, position: 0 },
                        { color: { value: "0070C0" }, position: 50000 },
                        { color: { value: "00B0F0" }, position: 100000 },
                      ],
                    },
                  },
                  outline: {
                    color: { value: "002060" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  transformation: {
                    height: 80,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 6. Gradient fill (radial/path)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "6. Gradient Fill (Radial)",
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
                          text: "Radial Gradient",
                        },
                      ],
                    },
                  ],
                  fill: {
                    type: "gradient",
                    options: {
                      shade: { path: "circle" },
                      stops: [
                        { color: { value: "FFFFFF" }, position: 0 },
                        { color: { value: "4472C4" }, position: 100000 },
                      ],
                    },
                  },
                  outline: {
                    color: { value: "2F5496" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  transformation: {
                    height: 80,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 7. Soft edge effect
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "7. Soft Edge Effect",
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
                      children: ["Soft Edge"],
                    },
                  ],
                  effects: {
                    softEdge: 50800,
                  },
                  fill: "70AD47",
                  transformation: {
                    height: 80,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 8. 3D shape with bevel and material
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "8. 3D Shape (Bevel + Material)",
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
                      children: ["3D Effect"],
                    },
                  ],
                  shape3d: {
                    bevelT: { prst: "circle", w: 76200, h: 76200 },
                    extrusionColor: { value: "2F5496" },
                    prstMaterial: "plastic",
                    z: 76200,
                  },
                  fill: "4472C4",
                  transformation: {
                    height: 120,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 9. Combined effects
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "9. Combined Effects (Shadow + Reflection + Gradient)",
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
                          text: "All Together!",
                        },
                      ],
                    },
                  ],
                  effects: {
                    outerShadow: {
                      blurRadius: 40000,
                      color: { value: "000000" },
                      distance: 30000,
                      direction: 5400000,
                    },
                    reflection: {
                      blurRadius: 4000,
                      distance: 20000,
                      fadeDirection: 5400000,
                      startAlpha: 30000,
                    },
                  },
                  fill: {
                    type: "gradient",
                    options: {
                      shade: { angle: 2700000 },
                      stops: [
                        { color: { value: "1F4E79" }, position: 0 },
                        { color: { value: "2E75B6" }, position: 50000 },
                        { color: { value: "9DC3E6" }, position: 100000 },
                      ],
                    },
                  },
                  outline: {
                    color: { value: "1F4E79" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  transformation: {
                    height: 120,
                    width: 400,
                  },
                },
              },
            ],
          },
        },

        // 10. Enhanced outline (dash style + line join)
        { paragraph: { children: [""] } },
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "10. Enhanced Outline (Dash + Join)",
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
                      children: ["Dashed Outline"],
                    },
                  ],
                  outline: {
                    color: { value: "C00000" },
                    dash: "dash",
                    join: "round",
                    type: "solidFill",
                    width: 25_400,
                  },
                  fill: "FFF2CC",
                  transformation: {
                    height: 80,
                    width: 300,
                  },
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
