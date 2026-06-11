// WPS (WordProcessing Shape) text boxes - modern DrawingML-based alternative to legacy VML text boxes (demo 94)
// Demonstrates: basic text box, styled fill/outline, rotation, floating positioning, and vertical alignment
import { readFileSync, writeFileSync } from "node:fs";

import {
  generateDocument,
  HorizontalPositionRelativeFrom,
  VerticalAnchor,
  VerticalPositionRelativeFrom,
} from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        // 1. Basic text box - minimal options
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: ["This is a basic WPS text box with just width and height."],
                  transformation: {
                    height: 80,
                    width: 400,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 2. Styled text box - solid fill and outline
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: [
                    {
                      children: [
                        {
                          bold: true,
                          color: "1F3864",
                          text: "Styled text box with a light blue background and dark blue border.",
                        },
                      ],
                    },
                  ],
                  outline: {
                    color: { value: "2E74B5" },
                    type: "solidFill",
                    width: 25_400,
                  },
                  fill: "D6E4F0",
                  transformation: {
                    height: 80,
                    width: 400,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 3. Rotated text box
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: ["This text box is rotated 15 degrees."],
                  outline: {
                    color: { value: "FF0000" },
                    type: "solidFill",
                    width: 12_700,
                  },
                  transformation: {
                    height: 60,
                    rotation: 15,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },
        { paragraph: { children: [""] } },

        // 4. Floating text box - positioned via offset
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  children: ["Floating text box positioned at a specific offset on the page."],
                  floating: {
                    horizontalPosition: {
                      offset: 3_500_000,
                      relative: HorizontalPositionRelativeFrom.PAGE,
                    },
                    verticalPosition: {
                      offset: 100_000,
                      relative: VerticalPositionRelativeFrom.PARAGRAPH,
                    },
                  },
                  outline: {
                    color: { value: "70AD47" },
                    type: "solidFill",
                    width: 19_050,
                  },
                  transformation: {
                    height: 70,
                    width: 350,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },
        { paragraph: { children: [""] } },
        { paragraph: { children: [""] } },

        // 5. Centered vertical alignment with custom margins
        {
          paragraph: {
            children: [
              {
                wpsShape: {
                  bodyProperties: {
                    margins: {
                      bottom: 72_000,
                      left: 144_000,
                      right: 144_000,
                      top: 72_000,
                    },
                    verticalAnchor: VerticalAnchor.CENTER,
                  },
                  children: ["Vertically centered text with custom margins."],
                  outline: {
                    color: { value: "BF8F00" },
                    type: "solidFill",
                    width: 19_050,
                  },
                  fill: "FFF2CC",
                  transformation: {
                    height: 120,
                    width: 400,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 6. Image fill (blipFill)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "6. Image Fill (blipFill)",
                size: 28,
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
                  children: ["Dog photo as shape fill"],
                  fill: {
                    type: "blip",
                    data: new Uint8Array(readFileSync("./demo/images/dog.png")),
                    imageType: "png",
                  },
                  transformation: {
                    height: 150,
                    width: 300,
                  },
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        // 7. Radial gradient fill (path)
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "7. Radial Gradient (path)",
                size: 28,
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
                          text: "Radial",
                        },
                      ],
                    },
                  ],
                  fill: {
                    type: "gradient",
                    path: "circle",
                    stops: [
                      { position: 0, color: "ED7D31" },
                      { position: 100, color: "C00000" },
                    ],
                  },
                  transformation: {
                    height: 100,
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
