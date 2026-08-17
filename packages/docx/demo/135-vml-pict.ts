import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";
import type { DocumentOptions } from "@office-open/docx";

// Minimal PNG (1x1 white pixel) standing in for a watermark image; real use
// supplies the source image bytes (WMF/EMF/PNG/JPEG all supported).
const watermarkBytes = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const doc: DocumentOptions = {
  sections: [
    {
      children: [
        { paragraph: { heading: "Heading1", children: ["VML pictures (w:pict)"] } },
        {
          paragraph: {
            children: [
              {
                // WordArt — v:shapetype preamble + v:shape with v:shadow and
                // v:textpath (the pre-DrawingML text-effect pipeline).
                pict: {
                  children: [
                    {
                      shapetype: {
                        id: "_x0000_t136",
                        coordsize: "21600,21600",
                        spt: 136,
                        adj: "10800",
                        path: "m@7,l@8,m@5,21600l@6,21600e",
                      },
                    },
                    {
                      shape: {
                        id: "_x0000_s1026",
                        type: "#_x0000_t136",
                        style: { width: "227.55pt", height: "22.4pt" },
                        filled: false,
                        shadow: { color: "#868686" },
                        textpath: {
                          style: { fontFamily: '"Arial Black"', fontSize: "16pt" },
                          trim: true,
                          string: "Office Open",
                        },
                      },
                    },
                  ],
                },
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                // Watermark-style picture — media entry + `{fileName}`
                // placeholder in the imagedata r:id (the compiler's media
                // bridge registers the bytes and rewrites the token).
                pict: {
                  children: [
                    {
                      shape: {
                        id: "_x0000_i1025",
                        type: "#_x0000_t75",
                        style: { width: "42.75pt", height: "57.75pt" },
                        imagedata: { relationshipId: "{image1.png}", grayscale: true },
                      },
                    },
                  ],
                  media: [{ fileName: "image1.png", data: watermarkBytes, type: "png" }],
                },
              },
              { text: " A picture watermark carried through w:pict." },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                // A plain VML rect (bullet-style block) and an empty pict —
                // both legal CT_Picture content.
                pict: {
                  children: [
                    {
                      rect: {
                        id: "_x0000_s1027",
                        style: { width: "120pt", height: "24pt" },
                        fillcolor: "#369",
                        stroked: false,
                      },
                    },
                  ],
                },
              },
              { pict: {} },
            ],
          },
        },
      ],
    },
  ],
};

const buffer = await generateDocument(doc);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/135-vml-pict.docx", buffer);
