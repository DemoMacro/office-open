import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";
import type { DocumentOptions } from "@office-open/docx";

// Mock OLE compound file (real use: supply the actual OLE container from the
// source application — Excel.Sheet.12, PowerPoint.Show, Equation, etc.).
// Header bytes D0CF11E0... are the OLE2 magic; the rest is zero-padded here.
const oleBytes = new Uint8Array([
  0xd0,
  0xcf,
  0x11,
  0xe0,
  0xa1,
  0xb1,
  0x1a,
  0xe1,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x3e,
  0x00,
  0x03,
  0x00,
  0xfe,
  0xff,
  0x09,
  0x00,
  0x06,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x00,
  0x01,
  0x00,
  0x00,
  0x00,
  ...Array.from({ length: 512 }, () => 0),
]);

// Minimal 1x1 transparent PNG for the v:imagedata preview icon.
const iconPng = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const doc: DocumentOptions = {
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [{ text: "OLE Object Embedding", bold: true }],
          },
        },
        {
          paragraph: {
            children: [
              {
                children: [
                  {
                    object: {
                      // Original OLE size in twips (w:dxaOrig / w:dyaOrig)
                      dxaOrig: 5400,
                      dyaOrig: 2700,
                      // Word's OLE shapetype preamble (_x0000_t75), round-tripped
                      // verbatim ahead of the preview shape
                      shapetype: {
                        id: "_x0000_t75",
                        coordsize: "21600,21600",
                        spt: 75,
                        preferrelative: true,
                        path: "m@4@5l@4@11@9@11@9@5xe",
                        filled: false,
                        stroked: false,
                        stroke: { joinstyle: "miter" },
                        formulas: {
                          equations: [
                            "if lineDrawn pixelLineWidth 0",
                            "sum @0 1 0",
                            "sum 0 0 @1",
                            "prod @2 1 2",
                            "prod @3 21600 pixelWidth",
                            "prod @3 21600 pixelHeight",
                            "sum @0 0 1",
                            "prod @6 1 2",
                            "prod @7 21600 pixelWidth",
                            "sum @8 21600 0",
                            "prod @7 21600 pixelHeight",
                            "sum @10 21600 0",
                          ],
                        },
                        pathElement: {
                          extrusionok: false,
                          gradientshapeok: true,
                          connecttype: "rect",
                        },
                        lock: { ext: "edit", aspectratio: true },
                      },
                      // Display size for the VML preview shape
                      width: "9.5cm",
                      height: "4.8cm",
                      // Preview icon rendered via v:imagedata
                      iconImage: { data: iconPng, type: "png", title: "Excel Sheet" },
                      // Embedded OLE object → word/embeddings/oleObject1.bin
                      embed: {
                        data: oleBytes,
                        progId: "Excel.Sheet.12",
                        drawAspect: "content",
                        shapeId: "_x0000_i1026",
                      },
                    },
                  },
                ],
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              {
                text: "Use `link` (with updateMode) for linked objects, `control` (rId) for ActiveX, or `movie` (rId) for legacy media.",
              },
            ],
          },
        },
      ],
    },
  ],
};

const buffer = await generateDocument(doc);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/134-ole-object.docx", buffer);
