import * as fs from "node:fs";

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
                      // Display size for the VML preview shape
                      width: 360,
                      height: 180,
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
fs.writeFileSync("OLE Document.docx", buffer);
