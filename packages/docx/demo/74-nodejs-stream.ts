// Exporting the document as a stream

import * as fs from "fs";
import { Readable } from "stream";
import type { ReadableStream as WebReadableStream } from "stream/web";

import { Document, Packer, Paragraph, Tab, TextRun } from "@office-open/docx";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun("Hello World"),
            new TextRun({
              bold: true,
              text: "Foo Bar",
            }),
            new TextRun({
              bold: true,
              children: [new Tab(), "Github is the best"],
            }),
          ],
        }),
      ],
      properties: {},
    },
  ],
});

const stream = Packer.toStream(doc);
Readable.fromWeb(stream as unknown as WebReadableStream).pipe(
  fs.createWriteStream("My Document.docx"),
);
