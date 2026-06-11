// Exporting the document as a stream

import { createWriteStream } from "node:fs";
import { Readable } from "stream";
import type { ReadableStream as WebReadableStream } from "stream/web";

import { generateDocumentStream } from "@office-open/docx";

const stream = generateDocumentStream({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                text: "Foo Bar",
              },
              {
                bold: true,
                children: [{ tab: true }, "Github is the best"],
              },
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
Readable.fromWeb(stream as unknown as WebReadableStream).pipe(
  createWriteStream("My Document.docx"),
);
