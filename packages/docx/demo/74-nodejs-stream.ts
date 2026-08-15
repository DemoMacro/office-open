// Exporting the document as a stream

import { createWriteStream, mkdirSync } from "node:fs";
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
mkdirSync(".temp", { recursive: true });
Readable.fromWeb(stream as unknown as WebReadableStream).pipe(
  createWriteStream(".temp/74-nodejs-stream.docx"),
);
