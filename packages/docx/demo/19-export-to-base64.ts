// Export to base64 string - Useful in a browser environment.

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument(
  {
    sections: [
      {
        children: [
          {
            paragraph: {
              children: [
                "Hello World",
                {
                  bold: true,
                  text: "Foo",
                },
                {
                  bold: true,
                  children: [{ tab: true }, "Bar"],
                },
              ],
            },
          },
        ],
      },
    ],
  },
  { type: "base64" },
);
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/19-export-to-base64.docx", Buffer.from(buffer, "base64"));
