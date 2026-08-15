// Sequential Captions

import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Hello World 1->",
              { seqIdentifier: "Caption" },
              " text after sequencial caption 2->",
              { seqIdentifier: "Caption" },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Hello World 1->",
              { seqIdentifier: "Label" },
              " text after sequencial caption 2->",
              { seqIdentifier: "Label" },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Hello World 1->",
              { seqIdentifier: "Another" },
              " text after sequencial caption 3->",
              { seqIdentifier: "Label" },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Hello World 2->",
              { seqIdentifier: "Another" },
              " text after sequencial caption 4->",
              { seqIdentifier: "Label" },
            ],
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/33-sequential-captions.docx", buffer);
