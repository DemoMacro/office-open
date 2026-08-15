// Simple example apply positional tabs to a document

import { mkdirSync, writeFileSync } from "node:fs";

import {
  generateDocument,
  PositionalTabAlignment,
  PositionalTabLeader,
  PositionalTabRelativeTo,
} from "@office-open/docx";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "Full name",
              {
                bold: true,
                children: [
                  {
                    positionalTab: {
                      alignment: PositionalTabAlignment.RIGHT,
                      relativeTo: PositionalTabRelativeTo.MARGIN,
                      leader: PositionalTabLeader.DOT,
                    },
                  },
                  "John Doe",
                ],
              },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "Hello World",
              {
                bold: true,
                children: [
                  {
                    positionalTab: {
                      alignment: PositionalTabAlignment.CENTER,
                      relativeTo: PositionalTabRelativeTo.INDENT,
                      leader: PositionalTabLeader.HYPHEN,
                    },
                  },
                  "Foo bar",
                ],
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
writeFileSync(".temp/84-positional-tabs.docx", buffer);
