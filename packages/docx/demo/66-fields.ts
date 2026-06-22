// Use fields to include dynamic text

import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  creator: "Me",
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              "This document is called ",
              { simpleField: { instruction: "FILENAME", cachedValue: "My Document.docx" } },
              ", was created on ",
              { simpleField: { instruction: String.raw`CREATEDATE  \@ "d MMMM yyyy"` } },
              " by ",
              { simpleField: { instruction: "AUTHOR" } },
            ],
          },
        },
        {
          paragraph: {
            children: [
              "The document has ",
              { simpleField: { instruction: "NUMWORDS", cachedValue: "34" } },
              " words and if you'd print it ",
              { bookmarkStart: { id: 0, name: "TimesPrinted" } },
              "42",
              { bookmarkEnd: { id: 0 } },
              " times two-sided, you would need ",
              { simpleField: { instruction: "=INT((TimesPrinted+1)/2)" } },
              " sheets of paper.",
            ],
          },
        },
      ],
      properties: {},
    },
  ],
});
writeFileSync("My Document.docx", buffer);
