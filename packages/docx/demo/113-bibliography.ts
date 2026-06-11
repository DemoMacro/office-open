// Demo: Bibliography - citation management
import { writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  bibliography: {
    styleName: "APA",
    sources: [
      {
        type: "Book",
        title: "The Design of Everyday Things",
        author: "Norman, Donald",
        year: "2013",
        publisher: "Basic Books",
        city: "New York",
        edition: "Revised",
      },
      {
        type: "JournalArticle",
        title: "A Survey of Techniques for Building Secure Software",
        author: "Smith, J.; Doe, A.",
        year: "2026",
        month: "April",
        journal: "Journal of Software Engineering",
        volume: "42",
        issue: "3",
        pages: "100-120",
      },
      {
        type: "InternetSite",
        title: "TypeScript Documentation",
        author: "Microsoft",
        year: "2026",
        url: "https://www.typescriptlang.org/docs/",
      },
    ],
  },
  sections: [
    {
      children: [
        {
          paragraph: {
            children: [
              {
                text: "Bibliography Demo",
                bold: true,
                size: 32,
              },
            ],
            spacing: { after: 400 },
          },
        },

        // Bibliography SDT
        {
          paragraph: {
            children: [
              {
                bold: true,
                text: "Bibliography (APA style)",
                size: 28,
              },
            ],
            spacing: { after: 200 },
          },
        },

        {
          sdt: {
            properties: {
              bibliography: true,
              alias: "Bibliography",
              tag: "bibliography-sdt",
            },
            children: [
              {
                paragraph: {
                  children: ["Citations will be rendered here by Word."],
                },
              },
            ],
          },
        },

        { paragraph: { children: [""] } },

        {
          paragraph: {
            children: [
              {
                text: "Note: Open this document in Microsoft Word to see the bibliography rendered.",
                italics: true,
                color: "888888",
              },
            ],
          },
        },
      ],
    },
  ],
});
writeFileSync("My Document.docx", buffer);
