// Demo: Bibliography - citation management
import { mkdirSync, writeFileSync } from "node:fs";

import { generateDocument } from "@office-open/docx";

const buffer = await generateDocument({
  bibliography: {
    styleName: "APA",
    sources: [
      {
        sourceType: "Book",
        title: "The Design of Everyday Things",
        author: { authors: [{ last: "Norman", first: "Donald" }] },
        year: "2013",
        publisher: "Basic Books",
        city: "New York",
        edition: "Revised",
      },
      {
        sourceType: "JournalArticle",
        title: "A Survey of Techniques for Building Secure Software",
        author: {
          authors: [
            { last: "Smith", first: "J." },
            { last: "Doe", first: "A." },
          ],
          editors: [{ last: "Lee", first: "Kai" }],
        },
        year: "2026",
        month: "April",
        journal: "Journal of Software Engineering",
        volume: "42",
        issue: "3",
        pages: "100-120",
      },
      {
        sourceType: "InternetSite",
        title: "TypeScript Documentation",
        author: { authors: [{ corporate: "Microsoft" }] },
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
                size: 16,
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
                size: 14,
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
                italic: true,
                color: "888888",
              },
            ],
          },
        },
      ],
    },
  ],
});
mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/113-bibliography.docx", buffer);
