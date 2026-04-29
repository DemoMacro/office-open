// Demo: Bibliography - citation management
import * as fs from "fs";

import { Document, Packer, Paragraph, TextRun, StructuredDocumentTagBlock } from "docx-plus";

const doc = new Document({
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
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Bibliography Demo",
                            bold: true,
                            size: 32,
                        }),
                    ],
                    spacing: { after: 400 },
                }),

                // Bibliography SDT
                new Paragraph({
                    children: [
                        new TextRun({
                            bold: true,
                            text: "Bibliography (APA style)",
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                }),

                new StructuredDocumentTagBlock({
                    properties: {
                        bibliography: true,
                        alias: "Bibliography",
                        tag: "bibliography-sdt",
                    },
                    children: [
                        new Paragraph({
                            children: [new TextRun("Citations will be rendered here by Word.")],
                        }),
                    ],
                }),

                new Paragraph({ children: [new TextRun("")] }),

                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Note: Open this document in Microsoft Word to see the bibliography rendered.",
                            italics: true,
                            color: "888888",
                        }),
                    ],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
